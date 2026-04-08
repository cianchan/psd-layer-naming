"""
seg_product.py — Shopee 扣图 API 集成模块
识别 PSD 输出文件中哪个 scenebg 图层是商品主体，
并将该图层重命名为 "product"。

依赖:
    pip install psd-tools pillow requests

CLI 用法:
    python seg_product.py <psd路径> <level1品类> <level3品类>
    python seg_product.py output/kettle.psd "Electronics" "Electric Kettles"
"""

import sys
import os
import json
import base64
import time
import requests
from io import BytesIO
from pathlib import Path

from PIL import Image
from psd_tools import PSDImage

# ── Shopee API 配置 ─────────────────────────────────────────────────────────

_HEADERS_LIVEISH = {
    "Content-Type": "application/json",
    "x-sp-sdu": "ai_engine_platform.mmuplt.controller.global.liveish.master.default",
    "x-sp-servicekey": "f0dd2d544097d2a938595c1d78949bd3",
    "x-sp-timeout": "60000",
    "x-sp-processid": "process_1",
}
_HEADERS_SEGLABEL = {**_HEADERS_LIVEISH, "x-sp-timeout": "30000", "x-sp-processid": "CID=global"}


def call_service(service_name, request_data, env='liveish', max_retries=3, log_fn=None):
    """调用 Shopee AI 服务（需要 Shopee VPN）。所有错误通过 log_fn 输出。"""
    def _log(msg):
        (log_fn or print)(msg)

    url = (
        f"https://http-gateway.spex.shopee.sg/sprpc/"
        f"ai_engine_platform.mmuplt.{service_name}.algo"
    )
    headers = _HEADERS_SEGLABEL if service_name == 'seglabel2' else _HEADERS_LIVEISH

    for attempt in range(max_retries):
        try:
            resp = requests.post(url, data=json.dumps(request_data), headers=headers, timeout=60)
            if resp.status_code != 200:
                _log(f"  [API] {service_name} HTTP {resp.status_code}: {resp.text[:300]}")
                if attempt < max_retries - 1:
                    time.sleep(1)
                    continue
                return None
            rj = resp.json()
            if 'task_result' not in rj:
                _log(f"  [API] {service_name} 返回结构异常（无 task_result）: {str(rj)[:300]}")
                if attempt < max_retries - 1:
                    time.sleep(1)
                    continue
                return None
            return rj['task_result']
        except requests.exceptions.ConnectionError as e:
            _log(f"  [API] {service_name} 连接失败（请检查 Shopee VPN）: {e}")
            if attempt < max_retries - 1:
                time.sleep(1)
                continue
            return None
        except Exception as e:
            _log(f"  [API] {service_name} 异常: {e}")
            if attempt < max_retries - 1:
                time.sleep(1)
                continue
            return None
    return None


# ── 图像工具函数 ─────────────────────────────────────────────────────────────

def pil_to_base64(pil_image, max_size=1024):
    """
    将 PIL 图片转为 base64 字符串。
    如果图片超过 max_size，等比缩放以控制 payload 大小。
    """
    img = pil_image.convert("RGBA")
    w, h = img.size
    if w > max_size or h > max_size:
        ratio = min(max_size / w, max_size / h)
        img = img.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)
    buf = BytesIO()
    img.save(buf, format='PNG', optimize=True)
    return base64.b64encode(buf.getvalue()).decode('utf-8')


def piseg_pil(pil_image, cates, env='liveish', log_fn=None):
    """
    对一张 PIL 图片调用分割服务（pisegv2）。
    传入 全画布尺寸 的图像（图层粘贴在正确位置的透明背景上）。
    返回 (seg_ims_base64, bboxes, seg_labels)，失败返回 None。
    """
    def _log(msg):
        (log_fn or print)(msg)

    b64 = pil_to_base64(pil_image)
    extra = json.dumps({'is_upload': False, 'cates': cates, 'bboxes': []})
    req = {
        "biz_type": "mmu_test",
        "region": "sg2",
        "task": {
            "image_list": [{"image_data": b64, "extra_info": extra}]
        }
    }
    res = call_service('pisegv2', req, env=env, log_fn=log_fn)
    if res is None:
        return None
    try:
        ei = json.loads(res['extra_info'])
        imgs   = ei.get('object_images', [])
        bboxes = ei.get('object_bboxes', [])
        labels = ei.get('object_labels', [])
        if not imgs:
            _log(f"  [API] pisegv2 返回 0 个分割块（图层中未检测到独立对象）")
            return None
        return imgs, bboxes, labels
    except Exception as e:
        _log(f"  [API] piseg 结果解析失败: {e}  原始: {str(res)[:200]}")
        return None


def tagging_base64(seg_b64, level3_cat, bbox, seg_label, env='liveish', log_fn=None):
    """对分割块（base64）调用打标签服务。返回标签列表或 None。"""
    extra = json.dumps({
        'bbox': bbox,
        'source_url': '',
        'level3_global_be_category': level3_cat,
        'seg_img_tags': seg_label
    })
    req = {
        "biz_type": "mmu_algo_test",
        "region": "sg",
        "task": {
            "image_list": [{"image_data": seg_b64, "extra_info": extra}],
            "extra_info": json.dumps({'type': 'pimg'})
        }
    }
    res = call_service('seglabel2', req, env=env, log_fn=log_fn)
    if res is None:
        return None
    try:
        lst = json.loads(res['extra_info'])[0]
        return lst if lst else None
    except Exception as e:
        (log_fn or print)(f"  [API] tagging 结果解析失败: {e}")
        return None


def rank_segments(label_list, bbox_list, level1_cat, level3_cat, env='liveish', log_fn=None):
    """对所有分割块一起打分排名。返回排名结果字典或 None。"""
    text_list = [
        {
            'text': json.dumps(label),
            'extra_info': json.dumps({
                'box': bbox,
                'ori_img_height': 1024,
                'ori_img_width': 1024,
                'l1': level1_cat,
                'l3': level3_cat
            })
        }
        for label, bbox in zip(label_list, bbox_list)
    ]
    req = {
        "biz_type": "mmu_algo_test",
        "region": "sg2",
        "task": {
            "text_list": text_list,
            "extra_info": json.dumps({"type": 'pimg'})
        }
    }
    return call_service('pfilter2', req, env=env, log_fn=log_fn)


# ── PSD 图层处理 ─────────────────────────────────────────────────────────────

def export_scenebg_layers(psd_path, log_fn=None):
    """
    打开处理后的 PSD，将所有名为 'scenebg' 的图层导出为 PIL image。
    关键：每个图层粘贴在与 PSD 等大的全画布上（透明背景），
    保留图层在画面中的位置上下文，让 API 能正确判断是否是主体。

    返回 [(layer_index, layer_name, full_canvas_pil), ...]
    """
    def _log(msg):
        (log_fn or print)(msg)

    psd = PSDImage.open(psd_path)
    doc_w, doc_h = psd.width, psd.height
    _log(f"  [PSD] 文档尺寸: {doc_w}×{doc_h}")

    results = []
    for i, layer in enumerate(psd):
        if layer.name != 'scenebg':
            continue
        try:
            layer_img = layer.composite()
            if layer_img is None:
                _log(f"  [PSD] 图层 {i} composite() 返回 None，跳过")
                continue

            layer_img_rgba = layer_img.convert('RGBA')   # 图层自身（用于分析）

            # 粘贴到全画布，保留位置信息（用于 API 调用）
            canvas = Image.new('RGBA', (doc_w, doc_h), (0, 0, 0, 0))
            bbox = layer.bbox   # BBox(left, top, right, bottom)
            canvas.paste(layer_img_rgba, (bbox.left, bbox.top))

            # 返回 4-tuple：(index, name, 图层自身RGBA, 全画布RGBA)
            results.append((i, layer.name, layer_img_rgba, canvas))
            _log(f"  [PSD] 图层 {i}: {layer.name}  位置=({bbox.left},{bbox.top})  "
                 f"图层尺寸={layer_img_rgba.size}  画布={canvas.size}")
        except Exception as e:
            _log(f"  [PSD] 图层 {i} ({layer.name}) 导出失败: {e}")

    return results


# ── Case 1：透明度 / 重复内容 / 尺寸排除 三路检测（无需 API） ───────────────

def detect_existing_cutout(layers, log_fn=None):
    """
    检测 scenebg 图层中是否已有商品主体图层。

    layers 格式（4-tuple）：(layer_idx, layer_name, layer_img_rgba, canvas_img)
      - layer_img_rgba：图层自身裁剪区域（用于分析，不含画布填充边框）
      - canvas_img：全画布版本（保留给 API 调用，此函数不使用）

    路径 A — 透明扣图检测（分析图层自身的 alpha）：
      图层本身有透明背景（已扣图），商品主体清晰可见。
      条件：图层自身 transparent_ratio 0.10-0.88，图层面积≥画布3%，宽高比 0.15-6.5。
      → 多候选取面积最大的。

    路径 B — 重复内容检测（分析图层自身内容相似度）：
      多个 scenebg 含相似内容（扣图层 + 场景图层同时包含该商品）。
      两两比较图层不透明区域 RGB 直方图，相似度≥0.80 时，面积较小者 = 商品主体。
      → 取面积最小的候选。

    路径 C — 尺寸排除法（处理无透明通道的商品图）：
      适用于商品在白色/纯色背景上（transparent_ratio≈0），无法用 alpha 检测的情况。
      在「实质内容」图层（面积>8%画布，几乎不透明）中：
        - 只有 1 个 → 直接是商品（如 MX3 PSD 只有一个产品图）
        - 多个 → 排除面积最大的（= 填满画布的场景背景），选次大的（= 商品主体）

    优先级：B > C > A。
    返回 (layer_idx, layer_name) 或 None。
    """
    def _log(msg):
        (log_fn or print)(msg)

    if not layers:
        return None

    # canvas_area 用于比较图层相对大小，layer_img_rgba 用于分析内容
    canvas_area = layers[0][3].width * layers[0][3].height  # 第4个元素是 canvas_img

    def cosine(h1, h2):
        if not h1 or not h2:
            return 0.0
        dot  = sum(a * b for a, b in zip(h1, h2))
        mag1 = sum(a * a for a in h1) ** 0.5
        mag2 = sum(b * b for b in h2) ** 0.5
        return dot / (mag1 * mag2) if (mag1 and mag2) else 0.0

    # ── 预计算每层指标（全部基于 layer_img_rgba，不含画布边框假透明） ──────────
    info = []
    for layer_idx, layer_name, layer_img_rgba, _canvas in layers:
        alpha      = layer_img_rgba.split()[3]
        alpha_hist = alpha.histogram()
        total      = sum(alpha_hist)
        if total == 0:
            continue

        # 图层自身的透明比（不受画布边框影响）
        transparent_ratio = sum(alpha_hist[:50]) / total

        opaque_mask = alpha.point(lambda p: 255 if p >= 50 else 0, mode='L')
        bb = opaque_mask.getbbox()
        if bb is None:
            opaque_area = aspect = 0
            rgb_hist = None
        else:
            bw = bb[2] - bb[0]
            bh = bb[3] - bb[1]
            opaque_area = bw * bh
            aspect = bw / bh if bh > 0 else 999.0
            # 仅取不透明区域的 RGB，缩到 32×32 做直方图
            r, g, b, _ = layer_img_rgba.split()
            rgb_crop = Image.merge('RGB', (r, g, b)).crop(bb)
            rgb_hist = rgb_crop.resize((32, 32), Image.LANCZOS).histogram()

        _log(f"  [检测] 图层 {layer_idx} ({layer_name}): "
             f"图层自身透明比={transparent_ratio:.1%}, "
             f"内容面积={opaque_area/canvas_area:.1%}画布, "
             f"图层尺寸={layer_img_rgba.size}")

        info.append({
            'idx': layer_idx, 'name': layer_name,
            'transparent_ratio': transparent_ratio,
            'opaque_area': opaque_area,
            'aspect': aspect,
            'rgb_hist': rgb_hist,
        })

    # ── 路径 A：透明扣图（图层自身有 alpha） ────────────────────────────────────
    a_candidates = []
    for d in info:
        if not (0.10 <= d['transparent_ratio'] <= 0.88):
            continue
        if d['opaque_area'] < canvas_area * 0.03:
            _log(f"  [检测] 图层 {d['idx']} ({d['name']}): 内容面积太小 → A路径跳过")
            continue
        if not (0.15 <= d['aspect'] <= 6.5):
            _log(f"  [检测] 图层 {d['idx']} ({d['name']}): 宽高比 {d['aspect']:.2f} 异常 → A路径跳过")
            continue
        _log(f"  [检测] 图层 {d['idx']} ({d['name']}): ✓ 路径A候选 "
             f"透明比={d['transparent_ratio']:.1%}, 宽高比={d['aspect']:.2f}")
        a_candidates.append(d)

    # ── 路径 B：重复内容检测（两两直方图比较） ──────────────────────────────────
    b_candidates = {}
    for i in range(len(info)):
        for j in range(i + 1, len(info)):
            sim = cosine(info[i]['rgb_hist'], info[j]['rgb_hist'])
            if sim >= 0.80:
                smaller = info[i] if info[i]['opaque_area'] <= info[j]['opaque_area'] else info[j]
                _log(f"  [检测] 图层 {info[i]['idx']}↔{info[j]['idx']} "
                     f"内容相似(sim={sim:.2f}) → 小面积图层 {smaller['idx']} ({smaller['name']}) 路径B候选")
                if smaller['idx'] not in b_candidates:
                    b_candidates[smaller['idx']] = smaller

    # ── 路径 C：尺寸排除法（无 alpha 的商品图）──────────────────────────────────
    # 场景背景 = 面积最大的几乎不透明图层（填满画布）
    # 商品主体 = 次大的图层；若只有1个实质内容层则直接是商品
    c_pool = [d for d in info
              if d['transparent_ratio'] < 0.10        # 几乎不透明（无真实 alpha）
              and d['opaque_area'] > canvas_area * 0.08]  # 有实质内容（> 8% 画布）

    if c_pool:
        if len(c_pool) == 1:
            d = c_pool[0]
            _log(f"  [检测] 图层 {d['idx']} ({d['name']}): ✓ 路径C候选 "
                 f"（唯一实质内容图层，面积={d['opaque_area']/canvas_area:.1%}）")
            c_candidates = [d]
        else:
            # 排除面积最大的（= 场景背景）
            sorted_pool = sorted(c_pool, key=lambda x: x['opaque_area'], reverse=True)
            excluded = sorted_pool[0]
            _log(f"  [检测] 图层 {excluded['idx']} ({excluded['name']}): "
                 f"面积最大({excluded['opaque_area']/canvas_area:.1%}) → C路径排除（疑为场景背景）")
            c_candidates = sorted_pool[1:]
            for d in c_candidates:
                _log(f"  [检测] 图层 {d['idx']} ({d['name']}): ✓ 路径C候选 "
                     f"面积={d['opaque_area']/canvas_area:.1%}")
    else:
        c_candidates = []

    # ── 选择最终结果（优先级 B > C > A） ─────────────────────────────────────
    if b_candidates:
        best = min(b_candidates.values(), key=lambda d: d['opaque_area'])
        _log(f"  [检测] ✅ 路径B 命中：图层 {best['idx']} ({best['name']}) 为商品主体")
        return best['idx'], best['name']

    if c_candidates:
        # 路径 C 中取面积最大的（最完整的商品图）
        best = max(c_candidates, key=lambda d: d['opaque_area'])
        _log(f"  [检测] ✅ 路径C 命中：图层 {best['idx']} ({best['name']}) 为商品主体 "
             f"(面积={best['opaque_area']/canvas_area:.1%})")
        return best['idx'], best['name']

    if a_candidates:
        best = max(a_candidates, key=lambda d: d['opaque_area'])
        _log(f"  [检测] ✅ 路径A 命中：图层 {best['idx']} ({best['name']}) 为商品主体")
        return best['idx'], best['name']

    return None


# ── Case 2：API 分割（兜底） ──────────────────────────────────────────────────

def _api_identify(layers, level1_cat, level3_cat, log_fn=None):
    """
    用 piseg API 在 scenebg 图层中寻找商品主体。
    优先对最不透明（内容最多）的图层调用 API。
    返回 { layer_idx: {'rank': int, 'score': float, 'method': 'api'} } 或 {}。
    """
    def _log(msg):
        (log_fn or print)(msg)

    cates = [level1_cat, '', ''] if level1_cat else ['', '', '']

    all_labels, all_bboxes, all_seg_meta = [], [], []

    for layer_idx, layer_name, _layer_img, canvas_img in layers:
        _log(f"[扣图] → 分割图层 {layer_idx}（全画布 {canvas_img.size}）")
        seg = piseg_pil(canvas_img, cates, log_fn=log_fn)
        if seg is None:
            _log(f"[扣图]   piseg 失败，跳过图层 {layer_idx}")
            continue

        seg_ims, bboxes, seg_labels = seg
        _log(f"[扣图]   检测到 {len(seg_ims)} 个分割块")

        for s_i, (s_b64, bbox, s_label) in enumerate(zip(seg_ims, bboxes, seg_labels)):
            _log(f"[扣图]   打标签 {s_i+1}/{len(seg_ims)}")
            label = None
            for _ in range(3):
                label = tagging_base64(s_b64, level3_cat or '', bbox, s_label, log_fn=log_fn)
                if label is not None:
                    break
                time.sleep(1)
            if label is None:
                _log(f"[扣图]   打标签失败，跳过分割块 {s_i}")
                continue
            all_labels.append(label)
            all_bboxes.append(bbox)
            all_seg_meta.append((layer_idx, s_i))

    if not all_labels:
        _log("[扣图] 无有效分割块，API 可能不可用（请检查 VPN）")
        return {}

    _log(f"[扣图] 对 {len(all_labels)} 个分割块进行排名...")
    rank_raw = rank_segments(all_labels, all_bboxes, level1_cat or '', level3_cat or '', log_fn=log_fn)
    if rank_raw is None:
        _log("[扣图] 排名 API 失败")
        return {}

    try:
        rank_data  = json.loads(rank_raw['label'])
        seg_ranks  = rank_data['label']
        seg_scores = rank_data['score']
    except Exception as e:
        _log(f"[扣图] 排名结果解析失败: {e}  原始: {str(rank_raw)[:200]}")
        return {}

    layer_best: dict = {}
    for rank_pos, seg_idx in enumerate(seg_ranks):
        if seg_idx < len(all_seg_meta):
            layer_idx, _ = all_seg_meta[seg_idx]
            if layer_idx not in layer_best:
                layer_best[layer_idx] = {
                    'rank': rank_pos,
                    'score': seg_scores[rank_pos],
                    'method': 'api'
                }

    _log("[扣图] API 识别结果:")
    for layer_idx, layer_name, _li, _ci in layers:
        if layer_idx in layer_best:
            r = layer_best[layer_idx]
            _log(f"  图层 {layer_idx} ({layer_name}): rank={r['rank']}, score={r['score']:.4f}")
        else:
            _log(f"  图层 {layer_idx} ({layer_name}): 无有效分割")

    return layer_best


def identify_product_layer(psd_path, level1_cat, level3_cat, log_fn=None):
    """
    主函数：识别 PSD 中哪个 'scenebg' 图层是商品主体。

    策略（优先级从高到低）：
      Case 1 — 透明度检测（无 API）：若某图层已是扣好的主体，直接返回。
      Case 2 — piseg API：Case 1 无结果时，调用 API 分割并排名。

    返回:
        { layer_index: {'rank': int, 'score': float, 'method': str}, ... }
        rank 越小越可能是商品主体；空 dict 表示识别失败。
    """
    def _log(msg):
        (log_fn or print)(msg)

    _log(f"[扣图] 开始识别: {Path(psd_path).name}")
    _log(f"[扣图] 品类: {level1_cat or '(未填)'} / {level3_cat or '(未填)'}")

    layers = export_scenebg_layers(psd_path, log_fn=log_fn)
    if not layers:
        _log("[扣图] 未找到 scenebg 图层")
        return {}

    # ── Case 1：透明度检测 ──────────────────────────────────────────────────
    _log("[扣图] Case 1：检测已有扣图图层…")
    cutout = detect_existing_cutout(layers, log_fn=log_fn)
    if cutout is not None:
        layer_idx, layer_name = cutout
        _log(f"[扣图] ✅ Case 1 命中：图层 {layer_idx} ({layer_name}) 已是商品扣图，直接命名")
        return {layer_idx: {'rank': 0, 'score': 1.0, 'method': 'cutout'}}

    # ── Case 2：API 分割 ────────────────────────────────────────────────────
    _log("[扣图] Case 1 未找到扣图图层，尝试 Case 2：调用 piseg API…")
    return _api_identify(layers, level1_cat, level3_cat, log_fn=log_fn)


def rename_product_in_psd(psd_path, product_layer_index, new_name='product', log_fn=None):
    """将 PSD 中指定索引的图层改名为 new_name 并保存。"""
    def _log(msg):
        (log_fn or print)(msg)

    psd = PSDImage.open(psd_path)
    layers_list = list(psd)
    if product_layer_index >= len(layers_list):
        _log(f"[扣图] 图层索引 {product_layer_index} 超出范围（共 {len(layers_list)} 层）")
        return False
    layer = layers_list[product_layer_index]
    old_name = layer.name
    layer.name = new_name
    psd.save(psd_path)
    _log(f"[扣图] 图层 {product_layer_index}: '{old_name}' → '{new_name}'  已保存")
    return True


def process_output_folder(output_folder, level1_cat, level3_cat, auto_rename=True, log_fn=None):
    """
    对输出文件夹中所有 PSD 批量识别商品主体图层。
    auto_rename=True 时自动将最优图层改名为 'product'。
    """
    def _log(msg):
        (log_fn or print)(msg)

    psd_files = sorted(Path(output_folder).glob("*.psd")) + \
                sorted(Path(output_folder).glob("*.PSD"))
    summary = {}

    for psd_path in psd_files:
        results = identify_product_layer(str(psd_path), level1_cat, level3_cat, log_fn=log_fn)
        if not results:
            _log(f"[扣图] {psd_path.name}: 识别失败")
            continue

        best_idx = min(results, key=lambda k: results[k]['rank'])
        best = results[best_idx]
        method = best.get('method', 'unknown')
        summary[psd_path.name] = {
            'layer_index': best_idx,
            'rank': best['rank'],
            'score': best['score'],
            'method': method
        }

        method_label = "已有扣图（直接命名）" if method == 'cutout' else "API 识别"
        _log(f"[扣图] {psd_path.name}: 商品主体 = 图层 {best_idx}"
             f"（{method_label}，score={best['score']:.4f}）")

        if auto_rename:
            rename_product_in_psd(str(psd_path), best_idx, 'product', log_fn=log_fn)

    return summary


# ── CLI ──────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python seg_product.py <psd路径> [level1品类] [level3品类]")
        print('示例: python seg_product.py output/kettle.psd "Electronics" "Electric Kettles"')
        sys.exit(1)

    psd_p = sys.argv[1]
    lvl1  = sys.argv[2] if len(sys.argv) > 2 else ''
    lvl3  = sys.argv[3] if len(sys.argv) > 3 else ''

    if not os.path.exists(psd_p):
        print(f"文件不存在: {psd_p}")
        sys.exit(1)

    results = identify_product_layer(psd_p, lvl1, lvl3)
    if not results:
        print("\n未能识别商品主体图层")
        sys.exit(0)

    best_idx = min(results, key=lambda k: results[k]['rank'])
    best = results[best_idx]
    print(f"\n✅ 最优商品主体: 图层 {best_idx}，rank={best['rank']}，score={best['score']:.4f}")

    ans = input(f"\n是否将图层 {best_idx} 改名为 'product'？(y/N): ").strip().lower()
    if ans == 'y':
        rename_product_in_psd(psd_p, best_idx, 'product')
