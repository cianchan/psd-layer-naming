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

            # 粘贴到全画布，保留位置信息
            canvas = Image.new('RGBA', (doc_w, doc_h), (0, 0, 0, 0))
            bbox = layer.bbox   # BBox(left, top, right, bottom)
            canvas.paste(layer_img.convert('RGBA'), (bbox.left, bbox.top))

            results.append((i, layer.name, canvas))
            _log(f"  [PSD] 图层 {i}: {layer.name}  位置=({bbox.left},{bbox.top})  "
                 f"图层尺寸={layer_img.size}  画布={canvas.size}")
        except Exception as e:
            _log(f"  [PSD] 图层 {i} ({layer.name}) 导出失败: {e}")

    return results


def identify_product_layer(psd_path, level1_cat, level3_cat, log_fn=None):
    """
    主函数：识别 PSD 中哪个 'scenebg' 图层是商品主体。

    返回:
        { layer_index: {'rank': int, 'score': float}, ... }
        rank 越小越可能是商品主体；空 dict 表示识别失败。
    """
    def _log(msg):
        (log_fn or print)(msg)

    cates = [level1_cat, '', ''] if level1_cat else ['', '', '']
    _log(f"[扣图] 开始识别: {Path(psd_path).name}")
    _log(f"[扣图] 品类: {level1_cat or '(未填)'} / {level3_cat or '(未填)'}")

    layers = export_scenebg_layers(psd_path, log_fn=log_fn)
    if not layers:
        _log("[扣图] 未找到 scenebg 图层")
        return {}

    all_labels, all_bboxes, all_seg_meta = [], [], []

    for layer_idx, layer_name, canvas_img in layers:
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
                layer_best[layer_idx] = {'rank': rank_pos, 'score': seg_scores[rank_pos]}

    _log("[扣图] 识别结果:")
    for layer_idx, layer_name, _ in layers:
        if layer_idx in layer_best:
            r = layer_best[layer_idx]
            _log(f"  图层 {layer_idx} ({layer_name}): rank={r['rank']}, score={r['score']:.4f}")
        else:
            _log(f"  图层 {layer_idx} ({layer_name}): 无有效分割")

    return layer_best


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
        summary[psd_path.name] = {
            'layer_index': best_idx,
            'rank': best['rank'],
            'score': best['score']
        }
        _log(f"[扣图] {psd_path.name}: 商品主体 = 图层 {best_idx}"
             f"（rank={best['rank']}, score={best['score']:.4f}）")

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
