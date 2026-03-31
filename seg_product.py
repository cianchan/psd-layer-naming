"""
seg_product.py — Shopee 扣图 API 集成模块
识别 PSD 输出文件中哪个 scenebg 图层是商品主体，
并将该图层重命名为 "product"。

依赖:
    pip install psd-tools pillow requests
    pip install imageSdk core  # Shopee 内部源（仅 pil2url 上传时需要）

CLI 用法:
    python seg_product.py <psd_path> <level1_category> <level3_category>
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


def call_service(service_name, request_data, env='liveish', max_retries=3):
    """调用 Shopee AI 服务。需要在 Shopee VPN 环境下运行。"""
    url = (
        f"https://http-gateway.spex.shopee.sg/sprpc/"
        f"ai_engine_platform.mmuplt.{service_name}.algo"
    )
    headers = _HEADERS_SEGLABEL if service_name == 'seglabel2' else _HEADERS_LIVEISH
    for attempt in range(max_retries):
        try:
            resp = requests.post(url, data=json.dumps(request_data), headers=headers, timeout=60)
            if resp.status_code != 200:
                print(f"  [API] {service_name} → {resp.status_code}")
                if attempt < max_retries - 1:
                    time.sleep(1)
                    continue
                return None
            rj = resp.json()
            if 'task_result' not in rj:
                if attempt < max_retries - 1:
                    time.sleep(1)
                    continue
                return None
            return rj['task_result']
        except Exception as e:
            print(f"  [API] {service_name} error: {e}")
            if attempt < max_retries - 1:
                time.sleep(1)
                continue
    return None


# ── 图像工具函数 ─────────────────────────────────────────────────────────────

def pil_to_base64(pil_image):
    buf = BytesIO()
    pil_image.convert("RGBA").save(buf, format='PNG')
    return base64.b64encode(buf.getvalue()).decode('utf-8')


def piseg_pil(pil_image, cates, env='liveish'):
    """
    对一张 PIL 图片调用分割服务（pisegv2）。
    返回 (seg_ims_base64, bboxes, seg_labels)，失败返回 None。
    无需 imageSdk —— 分割结果图片保留为 base64，不上传 CDN。
    """
    b64 = pil_to_base64(pil_image)
    extra = json.dumps({'is_upload': False, 'cates': cates, 'bboxes': []})
    req = {
        "biz_type": "mmu_test",
        "region": "sg2",
        "task": {
            "image_list": [{"image_data": b64, "extra_info": extra}]
        }
    }
    res = call_service('pisegv2', req, env=env)
    if res is None:
        return None
    try:
        ei = json.loads(res['extra_info'])
        return ei['object_images'], ei['object_bboxes'], ei['object_labels']
    except Exception as e:
        print(f"  [API] piseg parse error: {e}")
        return None


def tagging_base64(source_pil, seg_b64, level3_cat, bbox, seg_label, env='liveish'):
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
    res = call_service('seglabel2', req, env=env)
    if res is None:
        return None
    try:
        lst = json.loads(res['extra_info'])[0]
        return lst if lst else None
    except Exception as e:
        print(f"  [API] tagging parse error: {e}")
        return None


def rank_segments(label_list, bbox_list, level1_cat, level3_cat, env='liveish'):
    """对所有分割块一起打分排名。返回排名字典或 None。"""
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
    return call_service('pfilter2', req, env=env)


# ── PSD 图层处理 ─────────────────────────────────────────────────────────────

def export_scenebg_layers(psd_path):
    """
    用 psd-tools 打开已处理的 PSD，导出所有名为 'scenebg' 的图层为 PIL image。
    返回 [(layer_index, layer_name, pil_image), ...]
    """
    psd = PSDImage.open(psd_path)
    results = []
    layers_list = list(psd)
    for i, layer in enumerate(layers_list):
        if layer.name == 'scenebg':
            try:
                pil = layer.composite()
                if pil is not None and pil.size[0] > 0 and pil.size[1] > 0:
                    results.append((i, layer.name, pil))
                    print(f"  [PSD] 导出图层 {i}: {layer.name}  尺寸={pil.size}")
            except Exception as e:
                print(f"  [PSD] 图层 {i} ({layer.name}) 导出失败: {e}")
    return results


def identify_product_layer(psd_path, level1_cat, level3_cat, log_fn=None):
    """
    主函数：识别 PSD 中哪个 'scenebg' 图层是商品主体。

    参数:
        psd_path    — 已处理的输出 PSD 路径
        level1_cat  — 一级品类（如 "Electronics"）
        level3_cat  — 三级品类（如 "Electric Kettles"）
        log_fn      — 可选日志回调 fn(str)，不传则用 print

    返回:
        {
          layer_index: {'rank': int, 'score': float},
          ...
        }
        rank 越小越可能是商品主体；空 dict 表示识别失败。
    """
    def _log(msg):
        if log_fn:
            log_fn(msg)
        else:
            print(msg)

    cates = [level1_cat, '', '']
    _log(f"[扣图] 开始识别: {Path(psd_path).name}")
    _log(f"[扣图] 品类: {level1_cat} / {level3_cat}")

    layers = export_scenebg_layers(psd_path)
    if not layers:
        _log("[扣图] 未找到 scenebg 图层")
        return {}

    all_labels, all_bboxes, all_seg_meta = [], [], []

    for layer_idx, layer_name, pil_image in layers:
        _log(f"[扣图] → 分割图层 {layer_idx}，尺寸 {pil_image.size}")
        seg = piseg_pil(pil_image, cates)
        if seg is None:
            _log(f"[扣图]   piseg 失败，跳过图层 {layer_idx}")
            continue

        seg_ims, bboxes, seg_labels = seg
        _log(f"[扣图]   检测到 {len(seg_ims)} 个分割块")

        for s_i, (s_b64, bbox, s_label) in enumerate(zip(seg_ims, bboxes, seg_labels)):
            _log(f"[扣图]   打标签 {s_i+1}/{len(seg_ims)}")
            label = None
            for _ in range(3):
                label = tagging_base64(pil_image, s_b64, level3_cat, bbox, s_label)
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
    rank_raw = rank_segments(all_labels, all_bboxes, level1_cat, level3_cat)
    if rank_raw is None:
        _log("[扣图] 排名 API 失败")
        return {}

    try:
        rank_data  = json.loads(rank_raw['label'])
        seg_ranks  = rank_data['label']   # 按质量排序的分割块索引列表
        seg_scores = rank_data['score']
    except Exception as e:
        _log(f"[扣图] 排名结果解析失败: {e}")
        return {}

    # 将排名映射回图层索引
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
    """
    将 PSD 中指定索引的图层改名为 new_name 并保存。
    注意：psd-tools 保存时会保留图层数据，但部分 PS 特效可能需要重新打开确认。
    """
    def _log(msg):
        if log_fn:
            log_fn(msg)
        else:
            print(msg)

    psd = PSDImage.open(psd_path)
    layers_list = list(psd)
    if product_layer_index >= len(layers_list):
        _log(f"[扣图] 图层索引 {product_layer_index} 超出范围")
        return False
    layer = layers_list[product_layer_index]
    old_name = layer.name
    layer.name = new_name
    psd.save(psd_path)
    _log(f"[扣图] 图层 {product_layer_index}: '{old_name}' → '{new_name}'，已保存")
    return True


def process_output_folder(output_folder, level1_cat, level3_cat, auto_rename=False, log_fn=None):
    """
    对输出文件夹中所有 PSD 批量识别商品主体图层。
    auto_rename=True 时自动将最优图层改名为 'product'。
    返回 {psd_filename: {'layer_index': N, 'rank': X, 'score': Y}, ...}
    """
    def _log(msg):
        if log_fn:
            log_fn(msg)
        else:
            print(msg)

    psd_files = sorted(Path(output_folder).glob("*.psd")) + sorted(Path(output_folder).glob("*.PSD"))
    summary = {}

    for psd_path in psd_files:
        results = identify_product_layer(str(psd_path), level1_cat, level3_cat, log_fn)
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
        _log(f"[扣图] {psd_path.name}: 商品主体 = 图层 {best_idx}（rank={best['rank']}, score={best['score']:.4f}）")

        if auto_rename:
            rename_product_in_psd(str(psd_path), best_idx, 'product', log_fn)

    return summary


# ── CLI ──────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    if len(sys.argv) < 4:
        print("用法: python seg_product.py <psd路径> <level1品类> <level3品类>")
        print('示例: python seg_product.py output/kettle.psd "Electronics" "Electric Kettles"')
        sys.exit(1)

    psd_p   = sys.argv[1]
    lvl1    = sys.argv[2]
    lvl3    = sys.argv[3]

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
