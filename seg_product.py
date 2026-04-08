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
import math
import re
import socket
import subprocess
import tempfile
import time
import requests
from io import BytesIO
from pathlib import Path

from PIL import Image
from psd_tools import PSDImage

# ── Shopee API 配置 ─────────────────────────────────────────────────────────

_DEFAULT_GATEWAY = "https://http-gateway.spex.shopee.sg/sprpc"
_DEFAULT_SDU = "ai_engine_platform.mmuplt.controller.global.liveish.master.default"
_DEFAULT_SERVICEKEY = "f0dd2d544097d2a938595c1d78949bd3"


def _mask_secret(secret):
    if not secret:
        return "<empty>"
    if len(secret) <= 8:
        return "*" * len(secret)
    return f"{secret[:4]}...{secret[-4:]}"


def _api_config():
    gateway = os.getenv("SHOPEE_AI_GATEWAY", _DEFAULT_GATEWAY).rstrip("/")
    sdu = os.getenv("SHOPEE_AI_SDU", _DEFAULT_SDU)
    servicekey = os.getenv("SHOPEE_AI_SERVICEKEY", _DEFAULT_SERVICEKEY)
    return {
        "gateway": gateway,
        "sdu": sdu,
        "servicekey": servicekey,
    }


def _build_headers(processid, timeout):
    cfg = _api_config()
    return {
        "Content-Type": "application/json",
        "x-sp-sdu": cfg["sdu"],
        "x-sp-servicekey": cfg["servicekey"],
        "x-sp-timeout": str(timeout),
        "x-sp-processid": processid,
    }


def call_service(service_name, request_data, env='liveish', max_retries=3, log_fn=None):
    """调用 Shopee AI 服务（需要 Shopee VPN）。所有错误通过 log_fn 输出。"""
    def _log(msg):
        (log_fn or print)(msg)

    cfg = _api_config()
    url = f"{cfg['gateway']}/ai_engine_platform.mmuplt.{service_name}.algo"
    headers = _build_headers("CID=global", 30000) if service_name == 'seglabel2' else _build_headers("process_1", 60000)
    host = url.split("/")[2]

    try:
        resolved_ip = socket.gethostbyname(host)
    except Exception:
        resolved_ip = None

    for attempt in range(max_retries):
        try:
            if attempt == 0:
                _log(
                    f"  [API] {service_name} host={host} "
                    f"ip={resolved_ip or 'unresolved'} "
                    f"sdu={cfg['sdu']} servicekey={_mask_secret(cfg['servicekey'])}"
                )
            resp = requests.post(url, data=json.dumps(request_data), headers=headers, timeout=60)
            if resp.status_code != 200:
                gateway_msg = resp.headers.get("sgw-errmsg")
                _log(
                    f"  [API] {service_name} HTTP {resp.status_code}"
                    f"{f' sgw-errmsg={gateway_msg}' if gateway_msg else ''}: "
                    f"{resp.text[:300]}"
                )
                if resp.status_code == 403:
                    _log("  [API] 403 表示请求在网关层被拒绝，通常是 VPN / 白名单 / servicekey 权限问题，不是图片内容或 JSON 格式问题。")
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


def base64_to_pil(base64_str):
    """Decode a base64-encoded image returned by the segmentation service."""
    image_data = base64.b64decode(base64_str)
    return Image.open(BytesIO(image_data)).convert("RGBA")


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


def probe_api(log_fn=None):
    """Minimal connectivity probe for the first-stage segmentation endpoint."""
    def _log(msg):
        (log_fn or print)(msg)

    _log("[探测] 开始探测 Shopee segmentation API 网关")
    req = {
        "biz_type": "mmu_test",
        "region": "sg2",
        "task": {"image_list": []},
    }
    res = call_service("pisegv2", req, log_fn=log_fn, max_retries=1)
    if res is None:
        _log("[探测] 结果: 接口未通过。若日志中出现 403，优先检查 VPN / 白名单 / servicekey 权限。")
        return False
    _log("[探测] 结果: 网关可用。")
    return True


# ── PSD 图层处理 ─────────────────────────────────────────────────────────────

def _flatten_layers(root):
    """Depth-first flatten all layers so rename/read use the same stable index space."""
    flat_layers = []

    def walk(container, prefix=''):
        for layer in container:
            layer_name = (layer.name or '').strip() or '<unnamed>'
            layer_path = f"{prefix}/{layer_name}" if prefix else layer_name
            flat_layers.append({
                'layer_index': len(flat_layers),
                'layer_path': layer_path,
                'layer': layer,
            })
            if layer.is_group():
                walk(layer, layer_path)

    walk(root)
    return flat_layers


def _is_scenebg_name(layer_name):
    """Match scenebg, scenebg1, scenebg2... to be more tolerant of legacy files."""
    return bool(re.fullmatch(r'scenebg\d*', (layer_name or '').strip(), flags=re.IGNORECASE))


def _extract_local_metrics(canvas_img, doc_size, alpha_threshold=8):
    """Estimate how much a layer looks like a foreground product cutout."""
    doc_w, doc_h = doc_size
    doc_area = max(doc_w * doc_h, 1)
    alpha = canvas_img.getchannel('A')
    mask = alpha.point(lambda px: 255 if px > alpha_threshold else 0)
    bbox = mask.getbbox()
    if not bbox:
        return {
            'bbox': None,
            'opaque_pixels': 0,
            'opaque_ratio': 0.0,
            'bbox_ratio': 0.0,
            'fill_ratio': 0.0,
            'center_score': 0.0,
            'edge_touches': 0,
            'width_ratio': 0.0,
            'height_ratio': 0.0,
            'score': 0.0,
        }

    left, top, right, bottom = bbox
    bbox_w = max(right - left, 1)
    bbox_h = max(bottom - top, 1)
    bbox_area = max(bbox_w * bbox_h, 1)
    opaque_pixels = mask.histogram()[255]

    opaque_ratio = opaque_pixels / doc_area
    bbox_ratio = bbox_area / doc_area
    fill_ratio = opaque_pixels / bbox_area
    width_ratio = bbox_w / max(doc_w, 1)
    height_ratio = bbox_h / max(doc_h, 1)

    center_x = (left + right) / 2.0
    center_y = (top + bottom) / 2.0
    max_dist = math.hypot(doc_w / 2.0, doc_h / 2.0) or 1.0
    center_dist = math.hypot(center_x - doc_w / 2.0, center_y - doc_h / 2.0)
    center_score = max(0.0, 1.0 - (center_dist / max_dist))

    edge_margin_x = max(4, int(doc_w * 0.02))
    edge_margin_y = max(4, int(doc_h * 0.02))
    edge_touches = sum([
        left <= edge_margin_x,
        top <= edge_margin_y,
        (doc_w - right) <= edge_margin_x,
        (doc_h - bottom) <= edge_margin_y,
    ])

    # Prefer medium/large isolated objects. Penalize full-canvas or edge-hugging layers.
    prominence = opaque_ratio * (0.7 + 0.3 * center_score) * (0.5 + min(fill_ratio, 1.0))
    prominence *= 0.6 + min(bbox_ratio / 0.15, 1.0)

    if bbox_ratio > 0.65:
        prominence *= 0.15
    if width_ratio > 0.88:
        prominence *= 0.35
    if height_ratio > 0.88:
        prominence *= 0.35
    if edge_touches >= 3:
        prominence *= 0.15
    elif edge_touches == 2:
        prominence *= 0.45
    if fill_ratio < 0.08:
        prominence *= 0.4

    return {
        'bbox': bbox,
        'opaque_pixels': opaque_pixels,
        'opaque_ratio': opaque_ratio,
        'bbox_ratio': bbox_ratio,
        'fill_ratio': fill_ratio,
        'center_score': center_score,
        'edge_touches': edge_touches,
        'width_ratio': width_ratio,
        'height_ratio': height_ratio,
        'score': prominence,
    }


def _bbox_to_xyxy(bbox):
    if bbox is None:
        return None
    if hasattr(bbox, 'left'):
        return int(bbox.left), int(bbox.top), int(bbox.right), int(bbox.bottom)
    return int(bbox[0]), int(bbox[1]), int(bbox[2]), int(bbox[3])


def _segment_to_full_canvas(seg_b64, bbox, doc_size):
    """Paste a cropped segmentation image back onto a transparent full-canvas image."""
    doc_w, doc_h = doc_size
    seg_img = base64_to_pil(seg_b64)
    full_canvas = Image.new('RGBA', (doc_w, doc_h), (0, 0, 0, 0))
    left, top, _, _ = _bbox_to_xyxy(bbox)
    full_canvas.paste(seg_img, (left, top), seg_img)
    return full_canvas


def _should_split_candidate(source_metrics, segment_metrics):
    """
    Decide whether the best API segment should become a new product layer instead of
    simply renaming the original scenebg layer.
    """
    src_pixels = max(source_metrics.get('opaque_pixels', 0), 1)
    seg_pixels = segment_metrics.get('opaque_pixels', 0)
    coverage = seg_pixels / src_pixels

    if coverage >= 0.92:
        return False
    if coverage <= 0.75:
        return True
    if source_metrics.get('edge_touches', 0) >= 2 and coverage < 0.88:
        return True
    if source_metrics.get('bbox_ratio', 0.0) > 0.55 and coverage < 0.9:
        return True
    return False


def find_photoshop_app_name():
    """Return the exact installed Photoshop app name for AppleScript tell blocks."""
    try:
        result = subprocess.run(
            ["mdfind", "kMDItemCFBundleIdentifier == 'com.adobe.Photoshop'"],
            capture_output=True, text=True, timeout=10
        )
        for path in result.stdout.strip().splitlines():
            if path.endswith(".app"):
                name = os.path.basename(path).replace(".app", "")
                if "Photoshop" in name:
                    return name
    except Exception:
        pass
    return "Adobe Photoshop"


def _run_photoshop_jsx(jsx_code, log_fn=None, timeout=180):
    """Execute a JSX snippet in Photoshop via AppleScript."""
    def _log(msg):
        (log_fn or print)(msg)

    with tempfile.NamedTemporaryFile('w', encoding='utf-8', suffix='.jsx', delete=False) as jsx_file:
        jsx_file.write(jsx_code)
        jsx_path = jsx_file.name

    jsx_path_js = jsx_path.replace("\\", "/").replace("'", "\\'")
    jsx_loader = (
        f"var _f=new File('{jsx_path_js}');"
        "_f.open('r');var _s=_f.read();_f.close();eval(_s);"
    )
    ps_app_name = find_photoshop_app_name().replace('"', '\\"')

    with tempfile.NamedTemporaryFile('w', encoding='utf-8', suffix='.applescript', delete=False) as as_file:
        as_file.write(
            f'tell application "{ps_app_name}"\n'
            f'    activate\n'
            f'    do javascript "{jsx_loader}"\n'
            f'end tell\n'
        )
        as_path = as_file.name

    try:
        result = subprocess.run(
            ["osascript", as_path],
            capture_output=True, text=True, timeout=timeout
        )
        if result.returncode != 0:
            stderr = (result.stderr or result.stdout or "").strip()
            _log(f"[PS] JSX 执行失败: {stderr}")
            return False
        return True
    finally:
        Path(jsx_path).unlink(missing_ok=True)
        Path(as_path).unlink(missing_ok=True)


def split_product_in_psd(psd_path, source_layer_index, product_png_path, new_name='product', log_fn=None):
    """
    Insert a segmented product PNG as a new Photoshop layer and clear those pixels
    from the original source layer, so the product is actually split out.
    """
    def _log(msg):
        (log_fn or print)(msg)

    psd_path_js = str(psd_path).replace("\\", "/").replace("'", "\\'")
    product_png_js = str(product_png_path).replace("\\", "/").replace("'", "\\'")
    new_name_js = new_name.replace("\\", "\\\\").replace('"', '\\"')

    jsx = f"""#target photoshop
app.displayDialogs = DialogModes.NO;

function flattenLayers(container, acc) {{
    for (var i = 0; i < container.layers.length; i++) {{
        var layer = container.layers[i];
        acc.push(layer);
        if (layer.typename === "LayerSet") {{
            flattenLayers(layer, acc);
        }}
    }}
}}

function selectTransparency(layer) {{
    var desc = new ActionDescriptor();
    var ref = new ActionReference();
    ref.putProperty(charIDToTypeID("Chnl"), charIDToTypeID("fsel"));
    desc.putReference(charIDToTypeID("null"), ref);
    var ref2 = new ActionReference();
    ref2.putEnumerated(charIDToTypeID("Chnl"), charIDToTypeID("Chnl"), charIDToTypeID("Trsp"));
    ref2.putIdentifier(charIDToTypeID("Lyr "), layer.id);
    desc.putReference(charIDToTypeID("T   "), ref2);
    executeAction(charIDToTypeID("setd"), desc, DialogModes.NO);
}}

var psdFile = new File('{psd_path_js}');
var pngFile = new File('{product_png_js}');
if (!psdFile.exists) throw new Error("PSD file not found");
if (!pngFile.exists) throw new Error("Product PNG not found");

var doc = app.open(psdFile);
var flat = [];
flattenLayers(doc, flat);
if ({source_layer_index} >= flat.length) throw new Error("Source layer index out of range");
var sourceLayer = flat[{source_layer_index}];
if (sourceLayer.typename === "LayerSet") throw new Error("Source layer points to group");

var segDoc = app.open(pngFile);
segDoc.selection.selectAll();
segDoc.selection.copy();
segDoc.close(SaveOptions.DONOTSAVECHANGES);

app.activeDocument = doc;
var productLayer = doc.paste();
productLayer.name = "{new_name_js}";
doc.activeLayer = productLayer;
selectTransparency(productLayer);

doc.activeLayer = sourceLayer;
try {{ sourceLayer.allLocked = false; }} catch(e) {{}}
try {{ sourceLayer.transparentPixelsLocked = false; }} catch(e) {{}}
try {{
    if (sourceLayer.kind !== LayerKind.NORMAL) {{
        sourceLayer.rasterize(RasterizeType.ENTIRELAYER);
    }}
}} catch(e) {{}}
doc.selection.clear();
doc.selection.deselect();

doc.activeLayer = productLayer;
var opts = new PhotoshopSaveOptions();
opts.layers = true;
opts.embedColorProfile = true;
opts.annotations = false;
opts.alphaChannels = true;
opts.spotColors = true;
doc.saveAs(psdFile, opts, true);
doc.close(SaveOptions.DONOTSAVECHANGES);
"""

    ok = _run_photoshop_jsx(jsx, log_fn=log_fn)
    if ok:
        _log(f"[扣图] 已将分割主体写回 PSD，并命名为 '{new_name}'")
    return ok


def _rank_local_candidates(layers, log_fn=None):
    """Rank scenebg layers using local geometry so we still work without VPN/API."""
    def _log(msg):
        (log_fn or print)(msg)

    ranked = []
    for item in layers:
        metrics = item['metrics']
        if metrics['opaque_pixels'] <= 0:
            _log(f"[本地识别] 跳过空图层 #{item['layer_index']} {item['layer_path']}")
            continue

        left, top, right, bottom = metrics['bbox']
        _log(
            f"[本地识别] 图层 #{item['layer_index']} {item['layer_path']}: "
            f"bbox=({left},{top},{right},{bottom}) "
            f"opaque={metrics['opaque_ratio']:.4f} "
            f"fill={metrics['fill_ratio']:.4f} "
            f"center={metrics['center_score']:.4f} "
            f"edges={metrics['edge_touches']} "
            f"score={metrics['score']:.6f}"
        )
        ranked.append((item['layer_index'], item['layer_path'], metrics['score']))

    ranked.sort(key=lambda entry: (-entry[2], entry[0]))

    results = {}
    if ranked:
        _log("[本地识别] 排名结果:")
    for rank, (layer_index, layer_path, score) in enumerate(ranked):
        _log(f"  #{rank + 1}: 图层 {layer_index} ({layer_path}) score={score:.6f}")
        results[layer_index] = {
            'rank': rank,
            'score': score,
            'method': 'local',
            'layer_path': layer_path,
            'split_needed': False,
            'action': 'rename',
        }
    return results

def export_scenebg_layers(psd_path, log_fn=None):
    """
    打开处理后的 PSD，将所有名为 'scenebg' 的图层导出为 PIL image。
    关键：每个图层粘贴在与 PSD 等大的全画布上（透明背景），
    保留图层在画面中的位置上下文，让 API / 本地启发式都能判断主体。

    返回 [{layer_index, layer_name, layer_path, layer_img_rgba, canvas, metrics}, ...]
    """
    def _log(msg):
        (log_fn or print)(msg)

    psd = PSDImage.open(psd_path)
    doc_w, doc_h = psd.width, psd.height
    _log(f"  [PSD] 文档尺寸: {doc_w}×{doc_h}")

    results = []
    for flat in _flatten_layers(psd):
        layer = flat['layer']
        if layer.is_group() or not _is_scenebg_name(layer.name):
            continue
        try:
            layer_img = layer.composite()
            if layer_img is None:
                _log(f"  [PSD] 图层 {flat['layer_index']} composite() 返回 None，跳过")
                continue

            layer_img_rgba = layer_img.convert('RGBA')   # 图层自身（用于分析）

            # 粘贴到全画布，保留位置信息（用于 API 调用）
            canvas = Image.new('RGBA', (doc_w, doc_h), (0, 0, 0, 0))
            bbox = layer.bbox   # tuple or BBox(left, top, right, bottom)
            if hasattr(bbox, 'left'):
                left, top = bbox.left, bbox.top
            else:
                left, top = int(bbox[0]), int(bbox[1])
            canvas.paste(layer_img_rgba, (left, top))
            metrics = _extract_local_metrics(canvas, (doc_w, doc_h))

            results.append({
                'layer_index': flat['layer_index'],
                'layer_name': layer.name,
                'layer_path': flat['layer_path'],
                'layer_img_rgba': layer_img_rgba,
                'canvas': canvas,
                'metrics': metrics,
            })
            _log(
                f"  [PSD] 图层 {flat['layer_index']}: {flat['layer_path']}  "
                f"位置=({left},{top}) 图层尺寸={layer_img.size} 画布={canvas.size}"
            )
        except Exception as e:
            _log(f"  [PSD] 图层 {flat['layer_index']} ({flat['layer_path']}) 导出失败: {e}")

    return results


# ── Case 1：透明度 / 重复内容 / 尺寸排除 三路检测（无需 API） ───────────────

def detect_existing_cutout(layers, log_fn=None):
    """
    检测 scenebg 图层中是否已有商品主体图层。

    layers 格式：[{layer_index, layer_path, layer_img_rgba, canvas, ...}, ...]
      - layer_img_rgba：图层自身裁剪区域（用于分析，不含画布填充边框）
      - canvas：全画布版本（保留给 API 调用，此函数不使用）

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
    canvas_area = layers[0]['canvas'].width * layers[0]['canvas'].height

    def cosine(h1, h2):
        if not h1 or not h2:
            return 0.0
        dot  = sum(a * b for a, b in zip(h1, h2))
        mag1 = sum(a * a for a in h1) ** 0.5
        mag2 = sum(b * b for b in h2) ** 0.5
        return dot / (mag1 * mag2) if (mag1 and mag2) else 0.0

    # ── 预计算每层指标（全部基于 layer_img_rgba，不含画布边框假透明） ──────────
    info = []
    for item in layers:
        layer_idx = item['layer_index']
        layer_name = item['layer_path']
        layer_img_rgba = item['layer_img_rgba']
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

    local_results = _rank_local_candidates(layers, log_fn=log_fn)
    if not local_results:
        _log("[本地识别] 没有可用的 scenebg 候选图层")
        return {}

    if not level1_cat and not level3_cat:
        _log("[扣图] 未提供品类，尝试使用通用 API 分割；若失败则回退到本地启发式结果")

    cates = [level1_cat or '', '', '']
    all_labels, all_bboxes, all_seg_meta = [], [], []

    for item in layers:
        layer_idx = item['layer_index']
        canvas_img = item['canvas']
        if item['metrics']['opaque_pixels'] <= 0:
            _log(f"[扣图] → 跳过空图层 {layer_idx}（{item['layer_path']}）")
            continue

        _log(f"[扣图] → 分割图层 {layer_idx}（{item['layer_path']}，全画布 {canvas_img.size}）")
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
            seg_canvas = _segment_to_full_canvas(s_b64, bbox, canvas_img.size)
            seg_metrics = _extract_local_metrics(seg_canvas, canvas_img.size)
            all_seg_meta.append({
                'layer_index': layer_idx,
                'segment_index': s_i,
                'segment_base64': s_b64,
                'segment_bbox': bbox,
                'segment_label': s_label,
                'layer_path': item['layer_path'],
                'source_metrics': item['metrics'],
                'segment_metrics': seg_metrics,
                'doc_size': canvas_img.size,
            })

    if not all_labels:
        _log("[扣图] 无有效分割块，API 可能不可用，回退到本地启发式结果")
        return local_results

    _log(f"[扣图] 对 {len(all_labels)} 个分割块进行排名...")
    rank_raw = rank_segments(all_labels, all_bboxes, level1_cat or '', level3_cat or '', log_fn=log_fn)
    if rank_raw is None:
        _log("[扣图] 排名 API 失败，回退到本地启发式结果")
        return local_results

    try:
        rank_data  = json.loads(rank_raw['label'])
        seg_ranks  = rank_data['label']
        seg_scores = rank_data['score']
    except Exception as e:
        _log(f"[扣图] 排名结果解析失败: {e}  原始: {str(rank_raw)[:200]}")
        _log("[扣图] 回退到本地启发式结果")
        return local_results

    layer_best: dict = {}
    for rank_pos, seg_idx in enumerate(seg_ranks):
        if seg_idx < len(all_seg_meta):
            seg_meta = all_seg_meta[seg_idx]
            layer_idx = seg_meta['layer_index']
            if layer_idx not in layer_best:
                split_needed = _should_split_candidate(seg_meta['source_metrics'], seg_meta['segment_metrics'])
                layer_best[layer_idx] = {
                    'rank': rank_pos,
                    'score': seg_scores[rank_pos],
                    'method': 'api',
                    'layer_path': seg_meta['layer_path'],
                    'segment_base64': seg_meta['segment_base64'],
                    'segment_bbox': seg_meta['segment_bbox'],
                    'segment_label': seg_meta['segment_label'],
                    'source_metrics': seg_meta['source_metrics'],
                    'segment_metrics': seg_meta['segment_metrics'],
                    'doc_size': seg_meta['doc_size'],
                    'split_needed': split_needed,
                    'action': 'split' if split_needed else 'rename',
                }

    _log("[扣图] 识别结果:")
    for item in layers:
        layer_idx = item['layer_index']
        layer_name = item['layer_path']
        if layer_idx in layer_best:
            r = layer_best[layer_idx]
            _log(f"  图层 {layer_idx} ({layer_name}): rank={r['rank']}, score={r['score']:.4f}")
        else:
            _log(f"  图层 {layer_idx} ({layer_name}): 无有效分割")

    if not layer_best:
        _log("[扣图] API 未返回可用候选，回退到本地启发式结果")
        return local_results

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
    layers_list = _flatten_layers(psd)
    if product_layer_index >= len(layers_list):
        _log(f"[扣图] 图层索引 {product_layer_index} 超出范围（共 {len(layers_list)} 层）")
        return False
    layer_info = layers_list[product_layer_index]
    layer = layer_info['layer']
    old_name = layer.name
    layer.name = new_name
    psd.save(psd_path)
    _log(
        f"[扣图] 图层 {product_layer_index} ({layer_info['layer_path']}): "
        f"'{old_name}' → '{new_name}'  已保存"
    )
    return True


def materialize_segment_png(seg_base64, bbox, doc_size, log_fn=None):
    """Write the winning segmentation result to a transparent full-canvas PNG."""
    def _log(msg):
        (log_fn or print)(msg)

    full_canvas = _segment_to_full_canvas(seg_base64, bbox, doc_size)
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
    tmp_path = Path(tmp.name)
    tmp.close()
    full_canvas.save(tmp_path, format='PNG')
    _log(f"[扣图] 已生成临时 product PNG: {tmp_path}")
    return tmp_path


def process_output_folder(output_folder, level1_cat, level3_cat, auto_rename=True, log_fn=None):
    """
    对输出文件夹中所有 PSD 批量识别商品主体图层。
    auto_rename=True 时自动将最优图层改名为 'product'。
    无品类或 Shopee 接口不可用时，自动回退到本地启发式识别。
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
            'action': best.get('action', 'rename'),
            'method': method,
        }

        if method == 'cutout':
            method_label = "已有扣图（直接命名）"
        elif method == 'local':
            method_label = "本地启发式识别"
        else:
            method_label = "API 识别"
        _log(f"[扣图] {psd_path.name}: 商品主体 = 图层 {best_idx}"
             f"（{method_label}，score={best['score']:.4f}）")

        if auto_rename:
            if best.get('action') == 'split' and best.get('segment_base64'):
                tmp_png_path = None
                try:
                    tmp_png_path = materialize_segment_png(
                        best['segment_base64'],
                        best['segment_bbox'],
                        best['doc_size'],
                        log_fn=log_fn
                    )
                    ok = split_product_in_psd(
                        str(psd_path),
                        best_idx,
                        str(tmp_png_path),
                        'product',
                        log_fn=log_fn
                    )
                    if not ok:
                        _log("[扣图] 分割写回失败，回退为直接改名")
                        rename_product_in_psd(str(psd_path), best_idx, 'product', log_fn=log_fn)
                        summary[psd_path.name]['action'] = 'rename_fallback'
                finally:
                    if tmp_png_path:
                        Path(tmp_png_path).unlink(missing_ok=True)
            else:
                rename_product_in_psd(str(psd_path), best_idx, 'product', log_fn=log_fn)

    return summary


# ── CLI ──────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    if len(sys.argv) >= 2 and sys.argv[1] == '--probe-api':
        ok = probe_api()
        sys.exit(0 if ok else 2)

    if len(sys.argv) < 2:
        print("用法: python seg_product.py <psd路径> [level1品类] [level3品类]")
        print('示例: python seg_product.py output/kettle.psd "Electronics" "Electric Kettles"')
        print('探测: python seg_product.py --probe-api')
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
