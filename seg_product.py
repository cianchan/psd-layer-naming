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
import threading
import time
import requests
from io import BytesIO
from pathlib import Path

import numpy as np
from PIL import Image
from psd_tools import PSDImage

# #region agent log
_DBG_PATH = "/Users/ziyan.chen/Documents/Sophie's File/vibe coding/psd layer naming/.cursor/debug-c3ab93.log"
def _dbg(loc, msg, data=None, hid=""):
    try:
        with open(_DBG_PATH, 'a') as _f:
            _f.write(json.dumps({"sessionId":"c3ab93","location":loc,"message":msg,"data":data or {},"timestamp":int(time.time()*1000),"hypothesisId":hid}) + '\n')
    except: pass
# #endregion

# ── Shopee API 配置 ─────────────────────────────────────────────────────────

_DEFAULT_GATEWAY       = "https://http-gateway.spex.shopee.sg/sprpc"
_DEFAULT_PROXY_GATEWAY = "https://http-gateway-proxy.spex.shopee.sg/sprpc"
_DEFAULT_SDU           = "ai_engine_platform.mmuplt.controller.global.liveish.master.default"
_DEFAULT_SERVICEKEY    = "f0dd2d544097d2a938595c1d78949bd3"

# Per-thread token so Flask request threads don't interfere with each other
_tls = threading.local()

def set_request_space_token(token: str):
    _tls.space_token = token

def clear_request_space_token():
    _tls.space_token = None

def _effective_space_token():
    return getattr(_tls, 'space_token', None) \
        or os.getenv("SHOPEE_USER_SPACE_TOKEN") \
        or os.getenv("SHOPEE_AI_SPACE_TOKEN") \
        or ""


def _mask_secret(secret):
    if not secret:
        return "<empty>"
    if len(secret) <= 8:
        return "*" * len(secret)
    return f"{secret[:4]}...{secret[-4:]}"


def _api_config():
    explicit_gw = os.getenv("SHOPEE_AI_GATEWAY", "").strip()
    use_proxy   = os.getenv("SHOPEE_AI_OFFICE_PROXY", "").strip() == "1"
    token       = _effective_space_token()

    if explicit_gw:
        gateway = explicit_gw.rstrip("/")
    elif use_proxy or token:
        gateway = _DEFAULT_PROXY_GATEWAY
    else:
        gateway = _DEFAULT_GATEWAY

    return {
        "gateway":    gateway,
        "sdu":        os.getenv("SHOPEE_AI_SDU", _DEFAULT_SDU),
        "servicekey": os.getenv("SHOPEE_AI_SERVICEKEY", _DEFAULT_SERVICEKEY),
        "space_token": token,
    }


def _build_headers(cfg, processid, timeout):
    headers = {
        "Content-Type":    "application/json",
        "x-sp-sdu":        cfg["sdu"],
        "x-sp-servicekey": cfg["servicekey"],
        "x-sp-timeout":    str(timeout),
        "x-sp-processid":  processid,
    }
    if cfg.get("space_token"):
        headers["Authorization"] = f"Bearer {cfg['space_token']}"
    return headers


def call_service(service_name, request_data, env='liveish', max_retries=3, log_fn=None):
    """调用 Shopee AI 服务。办公网需走 proxy + Bearer token。"""
    def _log(msg):
        (log_fn or print)(msg)

    cfg = _api_config()
    url = f"{cfg['gateway']}/ai_engine_platform.mmuplt.{service_name}.algo"
    headers = _build_headers(cfg, "CID=global", 30000) if service_name == 'seglabel2' \
              else _build_headers(cfg, "process_1", 60000)
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
                    f"sdu={cfg['sdu']} servicekey={_mask_secret(cfg['servicekey'])} "
                    f"token={_mask_secret(cfg['space_token'])}"
                )
            resp = requests.post(url, data=json.dumps(request_data), headers=headers, timeout=60)
            if resp.status_code != 200:
                gateway_msg = resp.headers.get("sgw-errmsg")
                _log(
                    f"  [API] {service_name} HTTP {resp.status_code}"
                    f"{f' sgw-errmsg={gateway_msg}' if gateway_msg else ''}: "
                    f"{resp.text[:300]}"
                )
                # #region agent log
                _dbg("seg_product.py:call_service:http_error", "http_non_200", {"service": service_name, "status": resp.status_code, "attempt": attempt, "sgw_errmsg": gateway_msg or "", "body": resp.text[:200]}, hid="H-L")
                # #endregion
                if resp.status_code == 401:
                    _log("  [API] 401 → 需要 Bearer token（办公网必须经 proxy 域名 + user space token）")
                elif resp.status_code == 403:
                    _log("  [API] 403 → 网关拒绝：检查域名(proxy?)、token 是否过期、VPN/白名单")
                if attempt < max_retries - 1:
                    time.sleep(1)
                    continue
                return None
            rj = resp.json()
            if 'task_result' not in rj:
                _log(f"  [API] {service_name} 返回结构异常（无 task_result）: {str(rj)[:300]}")
                # #region agent log
                _dbg("seg_product.py:call_service:no_task_result", "missing_task_result", {"service": service_name, "attempt": attempt, "keys": list(rj.keys()), "snippet": str(rj)[:200]}, hid="H-L")
                # #endregion
                if attempt < max_retries - 1:
                    time.sleep(1)
                    continue
                return None
            return rj['task_result']
        except requests.exceptions.ConnectionError as e:
            _log(f"  [API] {service_name} 连接失败（请检查 VPN / 网络）: {e}")
            # #region agent log
            _dbg("seg_product.py:call_service:conn_error", "connection_error", {"service": service_name, "attempt": attempt, "error": str(e)[:200]}, hid="H-L")
            # #endregion
            if attempt < max_retries - 1:
                time.sleep(1)
                continue
            return None
        except Exception as e:
            _log(f"  [API] {service_name} 异常: {e}")
            # #region agent log
            _dbg("seg_product.py:call_service:exception", "general_exception", {"service": service_name, "attempt": attempt, "error": str(e)[:200], "type": type(e).__name__}, hid="H-L")
            # #endregion
            if attempt < max_retries - 1:
                time.sleep(1)
                continue
            return None
    return None


# ── 图像工具函数 ─────────────────────────────────────────────────────────────

def pil_to_base64(pil_image, max_size=1024):
    """
    将 PIL 图片转为 base64 字符串（RGB，白底，JPEG）。
    piseg API 期望普通照片（RGB），不能发 RGBA 透明背景。
    使用 JPEG 而非 PNG 以控制 payload 大小（避免网关 413）。
    """
    img = pil_image.convert("RGBA")
    w, h = img.size
    if w > max_size or h > max_size:
        ratio = min(max_size / w, max_size / h)
        img = img.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)
    bg = Image.new('RGB', img.size, (255, 255, 255))
    bg.paste(img, mask=img.split()[3])
    buf = BytesIO()
    bg.save(buf, format='JPEG', quality=90)
    return base64.b64encode(buf.getvalue()).decode('utf-8')


def base64_to_pil(base64_str):
    """Decode a base64-encoded image returned by the segmentation service."""
    image_data = base64.b64decode(base64_str)
    return Image.open(BytesIO(image_data)).convert("RGBA")


def _recompress_b64_to_jpeg(seg_b64, quality=85):
    """Re-encode a segment base64 (usually PNG from piseg) to JPEG to stay under
    the API gateway payload limit that causes HTTP 413."""
    img = base64_to_pil(seg_b64)
    bg = Image.new('RGB', img.size, (255, 255, 255))
    bg.paste(img, mask=img.split()[3])
    buf = BytesIO()
    bg.save(buf, format='JPEG', quality=quality)
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
        _log(f"  [API] pisegv2 call_service 返回 None")
        # #region agent log
        _dbg("seg_product.py:piseg_pil:call_none", "call_service_returned_none", {"service": "pisegv2"}, hid="H-K")
        # #endregion
        return None
    try:
        ei = json.loads(res['extra_info'])
        imgs   = ei.get('object_images', [])
        bboxes = ei.get('object_bboxes', [])
        labels = ei.get('object_labels', [])
        if not imgs:
            _log(f"  [API] pisegv2 返回 0 个分割块（图层中未检测到独立对象）")
            # #region agent log
            _dbg("seg_product.py:piseg_pil:zero_segs", "piseg_zero_segments", {"extra_info_keys": list(ei.keys())}, hid="H-K")
            # #endregion
            return None
        return imgs, bboxes, labels
    except Exception as e:
        _log(f"  [API] piseg 结果解析失败: {e}  原始: {str(res)[:200]}")
        # #region agent log
        _dbg("seg_product.py:piseg_pil:parse_error", "piseg_parse_error", {"error": str(e)}, hid="H-K")
        # #endregion
        return None


def tagging_base64(seg_b64, level3_cat, bbox, seg_label, env='liveish', log_fn=None):
    """对分割块（base64）调用打标签服务。返回标签列表或 None。"""
    seg_b64_jpeg = _recompress_b64_to_jpeg(seg_b64)
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
            "image_list": [{"image_data": seg_b64_jpeg, "extra_info": extra}],
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


def rank_segments(label_list, bbox_list, level1_cat, level3_cat,
                  img_width=1024, img_height=1024, env='liveish', log_fn=None):
    """对所有分割块一起打分排名。返回排名结果字典或 None。"""
    text_list = [
        {
            'text': json.dumps(label),
            'extra_info': json.dumps({
                'box': bbox,
                'ori_img_height': img_height,
                'ori_img_width': img_width,
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


def _load_project_dotenv():
    """Load .env from project directory (for CLI usage; app.py has its own loader)."""
    env_path = Path(__file__).parent / ".env"
    if not env_path.is_file():
        return
    for line in env_path.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" not in line:
            continue
        key, _, val = line.partition("=")
        key, val = key.strip(), val.strip()
        if key and key not in os.environ:
            os.environ[key] = val


def probe_api(log_fn=None):
    """Minimal connectivity probe for the first-stage segmentation endpoint."""
    _load_project_dotenv()

    def _log(msg):
        (log_fn or print)(msg)

    cfg = _api_config()
    _log(f"[探测] 开始探测 Shopee segmentation API 网关")
    _log(f"[探测] gateway={cfg['gateway']}")
    _log(f"[探测] token={_mask_secret(cfg['space_token'])}")
    req = {
        "biz_type": "mmu_test",
        "region": "sg2",
        "task": {"image_list": []},
    }
    res = call_service("pisegv2", req, log_fn=log_fn, max_retries=1)
    if res is None:
        _log("[探测] 结果: 接口未通过。若日志中出现 403/401，检查 proxy 域名 + token。")
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
    """Place a segmentation image onto a transparent full-canvas image.
    piseg returns two formats: cropped (size matches bbox) or full-canvas
    (size matches doc, product already positioned). Handle both."""
    doc_w, doc_h = doc_size
    seg_img = base64_to_pil(seg_b64)

    if seg_img.size[0] == doc_w and seg_img.size[1] == doc_h:
        return seg_img

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


def _validate_product_segment(seg_b64, bbox, doc_size, log_fn=None):
    """
    Validate that a segment from the API is a plausible product cutout.

    Reference patterns (20+ PSD files):
    - Products have transparency_ratio 0.01–0.70
    - Products fill 3–80% of canvas (most 10–40%)
    - Products are NOT the entire background or a solid color block
    """
    def _log(msg):
        (log_fn or print)(msg)

    try:
        seg_img = base64_to_pil(seg_b64)
        w, h = seg_img.size
        doc_w, doc_h = doc_size
        doc_area = max(doc_w * doc_h, 1)

        left, top, right, bottom = _bbox_to_xyxy(bbox)
        seg_w = max(right - left, 1)
        seg_h = max(bottom - top, 1)
        bbox_area = seg_w * seg_h
        bbox_ratio = bbox_area / doc_area

        alpha = seg_img.getchannel('A')
        alpha_hist = alpha.histogram()
        total_px = max(sum(alpha_hist), 1)
        opaque_count = sum(alpha_hist[50:])
        opaque_frac = opaque_count / total_px
        has_real_alpha = opaque_frac < 0.95

        _log(f"  [验证] seg={w}×{h} bbox=({left},{top},{right},{bottom}) "
             f"bbox_ratio={bbox_ratio:.1%} opaque={opaque_frac:.1%} has_alpha={has_real_alpha}")

        if bbox_ratio > 0.92:
            _log(f"  [验证] ✗ 分割块 bbox 覆盖 {bbox_ratio:.0%} 画布 → 不是独立商品")
            # #region agent log
            _dbg("seg_product.py:_validate:reject_bg", "rejected_as_background", {"bbox_ratio": round(bbox_ratio, 4), "opaque_frac": round(opaque_frac, 4)}, hid="H-G")
            # #endregion
            return False

        if bbox_ratio < 0.01:
            _log(f"  [验证] ✗ 太小 ({bbox_ratio:.1%} 画布)")
            # #region agent log
            _dbg("seg_product.py:_validate:reject_small", "rejected_too_small", {"bbox_ratio": round(bbox_ratio, 4)}, hid="H-A")
            # #endregion
            return False

        rgb = seg_img.convert('RGB')
        extrema = rgb.getextrema()
        r_range = extrema[0][1] - extrema[0][0]
        g_range = extrema[1][1] - extrema[1][0]
        b_range = extrema[2][1] - extrema[2][0]
        if r_range < 10 and g_range < 10 and b_range < 10:
            _log(f"  [验证] ✗ 近似纯色块 (R={r_range} G={g_range} B={b_range})")
            # #region agent log
            _dbg("seg_product.py:_validate:reject_solid", "rejected_solid_color", {"r_range": r_range, "g_range": g_range, "b_range": b_range}, hid="H-A")
            # #endregion
            return False

        alpha_arr = np.array(alpha)
        opaque_mask = alpha_arr > 50
        if opaque_mask.any() and bbox_ratio > 0.10:
            rgb_arr = np.array(rgb)
            opaque_rgb = rgb_arr[opaque_mask]
            brightness = opaque_rgb.mean(axis=1)
            white_frac = float((brightness > 235).mean())
            if white_frac > 0.80:
                _log(f"  [验证] ✗ {white_frac:.0%} 近白色像素 → 背景碎片")
                # #region agent log
                _dbg("seg_product.py:_validate:reject_white", "rejected_white_fragment", {"white_frac": round(white_frac, 3), "bbox_ratio": round(bbox_ratio, 4)}, hid="H-A")
                # #endregion
                return False

        _log(f"  [验证] ✓ 通过")
        # #region agent log
        _dbg("seg_product.py:_validate:pass", "segment_validated_ok", {"w": w, "h": h, "bbox_ratio": round(bbox_ratio, 4), "opaque_frac": round(opaque_frac, 4), "has_alpha": has_real_alpha, "color_range": [r_range, g_range, b_range]}, hid="H-A")
        # #endregion
        return True
    except Exception as e:
        _log(f"  [验证] 异常: {e}")
        # #region agent log
        _dbg("seg_product.py:_validate:exception", "validation_exception", {"error": str(e)}, hid="H-A")
        # #endregion
        return True


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


def apply_products_and_reorder(psd_path, products, renames=None, log_fn=None):
    """
    Single Photoshop session per PSD — insert product layers, rename
    existing layers, and reorder scenebg to bottom in one open/save/close.

    products: list of dicts with 'layer_index', 'png_path', 'name'
              (insert new extracted product layer above source).
    renames:  list of dicts with 'layer_index', 'name'
              (rename existing layer in-place, no extraction needed).
    """
    def _log(msg):
        (log_fn or print)(msg)

    psd_path_js = str(psd_path).replace("\\", "/").replace("'", "\\'")

    insert_blocks = []
    for p in products:
        png_js = str(p['png_path']).replace("\\", "/").replace("'", "\\'")
        name_js = p['name'].replace("\\", "\\\\").replace('"', '\\"')
        idx = p['layer_index']
        insert_blocks.append(f"""
// --- product from layer {idx} ---
var flat = [];
flattenLayers(doc, flat);
if ({idx} < flat.length && flat[{idx}].typename !== "LayerSet") {{
    var src = flat[{idx}];
    var segDoc = app.open(new File('{png_js}'));
    if (segDoc.activeLayer.isBackgroundLayer) {{
        segDoc.activeLayer.isBackgroundLayer = false;
    }}
    segDoc.activeLayer.duplicate(doc, ElementPlacement.PLACEATBEGINNING);
    segDoc.close(SaveOptions.DONOTSAVECHANGES);
    app.activeDocument = doc;
    var pLyr = doc.layers[0];
    pLyr.name = "{name_js}";
    _productSourcePairs.push({{product: pLyr, source: src}});
}}""")

    inserts_jsx = "\n".join(insert_blocks)

    rename_blocks = []
    for r in (renames or []):
        rname_js = r['name'].replace("\\", "\\\\").replace('"', '\\"')
        ridx = r['layer_index']
        rename_blocks.append(f"""
// --- rename layer {ridx} ---
var flat = [];
flattenLayers(doc, flat);
if ({ridx} < flat.length && flat[{ridx}].typename !== "LayerSet") {{
    flat[{ridx}].name = "{rname_js}";
}}""")
    renames_jsx = "\n".join(rename_blocks)

    jsx = f"""#target photoshop
app.displayDialogs = DialogModes.NO;

function flattenLayers(container, acc) {{
    for (var i = container.layers.length - 1; i >= 0; i--) {{
        var layer = container.layers[i];
        acc.push(layer);
        if (layer.typename === "LayerSet") {{
            flattenLayers(layer, acc);
        }}
    }}
}}

var psdFile = new File('{psd_path_js}');
if (!psdFile.exists) throw new Error("PSD file not found");
var doc = app.open(psdFile);

var _productSourcePairs = [];

{inserts_jsx}

{renames_jsx}

// --- reorder scenebg to bottom ---
var scenebgLayers = [];
for (var i = 0; i < doc.layers.length; i++) {{
    if (/^scenebg\\d*$/i.test(doc.layers[i].name)) {{
        scenebgLayers.push(doc.layers[i]);
    }}
}}
for (var j = 0; j < scenebgLayers.length; j++) {{
    var bottom = doc.layers[doc.layers.length - 1];
    if (bottom.isBackgroundLayer) {{
        bottom.isBackgroundLayer = false;
    }}
    if (scenebgLayers[j] !== bottom) {{
        scenebgLayers[j].move(bottom, ElementPlacement.PLACEAFTER);
    }}
}}

// --- move each product directly above its source scenebg ---
for (var k = 0; k < _productSourcePairs.length; k++) {{
    try {{
        var pair = _productSourcePairs[k];
        pair.product.move(pair.source, ElementPlacement.PLACEBEFORE);
    }} catch(e) {{}}
}}

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
        _log(f"[扣图] 已完成 {len(products)} 个 product 插入 + scenebg 规整（单次打开）")
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

            layer_img_rgba = layer_img.convert('RGBA')

            canvas = Image.new('RGBA', (doc_w, doc_h), (0, 0, 0, 0))
            bbox = layer.bbox
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

        nearly_opaque = sum(alpha_hist[200:])
        semi_transparent = sum(alpha_hist[50:200])
        content_pixels = nearly_opaque + semi_transparent
        solid_ratio = nearly_opaque / content_pixels if content_pixels > 0 else 0

        _log(f"  [检测] 图层 {layer_idx} ({layer_name}): "
             f"图层自身透明比={transparent_ratio:.1%}, "
             f"内容面积={opaque_area/canvas_area:.1%}画布, "
             f"solid_ratio={solid_ratio:.1%}, "
             f"图层尺寸={layer_img_rgba.size}")

        info.append({
            'idx': layer_idx, 'name': layer_name,
            'transparent_ratio': transparent_ratio,
            'opaque_area': opaque_area,
            'aspect': aspect,
            'rgb_hist': rgb_hist,
            'solid_ratio': solid_ratio,
        })

    # ── 路径 A：透明扣图（图层自身有 alpha） ────────────────────────────────────
    # Threshold lowered to 0.01 — reference products have 1–68% transparency;
    # some on solid backgrounds have only 1–3%.
    a_candidates = []
    for d in info:
        if not (0.01 <= d['transparent_ratio'] <= 0.88):
            continue
        if d['opaque_area'] < canvas_area * 0.03:
            _log(f"  [检测] 图层 {d['idx']} ({d['name']}): 内容面积太小 → A路径跳过")
            continue
        if d['opaque_area'] > canvas_area * 1.0:
            _log(f"  [检测] 图层 {d['idx']} ({d['name']}): 内容面积 {d['opaque_area']/canvas_area:.0%} 超出画布 → A路径跳过（疑为背景）")
            continue
        if not (0.25 <= d['aspect'] <= 4.0):
            _log(f"  [检测] 图层 {d['idx']} ({d['name']}): 宽高比 {d['aspect']:.2f} 异常 → A路径跳过")
            continue
        if d['solid_ratio'] < 0.30:
            _log(f"  [检测] 图层 {d['idx']} ({d['name']}): solid_ratio={d['solid_ratio']:.1%} (渐变/光效) → A路径跳过")
            continue
        _log(f"  [检测] 图层 {d['idx']} ({d['name']}): ✓ 路径A候选 "
             f"透明比={d['transparent_ratio']:.1%}, 宽高比={d['aspect']:.2f}, "
             f"solid_ratio={d['solid_ratio']:.1%}")
        a_candidates.append(d)

    # ── 路径 B：重复内容检测（两两直方图比较） ──────────────────────────────────
    # Require the larger layer to be substantial (>15% canvas) — two small strips
    # being similar doesn't indicate a product-scene relationship.
    b_candidates = {}
    for i in range(len(info)):
        for j in range(i + 1, len(info)):
            sim = cosine(info[i]['rgb_hist'], info[j]['rgb_hist'])
            if sim >= 0.80:
                if info[i]['opaque_area'] >= info[j]['opaque_area']:
                    larger, smaller = info[i], info[j]
                else:
                    larger, smaller = info[j], info[i]
                if larger['opaque_area'] < canvas_area * 0.15:
                    _log(f"  [检测] 图层 {info[i]['idx']}↔{info[j]['idx']} "
                         f"内容相似(sim={sim:.2f}) 但两者面积都较小 → B路径跳过")
                    continue
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
            sorted_pool = sorted(c_pool, key=lambda x: x['opaque_area'], reverse=True)
            excluded = sorted_pool[0]
            if excluded['opaque_area'] < canvas_area * 0.15:
                _log(f"  [检测] C路径排除的最大图层面积仅 {excluded['opaque_area']/canvas_area:.1%}，"
                     f"无明确场景背景 → C路径跳过全部")
                c_candidates = []
            else:
                _log(f"  [检测] 图层 {excluded['idx']} ({excluded['name']}): "
                     f"面积最大({excluded['opaque_area']/canvas_area:.1%}) → C路径排除（疑为场景背景）")
                c_candidates = sorted_pool[1:]
                for d in c_candidates:
                    _log(f"  [检测] 图层 {d['idx']} ({d['name']}): ✓ 路径C候选 "
                         f"面积={d['opaque_area']/canvas_area:.1%}")
    else:
        c_candidates = []

    # ── 选择最终结果（优先级 A > B > C），返回所有命中图层 ────────────────────
    # Returns (results_list, path_letter) — path_letter indicates detection method.
    # Only Path A = true pre-existing cutouts (transparent bg, no extraction needed).
    # Paths B/C = product candidates on opaque backgrounds (still need extraction).
    if a_candidates:
        a_candidates.sort(key=lambda d: d['opaque_area'], reverse=True)
        result = [(d['idx'], d['name']) for d in a_candidates]
        _log(f"  [检测] ✅ 路径A 命中 {len(result)} 个图层: "
             + ", ".join(f"{idx}({n})" for idx, n in result))
        return result, 'A'

    if b_candidates:
        vals = sorted(b_candidates.values(), key=lambda d: d['opaque_area'])
        result = [(d['idx'], d['name']) for d in vals]
        _log(f"  [检测] ✅ 路径B 命中 {len(result)} 个图层: "
             + ", ".join(f"{idx}({n})" for idx, n in result))
        return result, 'B'

    if c_candidates:
        c_candidates.sort(key=lambda d: d['opaque_area'], reverse=True)
        result = [(d['idx'], d['name']) for d in c_candidates]
        _log(f"  [检测] ✅ 路径C 命中 {len(result)} 个图层: "
             + ", ".join(f"{idx}({n})" for idx, n in result))
        return result, 'C'

    return None, None


def _try_composite_piseg(composite_img, ref_layer, cates, doc_size,
                         level1_cat, level3_cat, log_fn=None):
    """
    Send a composite of all scenebg layers to piseg.
    Returns a result dict if a valid product segment is found, else None.
    """
    def _log(msg):
        (log_fn or print)(msg)

    seg = piseg_pil(composite_img, cates, log_fn=log_fn)
    if seg is None:
        _log("[扣图]   composite piseg 无结果")
        # #region agent log
        _dbg("seg_product.py:composite:no_seg", "composite_piseg_returned_none", {}, hid="H-J")
        # #endregion
        return None

    seg_ims, bboxes, seg_labels = seg
    _log(f"[扣图]   composite piseg 检测到 {len(seg_ims)} 个分割块")
    # #region agent log
    _dbg("seg_product.py:composite:piseg_result", "composite_piseg_segments", {"seg_count": len(seg_ims), "bboxes": bboxes}, hid="H-J")
    # #endregion

    all_labels, all_bboxes, all_seg_meta = [], [], []
    for s_i, (s_b64, bbox, s_label) in enumerate(zip(seg_ims, bboxes, seg_labels)):
        if not _validate_product_segment(s_b64, bbox, doc_size, log_fn=log_fn):
            _log(f"[扣图]   composite 分割块 {s_i} 预验证失败")
            continue

        label = None
        for _ in range(3):
            label = tagging_base64(s_b64, level3_cat or '', bbox, s_label, log_fn=log_fn)
            if label is not None:
                break
            time.sleep(1)
        if label is None:
            _log(f"[扣图]   composite 分割块 {s_i} 打标签失败")
            continue

        all_labels.append(label)
        all_bboxes.append(bbox)
        seg_canvas = _segment_to_full_canvas(s_b64, bbox, doc_size)
        seg_metrics = _extract_local_metrics(seg_canvas, doc_size)
        all_seg_meta.append({
            'segment_base64': s_b64,
            'segment_bbox': bbox,
            'segment_label': s_label,
            'segment_metrics': seg_metrics,
        })

    if not all_seg_meta:
        _log("[扣图]   composite 无有效分割块")
        return None

    layer_idx = ref_layer['layer_index']
    best_meta = None

    if len(all_labels) > 0:
        rank_raw = rank_segments(all_labels, all_bboxes,
                                 level1_cat or '', level3_cat or '',
                                 img_width=doc_size[0], img_height=doc_size[1],
                                 log_fn=log_fn)
        if rank_raw is not None:
            try:
                rank_data = json.loads(rank_raw['label'])
                seg_ranks = rank_data['label']
                seg_scores = rank_data['score']
                # #region agent log
                _dbg("seg_product.py:composite:ranking", "composite_ranking", {"seg_ranks": seg_ranks, "seg_scores": seg_scores}, hid="H-J")
                # #endregion
                for rank_pos in range(len(seg_ranks)):
                    sidx = seg_ranks[rank_pos]
                    if sidx < len(all_seg_meta):
                        candidate = all_seg_meta[sidx]
                        if _validate_product_segment(candidate['segment_base64'],
                                                     candidate['segment_bbox'],
                                                     doc_size, log_fn=log_fn):
                            best_meta = candidate
                            break
            except Exception:
                pass

    if best_meta is None and all_seg_meta:
        best_meta = all_seg_meta[0]

    if best_meta is None:
        return None

    _log(f"[扣图]   ✅ composite 找到商品主体！")
    # #region agent log
    _dbg("seg_product.py:composite:found", "composite_product_found", {"layer_idx": layer_idx, "bbox": best_meta['segment_bbox']}, hid="H-J")
    # #endregion
    return {layer_idx: {
        'rank': 0,
        'score': 0.8,
        'method': 'api',
        'layer_path': ref_layer['layer_path'],
        'segment_base64': best_meta['segment_base64'],
        'segment_bbox': best_meta['segment_bbox'],
        'segment_label': best_meta['segment_label'],
        'source_metrics': ref_layer['metrics'],
        'segment_metrics': best_meta['segment_metrics'],
        'doc_size': doc_size,
        'split_needed': True,
        'action': 'split',
    }}


# ── Case 2：API 分割（兜底） ──────────────────────────────────────────────────

def _api_identify(layers, level1_cat, level3_cat, log_fn=None):
    """
    逐个 scenebg 图层调用 API，找到商品主体后立即返回。

    对每个 scenebg 图层：
      1. 跳过过小图层（<5% 画布，通常是装饰条/小元素）
      2. piseg → 分割出对象
      3. 预验证分割块（跳过纯色/过大/过小的）
      4. seglabel2 → 给每个有效分割块打标签
      5. pfilter2 → 排名
      6. 按排名依次验证 → 首个通过验证的分割块就是商品主体
      7. 如果这个图层没有有效分割 → 试下一个 scenebg

    所有 scenebg 都没找到时回退到本地启发式（仅改名）。
    """
    def _log(msg):
        (log_fn or print)(msg)

    local_results = _rank_local_candidates(layers, log_fn=log_fn)
    if not local_results:
        _log("[本地识别] 没有可用的 scenebg 候选图层")
        return {}

    cates = [level1_cat or '', '', '']
    found_products = {}

    first_canvas = layers[0]['canvas']
    doc_size = first_canvas.size
    canvas_area = first_canvas.width * first_canvas.height
    min_opaque_for_api = max(int(canvas_area * 0.01), 1000)

    sorted_layers = sorted(layers,
                           key=lambda it: it['metrics'].get('opaque_pixels', 0),
                           reverse=True)

    # #region agent log
    _dbg("seg_product.py:_api_identify:layer_order", "sorted_scenebg_layers", {"order": [{"idx": it['layer_index'], "path": it['layer_path'], "opaque_px": it['metrics'].get('opaque_pixels', 0)} for it in sorted_layers], "min_opaque_for_api": min_opaque_for_api, "canvas_area": canvas_area}, hid="H-I")
    # #endregion

    # Strategy A (composite) disabled — per-layer gives more reliable results
    # and avoids layer-attribution issues with composite segments.

    # ── Per-layer segmentation ──────────────────────────────────────
    _log("[扣图] ── 逐层发送到 piseg ──")
    for item in sorted_layers:
        layer_idx = item['layer_index']
        if layer_idx in found_products:
            continue
        canvas_img = item['canvas']
        doc_size = canvas_img.size
        opaque_px = item['metrics'].get('opaque_pixels', 0)

        if opaque_px <= 0:
            continue

        if opaque_px < min_opaque_for_api:
            _log(f"[扣图] 跳过图层 {layer_idx}（{item['layer_path']}）"
                 f"内容太少: {opaque_px}px < {min_opaque_for_api}px")
            # #region agent log
            _dbg("seg_product.py:_api_identify:skip_small", "layer_skipped_min_size", {"layer_idx": layer_idx, "opaque_px": opaque_px, "min_required": min_opaque_for_api, "canvas_area": canvas_area}, hid="H-D")
            # #endregion
            continue

        _log(f"[扣图] ── 检查图层 {layer_idx}（{item['layer_path']}，"
             f"opaque={opaque_px}px/{canvas_area}px={opaque_px/canvas_area:.1%}）──")

        # ── Step 1: piseg ────────────────────────────────────────────
        seg = piseg_pil(canvas_img, cates, log_fn=log_fn)
        if seg is None:
            _log(f"[扣图]   piseg 无结果，跳过")
            # #region agent log
            _dbg("seg_product.py:_api_identify:piseg_none", "piseg_returned_none", {"layer_idx": layer_idx, "opaque_px": opaque_px}, hid="H-K")
            # #endregion
            continue

        seg_ims, bboxes, seg_labels = seg
        _log(f"[扣图]   piseg 检测到 {len(seg_ims)} 个分割块")
        # #region agent log
        _dbg("seg_product.py:_api_identify:piseg_result", "piseg_returned_segments", {"layer_idx": layer_idx, "seg_count": len(seg_ims), "bboxes": bboxes, "labels": seg_labels}, hid="H-A")
        # #endregion

        # ── Step 2: pre-validate + tagging ───────────────────────────
        all_labels, all_bboxes, all_seg_meta = [], [], []
        for s_i, (s_b64, bbox, s_label) in enumerate(zip(seg_ims, bboxes, seg_labels)):
            if not _validate_product_segment(s_b64, bbox, doc_size, log_fn=log_fn):
                _log(f"[扣图]   分割块 {s_i} 预验证失败，跳过")
                continue

            _log(f"[扣图]   打标签 {s_i+1}/{len(seg_ims)}…")
            label = None
            for _ in range(3):
                label = tagging_base64(s_b64, level3_cat or '', bbox, s_label, log_fn=log_fn)
                if label is not None:
                    break
                time.sleep(1)
            if label is None:
                _log(f"[扣图]   打标签失败，跳过分割块 {s_i}")
                # #region agent log
                _dbg("seg_product.py:_api_identify:tagging_failed", "tagging_returned_none", {"layer_idx": layer_idx, "seg_idx": s_i}, hid="H-F")
                # #endregion
                continue
            # #region agent log
            _dbg("seg_product.py:_api_identify:tagging_ok", "tagging_success", {"layer_idx": layer_idx, "seg_idx": s_i, "label_sample": str(label)[:200]}, hid="H-F")
            # #endregion

            all_labels.append(label)
            all_bboxes.append(bbox)
            seg_canvas = _segment_to_full_canvas(s_b64, bbox, doc_size)
            seg_metrics = _extract_local_metrics(seg_canvas, doc_size)
            all_seg_meta.append({
                'segment_base64': s_b64,
                'segment_bbox': bbox,
                'segment_label': s_label,
                'segment_metrics': seg_metrics,
            })

        if not all_labels:
            _log(f"[扣图]   无有效分割块，跳过此图层")
            continue

        # ── Step 3: ranking ──────────────────────────────────────────
        _log(f"[扣图]   对 {len(all_labels)} 个分割块排名…")
        rank_raw = rank_segments(all_labels, all_bboxes,
                                 level1_cat or '', level3_cat or '',
                                 img_width=doc_size[0], img_height=doc_size[1],
                                 log_fn=log_fn)
        if rank_raw is None:
            _log(f"[扣图]   排名失败，跳过此图层")
            # #region agent log
            _dbg("seg_product.py:_api_identify:ranking_failed", "ranking_returned_none", {"layer_idx": layer_idx}, hid="H-F")
            # #endregion
            continue

        try:
            rank_data  = json.loads(rank_raw['label'])
            seg_ranks  = rank_data['label']
            seg_scores = rank_data['score']
            # #region agent log
            _dbg("seg_product.py:_api_identify:ranking_ok", "ranking_success", {"layer_idx": layer_idx, "seg_ranks": seg_ranks, "seg_scores": seg_scores}, hid="H-F")
            # #endregion
        except Exception as e:
            _log(f"[扣图]   排名解析失败: {e}")
            # #region agent log
            _dbg("seg_product.py:_api_identify:ranking_parse_fail", "ranking_parse_error", {"layer_idx": layer_idx, "error": str(e), "raw": str(rank_raw)[:300]}, hid="H-F")
            # #endregion
            continue

        if not seg_ranks:
            if all_seg_meta:
                best_meta = all_seg_meta[0]
                if not _should_split_candidate(item['metrics'], best_meta['segment_metrics']):
                    _log(f"[扣图]   分割块覆盖整个图层（非真实抠图），跳过")
                    continue
                _log(f"[扣图]   排名为空但有预验证分割块，直接使用")
                # #region agent log
                _dbg("seg_product.py:_api_identify:empty_rank_fallback", "using_prevalidated_seg_despite_empty_rank", {"layer_idx": layer_idx, "bbox": best_meta['segment_bbox']}, hid="H-F2")
                # #endregion
                found_products[layer_idx] = {
                    'rank': 0,
                    'score': 0.5,
                    'method': 'api',
                    'layer_path': item['layer_path'],
                    'segment_base64': best_meta['segment_base64'],
                    'segment_bbox': best_meta['segment_bbox'],
                    'segment_label': best_meta['segment_label'],
                    'source_metrics': item['metrics'],
                    'segment_metrics': best_meta['segment_metrics'],
                    'doc_size': doc_size,
                    'split_needed': True,
                    'action': 'split',
                }
            continue

        # ── Step 4: iterate ranked segments, pick first valid ────────
        for rank_pos in range(len(seg_ranks)):
            seg_idx = seg_ranks[rank_pos]
            score = seg_scores[rank_pos] if rank_pos < len(seg_scores) else 0.0

            if seg_idx >= len(all_seg_meta):
                continue

            best_meta = all_seg_meta[seg_idx]

            _log(f"[扣图]   rank {rank_pos}: seg_idx={seg_idx} score={score:.4f}")

            if not _validate_product_segment(best_meta['segment_base64'],
                                             best_meta['segment_bbox'],
                                             doc_size, log_fn=log_fn):
                _log(f"[扣图]   rank {rank_pos} 后验证失败，跳过")
                continue

            if not _should_split_candidate(item['metrics'], best_meta['segment_metrics']):
                _log(f"[扣图]   rank {rank_pos} 分割块覆盖整个图层，跳过")
                continue

            _log(f"[扣图]   ✅ 找到商品主体！rank={rank_pos} score={score:.4f}")

            found_products[layer_idx] = {
                'rank': 0,
                'score': score,
                'method': 'api',
                'layer_path': item['layer_path'],
                'segment_base64': best_meta['segment_base64'],
                'segment_bbox': best_meta['segment_bbox'],
                'segment_label': best_meta['segment_label'],
                'source_metrics': item['metrics'],
                'segment_metrics': best_meta['segment_metrics'],
                'doc_size': doc_size,
                'split_needed': True,
                'action': 'split',
            }
            break

        _log(f"[扣图]   该图层所有分割块均未通过验证")

    if found_products:
        _log(f"[扣图] 共找到 {len(found_products)} 个商品主体")
        return found_products

    _log("[扣图] 所有 scenebg 图层均未找到商品主体，回退到本地启发式结果")
    return local_results


def identify_product_layer(psd_path, level1_cat, level3_cat, log_fn=None):
    """
    主函数：识别 PSD 中哪个 'scenebg' 图层是商品主体并获取分割数据。

    策略：
      1. 本地透明度检测 — 标记哪些 scenebg 已经是人工抠图（不需要再抠）。
      2. API pipeline（piseg → seglabel2 → pfilter2）— 确认哪些图层包含商品主体。
         • API 确认含商品 + 已有抠图 → action='rename'（直接改名）
         • API 确认含商品 + 非抠图   → action='split'（需要抠图）
      3. 如果 API 不可用，回退到本地检测结果。

    返回:
        { layer_index: {'rank': int, 'score': float, 'method': str, ...}, ... }
    """
    def _log(msg):
        (log_fn or print)(msg)

    _log(f"[扣图] 开始识别: {Path(psd_path).name}")
    _log(f"[扣图] 品类: {level1_cat or '(未填)'} / {level3_cat or '(未填)'}")

    layers = export_scenebg_layers(psd_path, log_fn=log_fn)
    if not layers:
        _log("[扣图] 未找到 scenebg 图层")
        return {}

    # ── 1. 本地检测：标记已有人工抠图的图层 ──────────────────────────
    _log("[扣图] 检测是否存在已抠好的商品主体…")
    cutouts, cutout_path = detect_existing_cutout(layers, log_fn=log_fn)
    # Only Path A (transparent background) = true pre-existing cutouts
    # that don't need re-extraction. Paths B/C are product candidates
    # on opaque backgrounds that still need API extraction.
    cutout_indices = set(idx for idx, _ in cutouts) if cutouts and cutout_path == 'A' else set()
    if cutout_indices:
        _log(f"[扣图] 本地检测到 {len(cutout_indices)} 个已有人工抠图(路径A): {cutout_indices}")
    elif cutouts:
        _log(f"[扣图] 本地检测到 {len(cutouts)} 个商品候选(路径{cutout_path})，仍需 API 确认")

    # ── 2. API pipeline — 确认哪些图层真正包含商品主体 ────────────────
    _log("[扣图] 调用 API pipeline（piseg → seglabel2 → pfilter2）…")
    api_results = _api_identify(layers, level1_cat, level3_cat, log_fn=log_fn)

    has_api_product = any(v.get('action') == 'split' for v in api_results.values())
    if has_api_product:
        for idx, info in api_results.items():
            if info.get('action') == 'split' and idx in cutout_indices:
                _log(f"[扣图] 图层 {idx}: API 确认含商品主体 + 已有人工抠图 → 直接改名")
                info['action'] = 'rename'
                info['split_needed'] = False
                info.pop('segment_base64', None)
                info.pop('segment_bbox', None)
        _log("[扣图] API 成功找到商品主体")
        return api_results

    # ── 3. API 不可用时回退：仅 Path A（透明已抠图）可直接改名 ─────
    if cutouts and cutout_path == 'A':
        _log(f"[扣图] API 未返回有效分割，使用本地路径A（透明背景已抠好的商品）")
        results = {}
        for rank, (idx, name) in enumerate(cutouts):
            _log(f"[扣图] ✅ 已有透明抠图: 图层 {idx} ({name}) → 直接改名")
            results[idx] = {
                'rank': rank,
                'score': 1.0,
                'method': 'local_A',
                'layer_path': name,
                'split_needed': False,
                'action': 'rename',
            }
        return results
    elif cutouts:
        _log(f"[扣图] API 未返回有效分割，本地路径{cutout_path}候选不具备透明抠图 → 不做改名")

    _log("[扣图] 未能识别商品主体（API 未找到可抠图的商品，本地无透明抠图）")
    return {}


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

        split_layers = {k: v for k, v in results.items()
                        if v.get('action') == 'split' and v.get('segment_base64')}
        rename_layers = {k: v for k, v in results.items() if k not in split_layers}

        best_idx = min(results, key=lambda k: results[k]['rank'])
        best = results[best_idx]
        method = best.get('method', 'unknown')
        summary[psd_path.name] = {
            'layer_index': best_idx,
            'rank': best['rank'],
            'score': best['score'],
            'action': 'split' if split_layers else best.get('action', 'rename'),
            'method': method,
            'product_count': len(split_layers),
        }

        _log(f"[扣图] {psd_path.name}: 找到 {len(split_layers)} 个需抠图, "
             f"{len(rename_layers)} 个仅改名")

        if auto_rename:
            tmp_paths = []
            total_products = len(split_layers) + len(rename_layers)
            use_numbering = total_products > 1

            all_indices = sorted(set(split_layers.keys()) | set(rename_layers.keys()))
            products = []
            renames = []
            for seq, idx in enumerate(all_indices, start=1):
                name = f'product{seq}' if use_numbering else 'product'
                if idx in split_layers:
                    info = split_layers[idx]
                    try:
                        tmp_png = materialize_segment_png(
                            info['segment_base64'],
                            info['segment_bbox'],
                            info['doc_size'],
                            log_fn=log_fn
                        )
                        products.append({
                            'layer_index': idx,
                            'png_path': str(tmp_png),
                            'name': name,
                        })
                        tmp_paths.append(tmp_png)
                    except Exception as e:
                        _log(f"[扣图] 图层 {idx} PNG 生成失败: {e}")
                else:
                    renames.append({
                        'layer_index': idx,
                        'name': name,
                    })
                    _log(f"[扣图] 图层 {idx} → '{name}'（已有抠图，仅改名）")

            products.sort(key=lambda p: p['layer_index'], reverse=True)

            if products or renames:
                ok = apply_products_and_reorder(
                    str(psd_path), products, renames=renames, log_fn=log_fn)
                if not ok:
                    _log(f"[扣图] {psd_path.name}: Photoshop 处理失败")

            for p in tmp_paths:
                Path(p).unlink(missing_ok=True)

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
