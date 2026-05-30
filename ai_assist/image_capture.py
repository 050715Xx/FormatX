"""图片采集工具：剪贴板读取 + 截图 → base64"""
import base64
import io


def get_image_from_clipboard() -> bytes | None:
    """从剪贴板读取图片，返回 PNG 字节。不支持则返回 None。"""
    try:
        from PIL import ImageGrab, Image
        img = ImageGrab.grabclipboard()
        if img is None:
            return None
        if isinstance(img, Image.Image):
            buf = io.BytesIO()
            img.convert("RGB").save(buf, format="PNG")
            return buf.getvalue()
        return None
    except ImportError:
        return None


def screenshot_to_base64() -> str | None:
    """截取全屏并返回 base64 编码。"""
    try:
        from PIL import ImageGrab
        img = ImageGrab.grab()
        buf = io.BytesIO()
        img.convert("RGB").save(buf, format="PNG")
        return base64.b64encode(buf.getvalue()).decode("utf-8")
    except ImportError:
        return None


def image_to_base64(image_bytes: bytes) -> str:
    return base64.b64encode(image_bytes).decode("utf-8")
