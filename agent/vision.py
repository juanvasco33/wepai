import subprocess, base64, os, time
from PIL import Image

SCREENSHOT_PATH = os.path.expanduser("~/.wepai/preview.png")
CROPPED_PATH    = os.path.expanduser("~/.wepai/preview_cropped.png")

os.makedirs(os.path.expanduser("~/.wepai"), exist_ok=True)


def get_office_window_id(app_name: str) -> str | None:
    script = f'tell application "System Events" to tell process "{app_name}" to get id of window 1'
    result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    if result.returncode == 0:
        return result.stdout.strip()
    return None


def capture_office_window(app_name: str) -> str | None:
    time.sleep(1.5)
    win_id = get_office_window_id(app_name)
    if win_id:
        result = subprocess.run(
            ["screencapture", "-l", win_id, "-x", SCREENSHOT_PATH],
            capture_output=True
        )
    else:
        result = subprocess.run(
            ["screencapture", "-x", SCREENSHOT_PATH],
            capture_output=True
        )

    if result.returncode != 0 or not os.path.exists(SCREENSHOT_PATH):
        return None

    try:
        img = Image.open(SCREENSHOT_PATH)
        width, height = img.size
        crop_box = (0, 0, width, min(height, int(height * 0.88)))
        img_cropped = img.crop(crop_box)
        max_width = 1200
        if img_cropped.width > max_width:
            ratio = max_width / img_cropped.width
            new_size = (max_width, int(img_cropped.height * ratio))
            img_cropped = img_cropped.resize(new_size, Image.LANCZOS)
        img_cropped.save(CROPPED_PATH, "PNG", optimize=True)
        with open(CROPPED_PATH, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")
    except Exception as e:
        print(f"Screenshot error: {e}")
        return None


def get_preview_image() -> Image.Image | None:
    path = CROPPED_PATH if os.path.exists(CROPPED_PATH) else SCREENSHOT_PATH
    if os.path.exists(path):
        try:
            return Image.open(path).copy()
        except Exception:
            return None
    return None


APP_NAMES = {
    "word":        "Microsoft Word",
    "excel":       "Microsoft Excel",
    "powerpoint":  "Microsoft PowerPoint",
}
