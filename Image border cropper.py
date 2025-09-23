import ctypes
import hashlib
import io
import logging
import pathlib
import socket
import sys
import time
import traceback
import typing
import win32clipboard
import win32con
import win32gui
from datetime import datetime
from PIL import Image, ImageGrab, ImageChops

logger = logging.getLogger(__name__)

BORDER_SIZE = 10
WM_CLIPBOARDUPDATE = 0x031D
last_hash = None
ignore_next = False


def get_background_color(image: Image.Image):
    """Guess background color from image corners."""
    corners = [
        image.getpixel((0, 0)),
        image.getpixel((image.width - 1, 0)),
        image.getpixel((0, image.height - 1)),
        image.getpixel((image.width - 1, image.height - 1)),
    ]
    return max(set(corners), key=corners.count)


def crop_to_object(image: Image.Image, border: int = 40) -> Image.Image:
    bg_color = get_background_color(image)
    bg_image = Image.new(image.mode, image.size, bg_color)
    diff = ImageChops.difference(image, bg_image).convert("L")
    diff = diff.point(lambda x: 255 if x > 10 else 0)  # type: ignore[arg-type]

    bbox = diff.getbbox()
    if not bbox:
        logger.warning("No object detected.")
        return image

    # Desired crop with border
    left = bbox[0] - border
    upper = bbox[1] - border
    right = bbox[2] + border
    lower = bbox[3] + border

    # Compute needed padding if crop exceeds original image
    pad_left = max(0, -left)
    pad_top = max(0, -upper)
    pad_right = max(0, right - image.width)
    pad_bottom = max(0, lower - image.height)

    # Crop and expand if needed
    cropped = image.crop((
        max(left, 0),
        max(upper, 0),
        min(right, image.width),
        min(lower, image.height),
    ))

    if any((pad_left, pad_top, pad_right, pad_bottom)):
        new_width = cropped.width + pad_left + pad_right
        new_height = cropped.height + pad_top + pad_bottom
        new_img = Image.new(image.mode, (new_width, new_height), bg_color)
        new_img.paste(cropped, (pad_left, pad_top))
        return new_img

    return cropped


def image_to_clipboard(image: Image.Image) -> None:
    """Copy an image to the Windows clipboard in DIB format."""
    output = io.BytesIO()
    image.convert("RGB").save(output, format="BMP")
    data = output.getvalue()[14:]  # Strip BMP header
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32con.CF_DIB, data)
    win32clipboard.CloseClipboard()
    logger.info("Updated clipboard with cropped image.")


def image_hash(image: Image.Image) -> str:
    """Generate a hash of image content for change detection."""
    with io.BytesIO() as f:
        image.save(f, format='PNG')
        return hashlib.sha256(f.getvalue()).hexdigest()


def on_clipboard_update(hwnd, msg, wparam, lparam):
    global last_hash, ignore_next
    if msg == WM_CLIPBOARDUPDATE:
        try:
            if ignore_next:
                ignore_next = False
                return 0

            img = ImageGrab.grabclipboard()
            if isinstance(img, Image.Image):
                current_hash = image_hash(img)
                if current_hash == last_hash:
                    return 0  # same image, ignore

                logger.info("New image detected in clipboard.")
                cropped = crop_to_object(img, BORDER_SIZE)
                image_to_clipboard(cropped)
                last_hash = image_hash(cropped)
                ignore_next = True
        except Exception as e:
            logger.warning(f"Clipboard processing error: {e}")
    return 0


def start_clipboard_listener():
    wc = typing.cast(typing.Any, win32gui.WNDCLASS())
    wc.lpfnWndProc = on_clipboard_update
    wc.lpszClassName = "ClipboardWatcher"
    hinst = win32gui.GetModuleHandle(None)
    wc.hInstance = hinst
    classAtom = win32gui.RegisterClass(wc)

    hwnd = win32gui.CreateWindow(
        classAtom,
        "ClipboardWatcher",
        0, 0, 0, 0, 0,
        0, 0,
        hinst,
        None,
    )

    user32 = ctypes.windll.user32
    if not user32.AddClipboardFormatListener(hwnd):
        raise ctypes.WinError()

    logger.info("Started clipboard listener (event-driven).")
    win32gui.PumpMessages()


def setup_logging(
        logger: logging.Logger,
        log_file_path: typing.Union[str, pathlib.Path],
        number_of_logs_to_keep: typing.Union[int, None] = None,
        console_logging_level: int = logging.DEBUG,
        file_logging_level: int = logging.DEBUG,
        log_message_format: str = "%(asctime)s.%(msecs)03d %(levelname)s [%(funcName)s] [%(name)s]: %(message)s",
        date_format: str = "%Y-%m-%d %H:%M:%S") -> None:
    # Ensure log_dir is a Path object
    log_file_path = pathlib.Path(log_file_path)
    log_dir = log_file_path.parent
    log_dir.mkdir(parents=True, exist_ok=True)  # Create logs dir if it does not exist

    # Limit # of logs in logs folder
    if number_of_logs_to_keep is not None:
        log_files = sorted([f for f in log_dir.glob("*.log")], key=lambda f: f.stat().st_mtime)
        if len(log_files) >= number_of_logs_to_keep:
            for file in log_files[:len(log_files) - number_of_logs_to_keep + 1]:
                file.unlink()

    logger.setLevel(file_logging_level)  # Set the overall logging level

    # File Handler for date-based log file
    file_handler_date = logging.FileHandler(log_file_path, encoding="utf-8")
    file_handler_date.setLevel(file_logging_level)
    file_handler_date.setFormatter(logging.Formatter(log_message_format, datefmt=date_format))
    logger.addHandler(file_handler_date)

    # Console Handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(console_logging_level)
    console_handler.setFormatter(logging.Formatter(log_message_format, datefmt=date_format))
    logger.addHandler(console_handler)

    # Set specific logging levels if needed
    # logging.getLogger("requests").setLevel(logging.INFO)


if __name__ == "__main__":
    pc_name = socket.gethostname()
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    script_name = pathlib.Path(__file__).stem
    log_dir = pathlib.Path(f"{script_name} Logs")
    log_file_name = f"{timestamp}_{pc_name}.log"
    log_file_path = log_dir / log_file_name
    setup_logging(logger, log_file_path, number_of_logs_to_keep=10, log_message_format="%(asctime)s.%(msecs)03d %(levelname)s [%(funcName)s]: %(message)s")

    error = 0
    try:
        start_time = time.perf_counter()
        logger.info("Starting operation...")
        start_clipboard_listener()
        end_time = time.perf_counter()
        duration = end_time - start_time
        logger.info(f"Completed operation in {duration:.4f}s.")
    except Exception as e:
        logger.warning(f"A fatal error has occurred: {repr(e)}\n{traceback.format_exc()}")
        error = 1
    finally:
        sys.exit(error)
