import pathlib
import sys
import time
import traceback
import typing
import hashlib

from datetime import datetime
from PIL import Image, ImageGrab, ImageChops
import win32clipboard
import win32con
import io

import logging
logger = logging.getLogger(__name__)
BORDER_SIZE = 30
POLL_INTERVAL = 1  # seconds


def crop_to_object(image: Image.Image, border: int = 40) -> Image.Image:
    bg_color = image.getpixel((0, 0))  # assume top-left is background
    bg_image = Image.new(image.mode, image.size, bg_color)
    diff = ImageChops.difference(image, bg_image).convert("L")
    diff = diff.point(lambda x: 255 if x > 10 else 0)

    bbox = diff.getbbox()
    if not bbox:
        logger.warning("No object detected.")
        return image

    left = max(bbox[0] - border, 0)
    upper = max(bbox[1] - border, 0)
    right = min(bbox[2] + border, image.width)
    lower = min(bbox[3] + border, image.height)

    return image.crop((left, upper, right, lower))


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


def monitor_clipboard():
    last_hash = None
    logger.info("Monitoring clipboard for image changes...")

    while True:
        try:
            img = ImageGrab.grabclipboard()
            if isinstance(img, Image.Image):
                current_hash = image_hash(img)
                if current_hash != last_hash:
                    logger.info("New image detected in clipboard.")
                    cropped = crop_to_object(img, BORDER_SIZE)
                    image_to_clipboard(cropped)
                    last_hash = current_hash
        except Exception as e:
            logger.warning(f"Clipboard polling error: {e}")

        time.sleep(POLL_INTERVAL)


def main() -> None:
    start_time = time.perf_counter()
    logger.info("Starting clipboard watcher...")
    monitor_clipboard()
    end_time = time.perf_counter()
    duration = end_time - start_time
    logger.info(f"Completed operation in {duration:.4f}s.")


def setup_logging(
        logger: logging.Logger,
        log_file_path: typing.Union[str, pathlib.Path],
        console_logging_level: int = logging.DEBUG,
        file_logging_level: int = logging.DEBUG,
        log_message_format: str = "%(asctime)s.%(msecs)03d %(levelname)s [%(funcName)s] [%(name)s]: %(message)s",
        date_format: str = "%Y-%m-%d %H:%M:%S") -> None:
    logger.setLevel(file_logging_level)  # Set the overall logging level

    # File Handler for script-named log file (overwrite each run)
    file_handler = logging.FileHandler(log_file_path, encoding="utf-8", mode="w")
    file_handler.setLevel(file_logging_level)
    file_handler.setFormatter(logging.Formatter(log_message_format, datefmt=date_format))
    logger.addHandler(file_handler)

    # Console Handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(console_logging_level)
    console_handler.setFormatter(logging.Formatter(log_message_format, datefmt=date_format))
    logger.addHandler(console_handler)

    # Set specific logging levels if needed
    # logging.getLogger("requests").setLevel(logging.INFO)


if __name__ == "__main__":
    script_name = pathlib.Path(__file__).stem
    log_file_name = f"{script_name}.log"
    log_file_path = pathlib.Path(log_file_name)
    setup_logging(logger, log_file_path, log_message_format="%(asctime)s.%(msecs)03d %(levelname)s [%(funcName)s]: %(message)s")

    error = 0
    try:
        main()
    except Exception as e:
        logger.warning(f"A fatal error has occurred: {repr(e)}\n{traceback.format_exc()}")
        error = 1
    finally:
        sys.exit(error)
