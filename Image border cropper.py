import pathlib
import sys
import time
import traceback
import typing

from datetime import datetime
from PIL import Image, ImageGrab, ImageChops

import logging
logger = logging.getLogger(__name__)

BORDER_SIZE = 40


def crop_to_object(image: Image.Image, border: int = 40) -> Image.Image:
    bg_color = image.getpixel((0, 0))  # Assume top-left is background
    bg_image = Image.new(image.mode, image.size, bg_color)
    diff = ImageChops.difference(image, bg_image).convert("L")

    # Enhance contrast slightly for better bounding box detection
    diff = diff.point(lambda x: 255 if x > 10 else 0)

    bbox = diff.getbbox()
    if not bbox:
        logger.warning("No object detected in image.")
        return image

    left = max(bbox[0] - border, 0)
    upper = max(bbox[1] - border, 0)
    right = min(bbox[2] + border, image.width)
    lower = min(bbox[3] + border, image.height)

    cropped = image.crop((left, upper, right, lower))
    logger.info(f"Cropped image to box: ({left}, {upper}, {right}, {lower})")
    return cropped


def save_image(image: Image.Image, out_path: pathlib.Path) -> None:
    image.save(out_path)
    logger.info(f"Saved cropped image to: {out_path}")


def main() -> None:
    start_time = time.perf_counter()
    logger.info("Starting operation...")

    try:
        image = ImageGrab.grabclipboard()
        if not isinstance(image, Image.Image):
            logger.error("No image found in clipboard.")
            return

        logger.info(f"Clipboard image size: {image.size}")
        cropped = crop_to_object(image, BORDER_SIZE)

        out_path = pathlib.Path("cropped_output.png")
        save_image(cropped, out_path)

    except Exception as e:
        logger.exception(f"Error during processing: {e}")

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
