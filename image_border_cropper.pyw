"""
Image Border Cropper

A Windows utility script that monitors the clipboard for new images, automatically crops out uniform borders,
and replaces the clipboard image with the cropped version. The script runs in the background with a system tray icon,
providing quick access to the script folder and an exit option.

Features:
- Monitors clipboard for new images (ignores text and duplicate images).
- Crops images to remove uniform borders, preserving a configurable border size.
- Updates the clipboard with the cropped image.
- Runs as a background process with a system tray icon for user interaction.
- Configurable via a TOML file for logging, border size, and exit behavior.

How to use:
1. Place a configuration TOML file named `{script_name}_config.toml` in the same directory as this script.
2. Run the script. It will appear as a tray icon.
3. Copy an image to the clipboard (e.g., using Print Screen or Snipping Tool).
4. The script will automatically crop the image and update the clipboard.
5. Use the tray icon to open the script folder or exit the application.
"""

import datetime
import hashlib
import io
import json
import logging
import os
import pyperclip
import send2trash
import socket
import sys
import threading
import time
import tomllib
import webbrowser
import win32clipboard
from datetime import datetime
from pathlib import Path
from PIL import Image, ImageGrab, ImageChops, ImageOps
from pystray import Icon, MenuItem, Menu

logger = logging.getLogger(__name__)

__version__ = "2.1.0"  # Major.Minor.Patch

exit_event = threading.Event()
running_event = threading.Event()
running_event.set()

CONFIG = {}


def load_image(path: str | Path) -> Image.Image:
    path = Path(path)
    image = Image.open(path)
    logger.debug(f"Loaded image at path {json.dumps(str(path))}")
    return image


def open_source_url():
    webbrowser.open("https://github.com/RandomGgames/image_border_cropper.py")
    logger.debug("Opened source URL.")


def open_issues_url():
    webbrowser.open("https://github.com/RandomGgames/image_border_cropper.py/issues")
    logger.debug("Opened issues URL.")


def open_script_folder():
    folder_path = os.path.dirname(os.path.abspath(__file__))
    os.startfile(folder_path)
    logger.debug(f"Opened script folder: {json.dumps(str(folder_path))}")


def toggle_pause(icon):
    """Toggle pause state."""
    if running_event.is_set():
        running_event.clear()
        logger.info("Paused clipboard monitor.")
    else:
        running_event.set()
        logger.info("Resumed clipboard monitor.")

    icon.update_menu()


def pause_checked(_):
    """Return True when paused."""
    return not running_event.is_set()


def on_exit(icon):
    logger.debug("Exit pressed on system tray icon")
    icon.stop()
    logger.debug("System tray icon stopped.")
    exit_event.set()
    logger.debug("Exit event triggered")


def startup_tray_icon():
    logger.debug("Starting up system tray icon...")

    image = load_image("system_tray_icon.png")

    menu = Menu(
        MenuItem(f"Image Border Cropper v{__version__}", None, enabled=False),
        Menu.SEPARATOR,
        MenuItem("Pause Clipboard Monitor", toggle_pause, checked=pause_checked),
        Menu.SEPARATOR,
        MenuItem("Source Page", open_source_url),
        MenuItem("Issues Page", open_issues_url),
        MenuItem("Open File Path", open_script_folder),
        Menu.SEPARATOR,
        MenuItem("Exit", on_exit),
    )

    icon = Icon("Clipboard Whitespace Trimmer", image, menu=menu)

    logger.debug("Started system tray icon.")
    icon.run()


def trim_image_borders(img, border_width):
    """Shaves off pixels from all sides using ImageOps."""
    # This removes the border_width from Left, Top, Right, and Bottom
    return ImageOps.crop(img, border=border_width)


def send_image_to_clipboard(img):
    """Converts PIL image to Windows DIB format and updates clipboard."""
    output = io.BytesIO()
    img.convert("RGB").save(output, "BMP")  # Convert to RGB to ensure BMP compatibility
    data = output.getvalue()[14:]  # Remove the 14-byte BMP file header to satisfy Windows CF_DIB format
    output.close()

    win32clipboard.OpenClipboard()  # pylint: disable=c-extension-no-member
    try:
        win32clipboard.EmptyClipboard()  # pylint: disable=c-extension-no-member
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)  # pylint: disable=c-extension-no-member
    finally:
        win32clipboard.CloseClipboard()  # pylint: disable=c-extension-no-member


def trim_to_content(img, padding=10, tolerance=30):
    """
    Automatically crops the image to the object in the center.
    - tolerance: 0 to 255 (how different a pixel must be from the corner to be 'content')
    - padding: how many pixels of the original background to keep around the object
    """
    img = img.convert("RGBA")
    bg_color = img.getpixel((0, 0))
    bg = Image.new("RGBA", img.size, bg_color)
    diff = ImageChops.difference(img, bg)
    diff = diff.convert("L")
    lut = [255 if i > tolerance else 0 for i in range(256)]
    mask = diff.point(lut)
    bbox = mask.getbbox()
    if not bbox:
        return img  # Return original if no object found (image is all one color)
    left, top, right, bottom = bbox
    left = max(0, left - padding)
    top = max(0, top - padding)
    right = min(img.width, right + padding)
    bottom = min(img.height, bottom + padding)
    return img.crop((left, top, right, bottom))


def trim_and_expand_border_to_content(img, padding=10, tolerance=30):
    """
    Standardizes the image border.
    If the border is too big, it trims it.
    If the border is too small, it expands it.
    """
    img = img.convert("RGBA")
    bg_color = img.getpixel((0, 0))
    bg = Image.new("RGBA", img.size, bg_color)
    diff = ImageChops.difference(img, bg)
    diff = diff.convert("L")
    lut = [255 if i > tolerance else 0 for i in range(256)]
    mask = diff.point(lut)
    bbox = mask.getbbox()
    if not bbox:
        return img  # Return original if the image is a solid color
    object_only = img.crop(bbox)
    obj_w, obj_h = object_only.size
    new_width = obj_w + (padding * 2)
    new_height = obj_h + (padding * 2)
    standardized_img = Image.new("RGBA", (new_width, new_height), bg_color)
    standardized_img.paste(object_only, (padding, padding), object_only)
    return standardized_img.convert("RGB")


def get_image_hash(img):
    """Generates a stable hash based on pixel data only."""
    return hashlib.md5(img.convert("RGB").tobytes()).hexdigest()


def main():
    system_tray_thread = threading.Thread(target=startup_tray_icon, daemon=True)
    system_tray_thread.start()

    PADDING = CONFIG.get("padding", 10)
    TOLERANCE = CONFIG.get("tolerance", 30)

    logger.debug(f"Config Loaded - Padding: {PADDING}, Tolerance: {TOLERANCE}")

    last_image_hash = None

    while not exit_event.is_set():
        running_event.wait()

        try:
            img = ImageGrab.grabclipboard()

            if isinstance(img, Image.Image):
                current_hash = get_image_hash(img)

                if current_hash != last_image_hash:
                    logger.debug("New image detected. Analyzing content...")
                    processed_img = trim_and_expand_border_to_content(img, padding=PADDING, tolerance=TOLERANCE)
                    processed_hash = get_image_hash(processed_img)

                    if processed_hash != current_hash:
                        logger.debug("Updating clipboard...")
                        send_image_to_clipboard(processed_img)
                        last_image_hash = processed_hash
                        logger.info(f"Trimmed to content. New Hash: {last_image_hash}")
                    else:
                        last_image_hash = current_hash
                        logger.debug("No borders detected. Skipping clipboard update.")

            elif pyperclip.paste() == "":
                last_image_hash = None

        except Exception as e:
            logger.error(f"Main loop error: {e}")

        time.sleep(1)

    logger.info("Exit event detected. Shutting down main loop.")


def enforce_max_log_count(dir_path: Path | str, max_count: int | None, script_name: str) -> None:
    """
    Keep only the N most recent logs for this script.

    Args:
        dir_path (Path | str): The directory path to the log files.
        max_count (int | None): The maximum number of log files to keep. None for no limit.
        script_name (str): The name of the script to filter logs by.
    """
    if max_count is None or max_count <= 0:
        return
    dir_path = Path(dir_path)
    files = sorted([f for f in dir_path.glob(f"*{script_name}*.log") if f.is_file()])  # Newest will be at the end of the list
    if len(files) > max_count:
        to_delete = files[:-max_count]  # Everything except the last N files
        for f in to_delete:
            try:
                send2trash.send2trash(f)
                logger.debug(f"Deleted old log: {f.name}")
            except OSError as e:
                logger.error(f"Failed to delete {f.name}: {e}")


def setup_logging(
        logger_obj: logging.Logger,
        file_path: Path | str,
        script_name: str,
        max_log_files: int | None = None,
        console_logging_level: int = logging.DEBUG,
        file_logging_level: int = logging.DEBUG,
        message_format: str = "%(asctime)s.%(msecs)03d %(levelname)s [%(funcName)s]: %(message)s",
        date_format: str = "%Y-%m-%d %H:%M:%S"
) -> None:
    """
    Set up logging for a script.

    Args:
    logger_obj (logging.Logger): The logger object to configure.
    file_path (Path | str): The file path of the log file to write.
    max_log_files (int | None, optional): The maximum total size for all logs in the folder. Defaults to None.
    console_logging_level (int, optional): The logging level for console output. Defaults to logging.DEBUG.
    file_logging_level (int, optional): The logging level for file output. Defaults to logging.DEBUG.
    message_format (str, optional): The format string for log messages. Defaults to "%(asctime)s.%(msecs)03d %(levelname)s [%(funcName)s]: %(message)s".
    date_format (str, optional): The format string for log timestamps. Defaults to "%Y-%m-%d %H:%M:%S".
    """

    file_path = Path(file_path)
    dir_path = file_path.parent
    dir_path.mkdir(parents=True, exist_ok=True)

    logger_obj.handlers.clear()
    logger_obj.setLevel(file_logging_level)

    formatter = logging.Formatter(message_format, datefmt=date_format)

    # File Handler
    file_handler = logging.FileHandler(file_path, encoding="utf-8")
    file_handler.setLevel(file_logging_level)
    file_handler.setFormatter(formatter)
    logger_obj.addHandler(file_handler)

    # Console Handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(console_logging_level)
    console_handler.setFormatter(formatter)
    logger_obj.addHandler(console_handler)

    if max_log_files is not None:
        enforce_max_log_count(dir_path, max_log_files, script_name)


def read_toml(file_path: Path | str) -> dict:
    """
    Reads a TOML file and returns its contents as a dictionary.

    Args:
        file_path (Path | str): The file path of the TOML file to read.

    Returns:
        dict: The contents of the TOML file as a dictionary.

    Raises:
        FileNotFoundError: If the TOML file does not exist.
        OSError: If the file cannot be read.
        tomllib.TOMLDecodeError (or toml.TomlDecodeError): If the file is invalid TOML.
    """
    path = Path(file_path)

    if not path.is_file():
        raise FileNotFoundError(f"File not found: {json.dumps(str(path))}")

    try:
        # Read TOML as bytes
        with path.open("rb") as f:
            data = tomllib.load(f)  # Replace with 'toml.load(f)' if using the toml package
        return data

    except (OSError, tomllib.TOMLDecodeError):
        logger.exception(f"Failed to read TOML file: {json.dumps(str(file_path))}")
        raise


def load_config(file_path: Path | str) -> dict:
    """
    Load configuration from a TOML file.

    Args:
    file_path (Path | str): The file path of the TOML file to read.

    Returns:
    dict: The contents of the TOML file as a dictionary.
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {json.dumps(str(file_path))}")
    data = read_toml(file_path)
    return data


def bootstrap():
    """
    Handles environment setup, configuration loading,
    and logging before executing the main script logic.
    """
    exit_code = 0
    try:
        script_path = Path(__file__)
        script_name = script_path.stem
        pc_name = socket.gethostname()
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        config_path = script_path.with_name(f"{script_name}_config.toml")
        global CONFIG
        CONFIG = load_config(config_path)
        logger_config = CONFIG.get("logging", {})
        console_log_level = getattr(logging, logger_config.get("console_logging_level", "INFO").upper(), logging.INFO)
        file_log_level = getattr(logging, logger_config.get("file_logging_level", "INFO").upper(), logging.INFO)
        log_message_format = logger_config.get("log_message_format", "%(asctime)s.%(msecs)03d %(levelname)s [%(funcName)s] - %(message)s")
        logs_folder = Path(logger_config.get("logs_folder_name", "logs"))
        log_path = logs_folder / f"{timestamp}__{script_name}__{pc_name}.log"
        setup_logging(
            logger_obj=logger,
            file_path=log_path,
            script_name=script_name,
            max_log_files=logger_config.get("max_log_files"),
            console_logging_level=console_log_level,
            file_logging_level=file_log_level,
            message_format=log_message_format
        )

        exit_behavior_config = CONFIG.get("exit_behavior", {})
        pause_before_exit = exit_behavior_config.get("always_pause", False)
        pause_before_exit_on_error = exit_behavior_config.get("pause_on_error", True)

        logger.info(f"Script: {json.dumps(script_name)} | Version: {__version__} | Host: {json.dumps(pc_name)}")
        main()

    except KeyboardInterrupt:
        logger.warning("Operation interrupted by user.")
        exit_code = 130
    except Exception as e:
        logger.error(f"A fatal error has occurred: {e}")
        exit_code = 1
    finally:
        logger.info("Closing loggers and existing.")
        for handler in logger.handlers[:]:
            handler.close()
            logger.removeHandler(handler)

    if pause_before_exit or (pause_before_exit_on_error and exit_code != 0):
        input("Press Enter to exit...")

    return exit_code


if __name__ == "__main__":
    sys.exit(bootstrap())
