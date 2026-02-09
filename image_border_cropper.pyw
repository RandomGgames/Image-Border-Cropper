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

import ctypes
import datetime
import hashlib
import io
import json
import logging
import os
import send2trash
import socket
import sys
import threading
import time
import tomllib
import typing
import win32clipboard
import win32con
import win32gui
from datetime import datetime
from pathlib import Path
from PIL import Image
from PIL import Image, ImageGrab, ImageChops
from pystray import Icon, MenuItem, Menu
from ctypes import wintypes

logger = logging.getLogger(__name__)

__version__ = "2.0.0"  # Major.Minor.Patch

CONFIG = {}

WM_CLIPBOARDUPDATE = 0x031D
last_hash = None
ignore_next = False
hwnd = None
tray_icon = None

exit_event = threading.Event()


def load_image(path: str | Path) -> Image.Image:
    path = Path(path)
    image = Image.open(path)
    logger.debug(f"Loaded image at path {json.dumps(str(path))}")
    return image


def open_script_folder():
    folder_path = os.path.dirname(os.path.abspath(__file__))
    os.startfile(folder_path)
    logger.debug(f"Opened script folder: {json.dumps(str(folder_path))}")


def on_exit(icon):
    global hwnd, tray_icon
    logger.debug("Exit pressed on system tray icon")
    if tray_icon:
        tray_icon.stop()
        logger.debug("Tray icon stopped")
    if hwnd:
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    exit_event.set()


def startup_tray_icon():
    global tray_icon
    logger.debug("Starting system tray icon...")
    image = load_image("system_tray_icon.png")
    menu = Menu(
        MenuItem("Open Folder", open_script_folder),
        MenuItem("Exit", on_exit)
    )
    tray_icon = Icon("ClipboardWatcher", image, menu=menu)
    tray_icon.run()


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

            win32clipboard.OpenClipboard()
            has_text = (
                win32clipboard.IsClipboardFormatAvailable(win32con.CF_TEXT) or
                win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT)
            )
            win32clipboard.CloseClipboard()

            if has_text:
                logger.info("Clipboard contains text/numbers. Skipping image processing.")
                return 0

            img = ImageGrab.grabclipboard()
            if isinstance(img, Image.Image):
                current_hash = image_hash(img)
                if current_hash == last_hash:
                    return 0  # same image, ignore

                logger.info("New image detected in clipboard.")
                cropped = crop_to_object(img, CONFIG.get("border_size", 10))
                image_to_clipboard(cropped)
                last_hash = image_hash(cropped)
                ignore_next = True
        except Exception as e:
            logger.warning(f"Clipboard processing error: {e}")
    elif msg == win32con.WM_DESTROY:
        # Stop message pump
        win32gui.PostQuitMessage(0)
        return 0
    return 0


def main():
    border_size = CONFIG.get("border_size", 10)
    CONFIG["border_size"] = border_size

    # Start system tray icon
    system_tray_thread = threading.Thread(target=startup_tray_icon, daemon=True)
    system_tray_thread.start()

    # Setup hidden window for clipboard listener
    wc = typing.cast(typing.Any, win32gui.WNDCLASS())
    wc.lpfnWndProc = on_clipboard_update
    wc.lpszClassName = "ClipboardWatcher"
    hinst = win32gui.GetModuleHandle(None)
    wc.hInstance = hinst
    classAtom = win32gui.RegisterClass(wc)

    global hwnd
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
    user32 = ctypes.windll.user32

    while not exit_event.is_set():
        msg = wintypes.MSG()
        while user32.PeekMessageW(ctypes.byref(msg), 0, 0, 0, 1):
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))
        time.sleep(0.05)  # small sleep to avoid 100% CPU


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
        logger.info("Execution completed.")

    except KeyboardInterrupt:
        logger.warning("Operation interrupted by user.")
        exit_code = 130
    except Exception as e:  # pylint: disable=broad-exception-caught
        # Using 'err' or 'exc' is standard; logging the traceback handles the 'broad-except'
        logger.error(f"A fatal error has occurred: {e}")
        exit_code = 1
    finally:
        for handler in logger.handlers[:]:
            handler.close()
            logger.removeHandler(handler)

    if pause_before_exit or (pause_before_exit_on_error and exit_code != 0):
        input("Press Enter to exit...")

    return exit_code


if __name__ == "__main__":
    sys.exit(bootstrap())
