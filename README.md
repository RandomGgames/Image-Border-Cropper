# Image Border Cropper

A small Windows utility that watches the clipboard for images, automatically crops uniform borders (preserving a small configurable margin) and replaces the clipboard image with the cropped version. Runs in the background with a system tray icon.

## Features

- Event-driven clipboard watching (no polling) using the Windows clipboard listener.
- Automatically detects and crops uniform borders while preserving a configurable border size.
- Replaces the clipboard image with the cropped result.
- Background tray icon with quick "Open Folder" and "Exit" actions.
- Configurable via a TOML file and logs actions to a `logs/` folder.

### Quick start

1. Install Python 3.10+ and Git (Windows).
2. From the project folder, install dependencies:

```cmd
pip install -r requirements.txt
```

3. Edit `image_border_cropper_config.toml` if you want to change `border_size` or logging options.

4. Run the script (to run without a console window use `pythonw`):

```cmd
pythonw image_border_cropper.pyw
```

You should see the tray icon appear. Copy an image (Print Screen, Snipping Tool or copy an image file) and the script will crop borders and update the clipboard with the cropped image.

## Configuration

The script reads `image_border_cropper_config.toml` from the same directory. Example options:

- `border_size` (integer): Number of pixels to keep around the detected object (default: `10`).
- `[logging]` keys (console/file levels, message format, `logs_folder_name`, `max_log_files`).
- `[exit_behavior]` keys to control pause-on-exit behavior.

Example (included): `image_border_cropper_config.toml`.

### How it detects clipboard updates

The script registers a hidden window as a clipboard listener using the Windows API (AddClipboardFormatListener). Windows then sends the `WM_CLIPBOARDUPDATE` message to the window whenever the clipboard contents or available formats change. The script receives that message and then reads the clipboard (the message itself does not contain the clipboard data).

### Notes:

- The window message is event-driven and preferred over polling or the old clipboard-viewer chain.
- The script briefly opens the clipboard to query formats or to write the cropped image. If another process holds the clipboard, the operation can fail; the script logs errors.

### Troubleshooting

- Clipboard locked errors: Another process may have the clipboard open (e.g., some apps or remote-desktop tools). Wait a moment and try again, or close the app that is holding the clipboard.
- Tray icon not visible: Make sure Windows hides/notifications settings allow the tray icon or open the script folder and check logs in `logs/` for startup errors.
- Script re-processing its own clipboard update: The script attempts to detect duplicates using an image hash. If you see loops, check logs. You can adjust behavior in the code (the `last_hash` variable and the clipboard write sequence).

### Development notes

- Main script: `image_border_cropper.pyw` (runs as a GUI/daemon process).
- Config file: `image_border_cropper_config.toml`.
- Logs are written to the folder configured in the TOML (defaults to `logs/`).
- Dependencies are listed in `requirements.txt` (`pillow`, `imageio`, `pywin32`).

If you modify the clipboard-writing code, make sure to:

- Set `last_hash` before writing the newly generated image to the clipboard to avoid re-processing your own clipboard update.
- Always close the clipboard after OpenClipboard to avoid leaving it locked.

# Contributing

Contributions are welcome. Please open an issue to discuss larger changes. Small fixes or documentation updates can be submitted via pull requests.

# Contact

Found a bug or want an enhancement? Open an issue in the repository.
