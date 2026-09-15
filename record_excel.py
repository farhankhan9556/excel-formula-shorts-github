import os
import subprocess
import psutil
import win32gui
import win32process
import ctypes


def get_desktop_size():
    """
    Get the actual Windows desktop dimensions.
    """
    user32 = ctypes.windll.user32

    width = user32.GetSystemMetrics(0)
    height = user32.GetSystemMetrics(1)

    return width, height


def find_excel_window():
    windows = []

    def callback(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True

        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)

            if pid <= 0:
                return True

            process = psutil.Process(pid)

            if process.name().lower() == "excel.exe":
                title = win32gui.GetWindowText(hwnd)

                if title:
                    windows.append((hwnd, pid, title))

        except Exception:
            pass

        return True

    win32gui.EnumWindows(callback, None)

    if not windows:
        raise RuntimeError(
            "Microsoft Excel window was not found. "
            "Please open Excel first."
        )

    # Prefer the largest visible Excel window.
    best = None
    best_area = 0

    for hwnd, pid, title in windows:
        try:
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)

            width = max(0, right - left)
            height = max(0, bottom - top)
            area = width * height

            if area > best_area:
                best_area = area
                best = (hwnd, pid, title)

        except Exception:
            pass

    if best is None:
        best = windows[0]

    hwnd, pid, title = best

    print(f"Excel window found: {title}")
    print(f"Excel HWND: {hwnd}")
    print(f"Excel PID: {pid}")

    return hwnd, pid


def get_capture_area(hwnd):
    """
    Get Excel's visible screen rectangle and clamp it
    safely inside the actual Windows desktop.
    """

    desktop_width, desktop_height = get_desktop_size()

    print(
        f"Actual Windows desktop: "
        f"{desktop_width}x{desktop_height}"
    )

    left, top, right, bottom = win32gui.GetWindowRect(hwnd)

    print(
        f"Excel window rectangle before clamp: "
        f"left={left}, top={top}, "
        f"right={right}, bottom={bottom}"
    )

    # Clamp starting position.
    left = max(0, left)
    top = max(0, top)

    # Clamp ending position.
    right = min(desktop_width, right)
    bottom = min(desktop_height, bottom)

    width = right - left
    height = bottom - top

    if width <= 0 or height <= 0:
        raise RuntimeError(
            "Excel window is outside the visible desktop."
        )

    # H.264 works best with even dimensions.
    width -= width % 2
    height -= height % 2

    if width < 320 or height < 240:
        raise RuntimeError(
            f"Excel capture area is too small: "
            f"{width}x{height}"
        )

    print(
        f"Excel capture area: "
        f"x={left}, y={top}, "
        f"width={width}, height={height}"
    )

    return left, top, width, height


def start_ffmpeg_recording(output_file, duration=30):

    hwnd, pid = find_excel_window()

    left, top, width, height = get_capture_area(hwnd)

    output_file = os.path.abspath(output_file)

    output_dir = os.path.dirname(output_file)

    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    command = [
        "ffmpeg",
        "-y",

        "-f",
        "gdigrab",

        "-framerate",
        "30",

        "-draw_mouse",
        "1",

        "-offset_x",
        str(left),

        "-offset_y",
        str(top),

        "-video_size",
        f"{width}x{height}",

        "-t",
        str(duration),

        "-i",
        "desktop",

        "-c:v",
        "libx264",

        "-preset",
        "veryfast",

        "-pix_fmt",
        "yuv420p",

        output_file,
    ]

    print()
    print("Starting FFmpeg recording...")
    print(f"Duration: {duration} seconds")
    print(f"Resolution: {width}x{height}")
    print(f"Output: {output_file}")

    # Do not use stdout/stderr PIPE.
    # FFmpeg writes directly to the CMD window.
    process = subprocess.Popen(
        command,
        stdin=subprocess.DEVNULL,
        stdout=None,
        stderr=None,
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
    )

    return process


def record_excel(output_file, duration=30):
    return start_ffmpeg_recording(
        output_file=output_file,
        duration=duration,
    )


if __name__ == "__main__":

    import argparse

    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--output",
        required=True,
        help="Output MP4 file",
    )

    parser.add_argument(
        "--duration",
        type=int,
        default=30,
        help="Recording duration",
    )

    args = parser.parse_args()

    recorder = record_excel(
        args.output,
        args.duration,
    )

    return_code = recorder.wait()

    if return_code != 0:
        raise SystemExit(
            f"FFmpeg failed with exit code {return_code}"
        )

    print()
    print("============================================================")
    print("RECORDING COMPLETED SUCCESSFULLY")
    print("============================================================")