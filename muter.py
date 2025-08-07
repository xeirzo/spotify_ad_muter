# File: spotify_ad_muter.py
import time
import psutil
import win32gui
import win32process
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume

# ----------------------
# Window Handle Functions
# ----------------------
def get_spotify_hwnd():
    hwnd_list = []

    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            try:
                process = psutil.Process(pid)
                if process.name().lower() == "spotify.exe":
                    hwnd_list.append(hwnd)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        return True

    win32gui.EnumWindows(callback, None)
    return hwnd_list[0] if hwnd_list else None


def get_spotify_window_title():
    hwnd = get_spotify_hwnd()
    if hwnd:
        return win32gui.GetWindowText(hwnd).strip()
    return None

# ----------------------
# Audio Control Functions
# ----------------------
def get_spotify_volume_interface():
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        try:
            if session.Process and session.Process.name().lower() == "spotify.exe":
                return session._ctl.QueryInterface(ISimpleAudioVolume)
        except AttributeError:
            pass
    return None


def mute_spotify_with_retry(mute=True, retries=5, delay=0.2):
    """Retry muting/unmuting in case the session isn't ready yet."""
    for _ in range(retries):
        spotify_volume = get_spotify_volume_interface()
        if spotify_volume:
            spotify_volume.SetMute(mute, None)
            return True
        time.sleep(delay)
    return False

# ----------------------
# Ad Detection
# ----------------------
def is_ad_playing():
    title = get_spotify_window_title()
    return title is not None and title.lower() == "advertisement"

# ----------------------
# Main Loop
# ----------------------
def main():
    was_muted = False
    last_title = None
    print("🎵 Spotify Ad Muter started.")
    while True:
        try:
            current_title = get_spotify_window_title()

            # Print song change
            if current_title and current_title != last_title:
                print(f"🎶 Now Playing: {current_title}")
                last_title = current_title

            # Ad detection + muting
            if is_ad_playing():
                if not was_muted:
                    print("🔇 Muting ad...")
                    if mute_spotify_with_retry(True):
                        was_muted = True
                    else:
                        print("⚠️ Could not mute — session not found.")
            else:
                if was_muted:
                    print("🔊 Unmuting...")
                    if mute_spotify_with_retry(False):
                        was_muted = False
                    else:
                        print("⚠️ Could not unmute — session not found.")
        except Exception as e:
            print("⚠️ Error:", e)
        time.sleep(1)


if __name__ == "__main__":
    main()