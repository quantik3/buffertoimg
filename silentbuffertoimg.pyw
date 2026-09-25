import keyboard
from PIL import Image
import win32clipboard
import io
import os
from datetime import datetime
import win32gui
import pythoncom
import win32com.client
import gc

def get_active_explorer_path_via_com():
    try:
        pythoncom.CoInitialize()
        shell_app = win32com.client.Dispatch("Shell.Application")
        for window in shell_app.Windows():
            if window.Visible and window.HWND == win32gui.GetForegroundWindow():
                folder_location = window.Document.Folder.Self.Path
                if os.path.isdir(folder_location):
                    return folder_location
    except Exception:
        pass
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass
    return None

def get_image_from_clipboard():
    img = None
    try:
        win32clipboard.OpenClipboard()
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_DIB):
            data = win32clipboard.GetClipboardData(win32clipboard.CF_DIB)
            img = Image.open(io.BytesIO(data))
            
            return img
        else:
            return None
    except Exception:
        return None
    finally:
        try:
            win32clipboard.CloseClipboard()
        except:
            pass

def save_image_to_active_folder():
    folder_path = get_active_explorer_path_via_com()
    if not folder_path:
        return

    img = get_image_from_clipboard()
    if img is None:
        return

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
    filename = f"clipboard_img_{timestamp}.png"
    filepath = os.path.join(folder_path, filename)

    try:
        img.save(filepath, "PNG")
    except Exception:
        pass
    finally:
    
        if img:
            img.close()
        
        gc.collect()

keyboard.add_hotkey('ctrl+b', save_image_to_active_folder)

keyboard.wait('ctrl+shift+q')
