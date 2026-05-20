import os
from pathlib import Path
import win32com.client
import win32gui
import win32con
import pythoncom
import time

BASE_DIR = Path(__file__).parent.resolve()

def get_foreground_window():
    return win32gui.GetForegroundWindow()

def restore_foreground(hwnd):
    if hwnd:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)

def merge_presentations(files, output_file, progress_callback=None):
    if len(files) < 2:
        raise Exception("צריך לפחות 2 מצגות (PPTX) לצורך מיזוג")

    output_path = Path(output_file).resolve()
    temp_output_path = output_path.with_name(f"~temp_{output_path.name}")

    # שומרים את החלון הפעיל של המשתמש לפני שמתחילים
    user_window = get_foreground_window()

    pythoncom.CoInitialize()

    try:
        powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
        powerpoint.Visible = True
        powerpoint.WindowState = 2
        restore_foreground(user_window)
    except Exception as dispatch_err:
        raise Exception(f"שגיאה באתחול PowerPoint: {str(dispatch_err)}")

    main_presentation = None
    try:
        main_presentation = powerpoint.Presentations.Open(
            str(Path(files[0]).resolve()),
            ReadOnly=0,
            WithWindow=1
        )
        restore_foreground(user_window)

        if progress_callback:
            progress_callback(files[0])

        for pptx_file in files[1:]:
            full_path = str(Path(pptx_file).resolve())

            source_pres = powerpoint.Presentations.Open(
                full_path,
                ReadOnly=1,
                WithWindow=1
            )
            restore_foreground(user_window)

            slide_count = source_pres.Slides.Count
            if slide_count > 0:
                source_pres.Slides.Range().Copy()
                time.sleep(0.2)
                main_presentation.Windows(1).Activate()
                last_slide_idx = main_presentation.Slides.Count
                main_presentation.Slides(last_slide_idx).Select()
                powerpoint.CommandBars.ExecuteMso("PasteSourceFormatting")
                time.sleep(0.2)
                restore_foreground(user_window)  # מחזיר את המשתמש לקדמה אחרי ה-Paste

            source_pres.Close()
            restore_foreground(user_window)

            if progress_callback:
                progress_callback(pptx_file)
            time.sleep(0.1)

        if temp_output_path.exists():
            os.remove(temp_output_path)
        main_presentation.SaveAs(str(temp_output_path))
        main_presentation.Close()

    except Exception as e:
        raise Exception(f"המיזוג נכשל במהלך העבודה: {str(e)}")
    finally:
        try:
            powerpoint.Quit()
        except:
            pass
        pythoncom.CoUninitialize()
        time.sleep(0.3)

    try:
        if temp_output_path.exists():
            if output_path.exists():
                os.remove(output_path)
            os.rename(temp_output_path, output_path)
    except Exception as file_err:
        raise Exception(f"המיזוג הצליח, אך נכשל בשחזור שם הקובץ הסופי: {str(file_err)}")