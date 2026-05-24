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


def is_output_file_locked(file_path):
    """
    בודק האם קובץ היעד קיים ונעול (פתוח ע"י המשתמש)
    """
    path = Path(file_path).resolve()
    if not path.exists():
        return False

    try:
        os.rename(str(path), str(path))
        return False
    except OSError:
        return True


def merge_presentations(files, output_file, progress_callback=None):
    if len(files) < 2:
        raise Exception("צריך לפחות 2 מצגות (PPTX) לצורך מיזוג")

    if is_output_file_locked(output_file):
        raise Exception(
            f"לא ניתן לבצע מיזוג. קובץ היעד '{Path(output_file).name}' פתוח כרגע בתוכנה אחרת. נא לסגור אותו ולנסות שוב.")

    output_path = Path(output_file).resolve()
    temp_output_path = output_path.with_name(f"~temp_{output_path.name}")

    user_window = get_foreground_window()

    pythoncom.CoInitialize()

    powerpoint = None
    try:
        powerpoint = win32com.client.DispatchEx("PowerPoint.Application")
        powerpoint.Visible = True

        # --- פתרון בעיית המרכוז וה-DPI ---
        window_width = 200
        window_height = 200

        # טריק: נמקסם את החלון לשבריר שנייה רק כדי לקבל את המימדים המקסימליים ב-Points של האפליקציה
        powerpoint.WindowState = 3  # ppWindowMaximized
        max_width = powerpoint.Width
        max_height = powerpoint.Height

        # עכשיו נחזיר אותו למצב רגיל ונחשב מרכז מדויק לפי יחידות המידה של PowerPoint
        powerpoint.WindowState = 1  # ppWindowNormal
        powerpoint.Width = window_width
        powerpoint.Height = window_height

        # חישוב המרכז על בסיס יחידות המידה הפנימיות
        powerpoint.Left = int((max_width - window_width) / 2)
        powerpoint.Top = int((max_height - window_height) / 2)
        # ---------------------------------

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
                restore_foreground(user_window)

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
            if powerpoint:
                powerpoint.WindowState = 3  # מחזירים למסך מלא לפני הסגירה
                time.sleep(0.1)
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