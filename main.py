from pathlib import Path
import win32com.client
import time

BASE_DIR = Path(__file__).parent.resolve()


def merge_presentations(files, output_file, progress_callback=None):
    if len(files) < 2:
        raise Exception("צריך לפחות 2 מצגות (PPTX)")

    # אתחול אובייקט ה-Application של PowerPoint
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    main_presentation = None

    try:
        # פותחים את המצגת הראשונה כבסיס (במצב חבוי WithWindow=0)
        # זה מונע את פתיחת המצגת הריקה המיותרת ופותר את בעיית ה-SlideLayouts
        main_presentation = powerpoint.Presentations.Open(
            str(Path(files[0]).resolve()),
            ReadOnly=0,
            Untitled=0,
            WithWindow=0
        )

        if progress_callback:
            progress_callback(files[0])

    except Exception as e:
        try:
            powerpoint.Quit()
        except:
            pass
        raise Exception(f"נכשל בפתיחת המצגת הראשית: {str(e)}")

    try:
        # עוברים על שאר הקבצים (החל מהשני) וממזגים אותם פנימה
        for pptx_file in files[1:]:
            full_path = str(Path(pptx_file).resolve())

            # 1. פתיחה חבויה של מצגת המקור כדי לבדוק כמות שקפים
            source_pres = powerpoint.Presentations.Open(full_path, ReadOnly=1, WithWindow=0)
            slide_count = source_pres.Slides.Count
            source_pres.Close()

            if slide_count > 0:
                # שומרים את מיקום השקף האחרון הנוכחי במצגת הממוזגת לפני ההזרקה
                start_index = main_presentation.Slides.Count

                # 2. הזרקת השקפים. פקודה זו שומרת על הטקסטים והאלמנטים
                main_presentation.Slides.InsertFromFile(
                    full_path,
                    start_index,
                    1,
                    slide_count
                )

                # 3. פתיחת מצגת המקור שוב כדי להחיל את ערכת העיצוב (Design) בצורה ישירה
                # ללא צורך בלולאות פנימיות על ה-Layouts שגורמות לשגיאות ברקע
                source_pres = powerpoint.Presentations.Open(full_path, ReadOnly=1, WithWindow=0)

                for i in range(1, slide_count + 1):
                    inserted_slide_idx = start_index + i
                    target_slide = main_presentation.Slides(inserted_slide_idx)
                    source_slide = source_pres.Slides(i)

                    try:
                        # השמה ישירה של ה-Design והרמוניה מול ה-Master המקורי
                        target_slide.Design = source_slide.Design

                        # במקום לגעת ב-CustomLayout הבעייתי, אנחנו מאפשרים ל-PowerPoint
                        # להשתמש במבנה הקיים שהגיע מה-InsertFromFile
                        if source_slide.FollowMasterBackground == 0:
                            target_slide.FollowMasterBackground = 0
                            target_slide.Background.Fill.ForeColor.RGB = source_slide.Background.Fill.ForeColor.RGB
                    except Exception as slide_err:
                        # החלקה של שגיאות קטנות כדי למנוע קריסה קריטית
                        print(f"הערה: עיצוב שקף {i} הוחל חלקית ({slide_err})")
                        continue

                source_pres.Close()
                time.sleep(0.05)

            if progress_callback:
                progress_callback(pptx_file)

        # שמירת המצגת הממוזגת הסופית בנתיב החדש
        main_presentation.SaveAs(str(Path(output_file).resolve()))

    finally:
        # סגירה הרמטית ובטוחה של תהליכי הרקע מהזיכרון
        try:
            if main_presentation:
                main_presentation.Close()
        except:
            pass
        try:
            powerpoint.Quit()
        except:
            pass