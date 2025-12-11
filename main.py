import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import fitz  # PyMuPDF
from PIL import Image, ImageDraw
import re
import os
import io
import threading

# --- הגדרות ברירת מחדל (ניתן לשנות בממשק) ---
DEFAULT_MAP_START = 149
DEFAULT_MAP_END = 149
DEFAULT_TBL_START = 54
DEFAULT_TBL_END = 71

# --- פונקציות לוגיקה ועיבוד ---

def clean_currency(value):
    """מנקה טקסט וממיר למספר עשרוני"""
    if pd.isna(value) or value == '' or value is None:
        return 0.0
    # משאיר רק ספרות ונקודה
    clean_val = re.sub(r'[^\d.]', '', str(value))
    try:
        return float(clean_val)
    except ValueError:
        return 0.0

def clean_text(value):
    if value is None: return ""
    return str(value).replace('\n', ' ').strip()

def find_plot_rect(doc, start_p, end_p, plot_num):
    """
    מחפש את מיקום הטקסט (מספר המגרש) בטווח עמודי המפות.
    מחזיר: (אינדקס_עמוד, אובייקט_מלבן, אובייקט_עמוד)
    """
    plot_str = str(plot_num)
    
    # המרה לאינדקס 0
    start_idx = max(0, start_p - 1)
    end_idx = min(len(doc), end_p)

    for i in range(start_idx, end_idx):
        page = doc.load_page(i)
        # חיפוש מדויק של הטקסט
        found = page.search_for(plot_str)
        if found:
            # מחזיר את המופע הראשון שנמצא
            return i, found[0], page
            
    return None, None, None

def create_plot_image(page, rect, output_path):
    """גוזר את העמוד כתמונה ומצייר ריבוע אדום סביב המיקום שנמצא"""
    try:
        # רזולוציה גבוהה (zoom=2)
        mat = fitz.Matrix(2, 2)
        pix = page.get_pixmap(matrix=mat)
        img_data = pix.tobytes("png")
        
        img = Image.open(io.BytesIO(img_data))
        draw = ImageDraw.Draw(img)
        
        # התאמת קואורדינטות (מכיוון שהגדלנו את התמונה פי 2)
        x0, y0, x1, y1 = rect.x0 * 2, rect.y0 * 2, rect.x1 * 2, rect.y1 * 2
        
        # הרחבת הריבוע שיהיה בולט
        pad = 20
        shape = [x0 - pad, y0 - pad, x1 + pad, y1 + pad]
        
        # ציור מלבן אדום עבה
        draw.rectangle(shape, outline="red", width=6)
        
        img.save(output_path)
        return True
    except Exception as e:
        print(f"Error drawing image: {e}")
        return False

def process_tender_logic(pdf_path, t_start, t_end, m_start, m_end, status_cb):
    """הפונקציה המרכזית שמבצעת את כל העבודה"""
    output_dir = os.path.dirname(pdf_path)
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    
    # יצירת תיקייה לתמונות
    images_folder = f"{base_name}_Images"
    images_path = os.path.join(output_dir, images_folder)
    if not os.path.exists(images_path):
        os.makedirs(images_path)
    
    output_excel = os.path.join(output_dir, f"{base_name}_Analyzed.xlsx")
    
    # טעינת PDF לחיפוש גרפי
    doc_gfx = fitz.open(pdf_path)
    
    all_rows = []
    
    status_cb("קורא נתונים מהטבלאות...")
    
    with pdfplumber.open(pdf_path) as pdf:
        # לולאה על עמודי הטבלה
        for i in range(t_start - 1, t_end):
            if i >= len(pdf.pages): break
            page = pdf.pages[i]
            table = page.extract_table()
            if not table: continue
            
            for row in table:
                # ניקוי שורה
                c_row = [x for x in row if x is not None and str(x).strip() != '']
                if not c_row: continue
                r = [clean_text(x) for x in row]
                
                # זיהוי שורה רלוונטית (מספר סידורי בסוף)
                if len(r) > 10 and r[-1].isdigit():
                    try:
                        # זיהוי היסט בעמודות (לפעמים יש עמודת % בהתחלה)
                        offset = 0
                        if "%" in r[0] or "0.00" in r[0]: 
                            offset = 2
                        
                        plot_num = r[-3]
                        
                        # שליפת מחירים וחישוב
                        dev_cost = clean_currency(r[0 + offset])
                        warrior_price = clean_currency(r[-13])
                        total_price = dev_cost + warrior_price
                        
                        item = {
                            "מספר": int(r[-1]),
                            "מספר תוכנית": r[-2],
                            "מספר מגרש": plot_num,
                            "מספר חוזה": r[-4],
                            "שטח המגרש": r[-5],
                            "מלוא ערך הקרקע": clean_currency(r[-7]),
                            "מחיר קרקע לאחר הנחת אזור עדיפות לאומית למי שאיננו מחוסר דיור": clean_currency(r[-8]),
                            "מחיר הקרקע למחוסר דיור": clean_currency(r[-9]),
                            "מחיר הקרקע לחייל מילואים שאינו מחוסר דיור": clean_currency(r[-10]),
                            "מחיר הקרקע לחייל מילואים חסר דיור": clean_currency(r[-11]),
                            "מחיר הקרקע ללוחם חסר דיור": clean_currency(r[-12]),
                            "מחיר הקרקע ללוחם שאינו מחוסר דיור": warrior_price,
                            "הוצאות פיתוח": dev_cost,
                            "סה\"כ לי": total_price
                        }
                        
                        # --- טיפול במפה ---
                        link_val = "לא נמצא"
                        if m_start > 0 and m_end > 0:
                            # עדכון סטטוס למשתמש (כי זה לוקח זמן)
                            if int(r[-1]) % 5 == 0:
                                status_cb(f"מאתר במפה מגרש {plot_num}...")
                                
                            page_idx, rect, page_obj = find_plot_rect(doc_gfx, m_start, m_end, plot_num)
                            
                            if page_idx is not None:
                                # שם קובץ תמונה ייחודי
                                img_name = f"Plot_{plot_num}_p{page_idx+1}.png"
                                img_full = os.path.join(images_path, img_name)
                                
                                # יצירת התמונה (אם לא קיימת כבר)
                                if not os.path.exists(img_full):
                                    create_plot_image(page_obj, rect, img_full)
                                
                                # יצירת לינק לאקסל
                                rel_path = os.path.join(images_folder, img_name)
                                link_val = f'=HYPERLINK("{rel_path}", "הצג מפה")'
                        
                        item["מפה ויזואלית"] = link_val
                        all_rows.append(item)
                        
                    except Exception as e:
                        print(f"Skipping row: {e}")
                        continue

    doc_gfx.close()
    
    # יצירת אקסל
    status_cb("שומר קובץ אקסל...")
    df = pd.DataFrame(all_rows)
    
    if df.empty:
        raise ValueError("לא נמצאו נתונים בטווח העמודים שהוגדר.")
        
    # סדר העמודות המבוקש
    cols = [
        "מספר", "מספר תוכנית", "מספר מגרש", "מפה ויזואלית", "מספר חוזה", "שטח המגרש", 
        "מלוא ערך הקרקע", 
        "מחיר קרקע לאחר הנחת אזור עדיפות לאומית למי שאיננו מחוסר דיור",
        "מחיר הקרקע למחוסר דיור", "מחיר הקרקע לחייל מילואים שאינו מחוסר דיור",
        "מחיר הקרקע לחייל מילואים חסר דיור", "מחיר הקרקע ללוחם חסר דיור",
        "מחיר הקרקע ללוחם שאינו מחוסר דיור", "הוצאות פיתוח", "סה\"כ לי"
    ]
    # סינון עמודות קיימות בלבד
    final_cols = [c for c in cols if c in df.columns]
    df = df[final_cols]
    
    # כתיבה
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='מגרשים')
        ws = writer.sheets['מגרשים']
        ws.sheet_view.rightToLeft = True # כיוון מימין לשמאל
        
        # רוחב עמודות
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 18
            
    return output_excel

# --- ממשק משתמש (GUI) ---

class TenderAppRTL:
    def __init__(self, root):
        self.root = root
        self.root.title("מנתח מכרזים - רמ\"י")
        self.root.geometry("600x450")
        self.root.resizable(False, False)
        
        # כותרת
        header = tk.Label(root, text="מערכת לניתוח מכרזים ומיפוי", font=("Segoe UI", 18, "bold"), fg="#333")
        header.pack(pady=15)
        
        # --- אזור בחירת קובץ ---
        file_frame = tk.Frame(root)
        file_frame.pack(fill="x", padx=20, pady=5)
        
        # כפתור בצד ימין
        self.btn_browse = tk.Button(file_frame, text="טען קובץ PDF", command=self.browse_file, 
                                    bg="#0078D7", fg="white", font=("Segoe UI", 10, "bold"), width=15)
        self.btn_browse.pack(side="right", padx=(5, 0))
        
        # תווית שם הקובץ משמאל
        self.lbl_file = tk.Label(file_frame, text="לא נבחר קובץ...", relief="sunken", anchor="e", bg="white")
        self.lbl_file.pack(side="left", fill="x", expand=True)
        
        # --- אזור הגדרות (LabelFrame) ---
        settings_frame = tk.LabelFrame(root, text="  הגדרות עמודים (לפי ה-PDF)  ", font=("Segoe UI", 10, "bold"), padx=10, pady=10)
        settings_frame.pack(fill="x", padx=20, pady=15)
        
        # שימוש ב-Grid לסידור יפה מימין לשמאל
        # עמודה 3 (הכי ימנית): תווית
        # עמודה 2: שדה קלט
        # עמודה 1: רווח
        # עמודה 0: צד שמאל
        
        # שורה 1: טבלאות
        tk.Label(settings_frame, text=":התחלת טבלה (עמוד)", font=("Segoe UI", 10)).grid(row=0, column=3, sticky="e", pady=5)
        self.ent_tbl_start = tk.Entry(settings_frame, width=6, justify="center")
        self.ent_tbl_start.insert(0, str(DEFAULT_TBL_START))
        self.ent_tbl_start.grid(row=0, column=2, padx=5)
        
        tk.Label(settings_frame, text=":סיום טבלה (עמוד)", font=("Segoe UI", 10)).grid(row=0, column=1, sticky="e", pady=5)
        self.ent_tbl_end = tk.Entry(settings_frame, width=6, justify="center")
        self.ent_tbl_end.insert(0, str(DEFAULT_TBL_END))
        self.ent_tbl_end.grid(row=0, column=0, padx=5)
        
        # שורה 2: מפות
        tk.Label(settings_frame, text=":התחלת מפות (עמוד)", font=("Segoe UI", 10), fg="red").grid(row=1, column=3, sticky="e", pady=5)
        self.ent_map_start = tk.Entry(settings_frame, width=6, justify="center")
        self.ent_map_start.insert(0, str(DEFAULT_MAP_START))
        self.ent_map_start.grid(row=1, column=2, padx=5)
        
        tk.Label(settings_frame, text=":סיום מפות (עמוד)", font=("Segoe UI", 10), fg="red").grid(row=1, column=1, sticky="e", pady=5)
        self.ent_map_end = tk.Entry(settings_frame, width=6, justify="center")
        self.ent_map_end.insert(0, str(DEFAULT_MAP_END))
        self.ent_map_end.grid(row=1, column=0, padx=5)
        
        # כפתור ביצוע
        self.btn_run = tk.Button(root, text="התחל עיבוד", command=self.start_process, state="disabled",
                                 bg="#28a745", fg="white", font=("Segoe UI", 14, "bold"), width=20, height=1)
        self.btn_run.pack(pady=20)
        
        # שורת סטטוס
        self.lbl_status = tk.Label(root, text="ממתין לפעולה...", fg="gray", font=("Segoe UI", 9))
        self.lbl_status.pack(side="bottom", pady=10)
        
        self.file_path = None

    def browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if path:
            self.file_path = path
            self.lbl_file.config(text=os.path.basename(path))
            self.btn_run.config(state="normal")
            self.lbl_status.config(text="קובץ נטען. אנא וודא מספרי עמודים ולחץ על התחל.")

    def start_process(self):
        try:
            ts = int(self.ent_tbl_start.get())
            te = int(self.ent_tbl_end.get())
            ms = int(self.ent_map_start.get())
            me = int(self.ent_map_end.get())
            
            if ts > te:
                messagebox.showerror("שגיאה", "עמוד התחלת טבלה חייב להיות קטן מעמוד סיום")
                return
                
        except ValueError:
            messagebox.showerror("שגיאה", "נא להזין מספרים שלמים בלבד בשדות העמודים")
            return
            
        self.btn_run.config(state="disabled")
        self.btn_browse.config(state="disabled")
        
        # הרצה ברקע
        threading.Thread(target=self.run_logic, args=(ts, te, ms, me)).start()

    def run_logic(self, ts, te, ms, me):
        try:
            output_file = process_tender_logic(self.file_path, ts, te, ms, me, self.update_status)
            self.update_status("הסתיים בהצלחה!")
            messagebox.showinfo("הצלחה", f"התהליך הסתיים!\nהקובץ נשמר ב:\n{output_file}")
        except Exception as e:
            self.update_status("שגיאה")
            messagebox.showerror("שגיאה בתהליך", str(e))
        finally:
            self.btn_run.config(state="normal")
            self.btn_browse.config(state="normal")

    def update_status(self, text):
        self.lbl_status.config(text=text)

if __name__ == "__main__":
    root = tk.Tk()
    app = TenderAppRTL(root)
    root.mainloop()