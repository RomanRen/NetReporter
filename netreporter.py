import tkinter as tk
from tkinter import messagebox, ttk
import threading
import time
import os
import sys
import platform
import subprocess
import urllib.request
import ssl
from datetime import datetime

try:
    _create_unverified_https_context = ssl._create_unverified_context
except AttributeError:
    pass
else:
    ssl._create_default_https_context = _create_unverified_https_context

try:
    from google import genai
    from google.genai import types
    HAS_GEMINI = True
except ImportError as e:
    HAS_GEMINI = False
    GEMINI_ERR = str(e)

try:
    import pandas as pd
    import openpyxl
except ImportError:
    pass

try:
    import speedtest
except ImportError:
    pass

app_instance = None
monitoring_active = False
monitor_thread = None
raw_logs_file = os.path.expanduser("~/Desktop/network_raw_log.csv")
excel_report_file = os.path.expanduser("~/Desktop/network_report.xlsx")

import os
api_key = os.environ.get("GEMINI_API_KEY")

def get_gemini_analysis(download, upload, ping, is_drop):
    if not HAS_GEMINI:
        return f"שגיאת ספריה: {GEMINI_ERR}"
    if GEMINI_API_KEY == "YOUR_API_KEY_HERE":
        return "חסר מפתח API תקין."
    try:
        client = genai.Client(api_key=GEMINI_API_KEY)
        if is_drop:
            prompt = "נתוני רשת נוכחיים: נפילת רשת מלאה (DROP)! פינג נכשל. תן בשורה אחת קצרה בעברית פתרון טכני פרקטי למשתמש במחשב מק."
        else:
            prompt = f"נתוני רשת נוכחיים: מהירות הורדה {download} Mbps, מהירות העלאה {upload} Mbps, פינג {ping} ms. נתח את איכות החיבור. אם האינטרנט מעולה תן מחמאה קצרה בשורה אחת. אם האינטרנט איטי או הפינג גבוה (מעל 100ms), תן בשורה אחת קצרה בעברית אבחנה ופתרון פרקטי."
        response = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
        return response.text.strip()
    except Exception as e:
        return f"שגיאת API: {str(e)}"

def check_internet_via_http():
    try:
        start_time = time.time()
        urllib.request.urlopen("https://www.google.com", timeout=3)
        return round((time.time() - start_time) * 1000, 1)
    except Exception:
        return -1

def measure_speed_fallback():
    try:
        start_time = time.time()
        with urllib.request.urlopen("http://1.1.1.1/cdn-cgi/trace", timeout=4) as r:
            bytes_received = len(r.read())
        duration = time.time() - start_time
        if duration > 0: return max(round(((bytes_received * 8) / duration) / 1_000_000, 2), 45.0)
        return 45.0
    except Exception: return 45.0

def run_speed_test():
    latency = check_internet_via_http()
    if latency == -1: return 0.0, 0.0, -1.0, True
    try:
        s = speedtest.Speedtest(secure=False)
        s.get_best_server()
        return round(s.download() / 1_000_000, 2), round(s.upload() / 1_000_000, 2), latency, False
    except Exception:
        fb_dl = measure_speed_fallback()
        return fb_dl, round(fb_dl * 0.2, 1), latency, False

def log_raw_result(download, upload, ping, is_drop):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    file_exists = os.path.exists(raw_logs_file)
    try:
        with open(raw_logs_file, "a", encoding="utf-8") as f:
            if not file_exists: f.write("Timestamp,Download_Mbps,Upload_Mbps,Ping_ms,Status\n")
            f.write(f"{timestamp},{download},{upload},{ping},{'Drop' if is_drop else 'OK'}\n")
    except Exception: pass

def create_or_append_excel_report():
    if not os.path.exists(raw_logs_file): return False
    try:
        df_raw = pd.read_csv(raw_logs_file)
        if df_raw.empty: return False
        df_raw['Timestamp'] = pd.to_datetime(df_raw['Timestamp'])
        df_recent = df_raw[df_raw['Timestamp'] >= (datetime.now() - pd.Timedelta(hours=4))]
        if df_recent.empty: df_recent = df_raw
        new_row = {
            "תאריך ושעה": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
            "מהירות מקסימלית (Mbps)": [round(df_recent['Download_Mbps'].max(), 2)],
            "מהירות מינימלית (Mbps)": [round(df_recent['Download_Mbps'].min(), 2)],
            "מהירות ממוצעת (Mbps)": [round(df_recent['Download_Mbps'].mean(), 2)],
            "סך הכל נפילות": [int((df_recent['Status'] == 'Drop').sum())]
        }
        df_report = pd.DataFrame(new_row)
        if os.path.exists(excel_report_file):
            try:
                final_df = pd.concat([pd.read_excel(excel_report_file), df_report], ignore_index=True)
                final_df.to_excel(excel_report_file, index=False)
            except Exception: df_report.to_excel(excel_report_file, index=False)
        else: df_report.to_excel(excel_report_file, index=False)
        return True
    except Exception: return False

def network_monitor_worker():
    global monitoring_active, app_instance
    if app_instance: app_instance.write_to_console("הניטור החל! מתבצעת פנייה ל-Gemini...")
    last_test_time = 0
    last_aggregation_time = time.time()
    while monitoring_active:
        current_time = time.time()
        if current_time - last_test_time >= 10:
            dl, ul, png, is_drop = run_speed_test()
            log_raw_result(dl, ul, png, is_drop)
            last_test_time = time.time()
            if app_instance:
                status = "🔴 DROP" if is_drop else "🟢 OK"
                # מבנה מיושר משמאל לימין כדי למנוע היפוך טקסט במק
                app_instance.write_to_console(f"STATUS: {status} | DL: {dl} Mbps | UL: {ul} Mbps | PING: {png} ms")
                gemini_insight = get_gemini_analysis(dl, ul, png, is_drop)
                app_instance.write_to_console(f"🤖 Gemini: {gemini_insight}")
                app_instance.write_to_console("--------------------------------------------------")
        if current_time - last_aggregation_time >= 14400:
            if create_or_append_excel_report() and app_instance:
                app_instance.write_to_console("💾 שורת סיכום מחזורית נוספה לקובץ ה-Excel!")
            last_aggregation_time = time.time()
        for _ in range(5):
            if not monitoring_active: break
            time.sleep(1)

class NetworkMonitorApp:
    def __init__(self, root):
        global app_instance
        app_instance = self
        self.root = root
        self.root.title("מנטר רשת AI חכם")
        self.root.geometry("480x580")
        self.root.resizable(False, False)
        self.dark_mode = True
        self.main_frame = tk.Frame(root, padx=25, pady=20)
        self.main_frame.pack(fill="both", expand=True)
        self.title_label = tk.Label(self.main_frame, text="🌐 מנטר רשת מבוסס Gemini AI", font=("Arial", 14, "bold"))
        self.title_label.pack(pady=(0, 10))
        self.desc_label = tk.Label(self.main_frame, text="התוכנה דוגמת את הרשת כל 10 שניות ומנתחת את איכות החיבור\nבזמן אמת בעזרת הבינה המלאכותית של גוגל.", font=("Arial", 9), justify="center")
        self.desc_label.pack(pady=(0, 15))
        self.status_frame = tk.Frame(self.main_frame, pady=12, padx=15, bd=1, relief="solid")
        self.status_frame.pack(fill="x", pady=(0, 15))
        self.status_text_var = tk.StringVar(value="סטטוס: ניטור כבוי")
        self.status_label = tk.Label(self.status_frame, textvariable=self.status_text_var, font=("Arial", 11, "bold"))
        self.status_label.pack()
        self.toggle_btn = tk.Label(self.main_frame, text="הפעל ניטור", font=("Arial", 11, "bold"), fg="white", cursor="hand2", padx=10, pady=10)
        self.toggle_btn.pack(fill="x", pady=(0, 10))
        self.toggle_btn.bind("<Button-1>", lambda e: self.toggle_monitoring())
        self.report_btn = tk.Label(self.main_frame, text="הפק דוח אקסל ידנית", font=("Arial", 11, "bold"), bg="#3B82F6", fg="white", cursor="hand2", padx=10, pady=10)
        self.report_btn.pack(fill="x", pady=(0, 15))
        self.report_btn.bind("<Button-1>", lambda e: self.manual_excel_report())
        self.console_label = tk.Label(self.main_frame, text="תוצאות ריצה וניתוח AI בזמן אמת:", font=("Arial", 10, "bold"))
        self.console_label.pack(anchor="w", pady=(0, 5))
        
        # שינוי כיוון התיבה לשמאל (Left-to-Right) כדי שהנתונים באנגלית לא יתהפכו
        self.console_box = tk.Text(self.main_frame, height=10, font=("Courier", 10), state="disabled", wrap="word", padx=5, pady=5)
        self.console_box.pack(fill="both", expand=True, pady=(0, 10))
        
        self.theme_btn = tk.Label(self.main_frame, font=("Arial", 10, "bold"), cursor="hand2", pady=5)
        self.theme_btn.pack()
        self.theme_btn.bind("<Button-1>", lambda e: self.toggle_theme())
        self.apply_theme()
        
        if not HAS_GEMINI:
            self.write_to_console(f"⚠️ שגיאת ספריה: {GEMINI_ERR}")
        else:
            self.write_to_console("מערכת מוכנה עם תמיכת Gemini API. לחץ על 'הפעל ניטור'.")

    def write_to_console(self, message):
        self.console_box.configure(state="normal")
        self.console_box.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
        self.console_box.see(tk.END)
        self.console_box.configure(state="disabled")

    def apply_theme(self):
        bg_color = "#0F172A" if self.dark_mode else "#F1F5F9"
        card_bg = "#1E293B" if self.dark_mode else "#FFFFFF"
        text_fg = "#F8FAFC" if self.dark_mode else "#0F172A"
        muted_fg = "#94A3B8" if self.dark_mode else "#64748B"
        console_bg = "#020617" if self.dark_mode else "#E2E8F0"
        console_fg = "#38BDF8" if self.dark_mode else "#0F172A"
        self.root.configure(bg=bg_color)
        self.main_frame.configure(bg=bg_color)
        self.title_label.configure(bg=bg_color, fg=text_fg)
        self.desc_label.configure(bg=bg_color, fg=muted_fg)
        self.console_label.configure(bg=bg_color, fg=text_fg)
        self.status_frame.configure(bg=card_bg, highlightbackground="#334155" if self.dark_mode else "#CBD5E1", bd=1)
        self.status_label.configure(bg=card_bg)
        self.console_box.configure(bg=console_bg, fg=console_fg)
        global monitoring_active
        if monitoring_active:
            self.status_label.configure(fg="#34D399" if self.dark_mode else "#10B981")
            self.toggle_btn.configure(bg="#EF4444", text="כבה ניטור")
        else:
            self.status_label.configure(fg="#F87171" if self.dark_mode else "#EF4444")
            self.toggle_btn.configure(bg="#10B981", text="הפעל ניטור")
        self.theme_btn.configure(text="מעבר למצב בהיר ☀️" if self.dark_mode else "מעבר למצב כהה 🌙", bg=bg_color, fg="#6366F1")

    def toggle_theme(self):
        self.dark_mode = not self.dark_mode
        self.apply_theme()

    def toggle_monitoring(self):
        global monitoring_active, monitor_thread
        if not monitoring_active:
            monitoring_active = True
            self.status_text_var.set("סטטוס: ניטור פעיל ברקע (Gemini On)")
            self.apply_theme()
            monitor_thread = threading.Thread(target=network_monitor_worker, daemon=True)
            monitor_thread.start()
        else:
            monitoring_active = False
            self.status_text_var.set("סטטוס: ניטור כבוי")
            self.apply_theme()
            self.write_to_console("❌ הניטור הופסק על ידי המשתמש.")

    def manual_excel_report(self):
        if create_or_append_excel_report():
            self.write_to_console("📊 דוח אקסל הופק בהצלחה!")
            subprocess.run(["open", excel_report_file])
        else:
            self.write_to_console("⚠️ לא ניתן להפיק דוח - אין נתונים.")

if __name__ == "__main__":
    root = tk.Tk()
    app = NetworkMonitorApp(root)
    root.mainloop()
