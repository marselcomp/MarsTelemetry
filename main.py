import sys
import os
import ctypes
import tkinter as tk
from tkinter import ttk, messagebox
import platform
import subprocess
import webbrowser
import telemetry_control

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    params = ' '.join([f'"{arg}"' for arg in sys.argv[1:]])
    executable = sys.executable
    script = os.path.abspath(sys.argv[0])
    ret = ctypes.windll.shell32.ShellExecuteW(None, "runas", executable, f'"{script}" {params}', None, 1)
    return ret > 32

def check_windows_version():
    if platform.system() == "Windows":
        try:
            output = subprocess.check_output('wmic os get Caption', shell=True).decode(errors='ignore')
            if "Pro" in output and ("Windows 10" in output or "Windows 11" in output):
                return True
        except:
            pass
    return False

class TelemetryApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MarsTelemetry Control // v0.1")
        self.geometry("720x700")
        self.configure(bg="#1e1e1e")
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        self.style.configure(".", background="#1e1e1e", foreground="white", font=("Segoe UI", 10))
        self.style.configure("TButton", background="#333", foreground="white", padding=6, relief="flat")
        self.style.map("TButton", background=[("active", "#444")])
        self.style.configure("TLabel", background="#1e1e1e", foreground="white")
        self.style.configure("TCheckbutton", background="#1e1e1e", foreground="white")
        self.style.configure("TFrame", background="#1e1e1e")

        self.translations = {
            "RU": {
                "title": "Отключение телеметрии Windows и Microsoft Office",
                "disable": "Отключить телеметрию",
                "restore": "Восстановить по умолчанию",
                "status": "Текущее состояние: не определено",
                "options": "Опции телеметрии",
                "log": "Журнал действий",
                "lang_btn": "RU / EN",
                "confirm_disable": "Вы уверены, что хотите отключить телеметрию?",
                "confirm_restore": "Восстановить телеметрию по умолчанию?",
                "Disable DiagTrack service": "Отключить службу DiagTrack",
                "Disable dmwappushservice": "Отключить dmwappushservice",
                "Disable Diagnostic Task": "Отключить задачи диагностики",
                "Disable Office Telemetry": "Отключить телеметрию Office",
                "Disable LinkedIn info": "Отключить отображение LinkedIn",
                "restore_point_recommendation": "Рекомендуется создать точку восстановления системы перед изменениями.",
                "create_restore_point": "Создать точку восстановления"
            },
            "EN": {
                "title": "Disable Windows and Microsoft Office Telemetry",
                "disable": "Disable Telemetry",
                "restore": "Restore Defaults",
                "status": "Current status: unknown",
                "options": "Telemetry Options",
                "log": "Action Log",
                "lang_btn": "RU / EN",
                "confirm_disable": "Are you sure you want to disable telemetry?",
                "confirm_restore": "Restore telemetry to defaults?",
                "Disable DiagTrack service": "Disable DiagTrack service",
                "Disable dmwappushservice": "Disable dmwappushservice",
                "Disable Diagnostic Task": "Disable Diagnostic Task",
                "Disable Office Telemetry": "Disable Office Telemetry",
                "Disable LinkedIn info": "Disable LinkedIn info",
                "restore_point_recommendation": "It is recommended to create a system restore point before making changes.",
                "create_restore_point": "Create Restore Point"
            }
        }
        self.current_lang = "RU"
        self.option_vars = {}
        self.checkbuttons = {}
        self.create_widgets()
        self.set_language(self.current_lang)

        messagebox.showinfo(
            self.translations[self.current_lang]["title"],
            self.translations[self.current_lang]["restore_point_recommendation"]
        )

    def create_widgets(self):
        self.header = ttk.Label(self, font=("Segoe UI", 14, "bold"))
        self.header.pack(pady=(20, 10))

        self.status_var = tk.StringVar()
        self.status_label = ttk.Label(self, textvariable=self.status_var, font=("Segoe UI", 11, "italic"))
        self.status_label.pack(pady=(0, 15))

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10)

        self.disable_btn = ttk.Button(btn_frame, command=self.confirm_disable)
        self.disable_btn.grid(row=0, column=0, padx=10, pady=5)

        self.restore_btn = ttk.Button(btn_frame, command=self.confirm_restore)
        self.restore_btn.grid(row=0, column=1, padx=10, pady=5)

        self.restore_point_btn = ttk.Button(btn_frame, command=self.open_restore_point)
        self.restore_point_btn.grid(row=0, column=2, padx=10, pady=5)

        self.options_frame = ttk.LabelFrame(self)
        self.options_frame.pack(pady=20, fill=tk.X, padx=30)

        option_keys = [
            "Disable DiagTrack service",
            "Disable dmwappushservice",
            "Disable Diagnostic Task",
            "Disable Office Telemetry",
            "Disable LinkedIn info"
        ]

        for key in option_keys:
            var = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(self.options_frame, text="", variable=var)
            cb.pack(anchor='w', padx=10, pady=3)
            self.option_vars[key] = var
            self.checkbuttons[key] = cb

        self.log_frame = ttk.LabelFrame(self)
        self.log_frame.pack(padx=30, pady=10, fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(self.log_frame, height=10, bg="#121212", fg="#dcdcdc", insertbackground="white", relief="flat", wrap="word")
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        bottom_frame = ttk.Frame(self)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 20), padx=30)

        bottom_frame.columnconfigure(0, weight=1)
        bottom_frame.columnconfigure(1, weight=1)

        self.lang_btn = ttk.Button(bottom_frame, command=self.toggle_language)
        self.lang_btn.grid(row=0, column=0, sticky="w")

        social_frame = ttk.Frame(bottom_frame)
        social_frame.grid(row=0, column=1, sticky="e")

        github_btn = ttk.Button(social_frame, text="GitHub", command=lambda: webbrowser.open("https://github.com/yourgithub"))
        github_btn.pack(side=tk.LEFT, padx=(0, 5))

        telegram_btn = ttk.Button(social_frame, text="Telegram", command=lambda: webbrowser.open("https://t.me/yourtelegram"))
        telegram_btn.pack(side=tk.LEFT)

    def open_restore_point(self):
        try:
            subprocess.Popen("SystemPropertiesProtection.exe")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть раздел восстановления:\n{e}")

    def log(self, text):
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.see(tk.END)

    def update_status(self):
        self.status_var.set(self.translations[self.current_lang]["status"])

    def confirm_disable(self):
        if messagebox.askyesno("", self.translations[self.current_lang]["confirm_disable"]):
            messagebox.showinfo(
                self.translations[self.current_lang]["title"],
                self.translations[self.current_lang]["restore_point_recommendation"]
            )
            self.disable_telemetry()

    def confirm_restore(self):
        if messagebox.askyesno("", self.translations[self.current_lang]["confirm_restore"]):
            self.restore_telemetry()

    def disable_telemetry(self):
        self.log("Отключение телеметрии...")
        options = {k: v.get() for k, v in self.option_vars.items()}
        telemetry_control.disable_telemetry(log_callback=self.log, options=options)
        self.log("Отключение завершено.")
        self.update_status()

    def restore_telemetry(self):
        self.log("Восстановление параметров...")
        options = {k: v.get() for k, v in self.option_vars.items()}
        telemetry_control.restore_telemetry(log_callback=self.log, options=options)
        self.log("Восстановление завершено.")
        self.update_status()

    def toggle_language(self):
        self.set_language("EN" if self.current_lang == "RU" else "RU")

    def set_language(self, lang):
        t = self.translations[lang]
        self.header.config(text=t["title"])
        self.disable_btn.config(text=t["disable"])
        self.restore_btn.config(text=t["restore"])
        self.restore_point_btn.config(text=t["create_restore_point"])
        self.status_var.set(t["status"])
        self.options_frame.config(text=t["options"])
        self.log_frame.config(text=t["log"])
        self.lang_btn.config(text=t["lang_btn"])
        for key, cb in self.checkbuttons.items():
            cb.config(text=t[key])
        self.current_lang = lang

def main():
    if not is_admin():
        if run_as_admin():
            sys.exit(0)
        else:
            messagebox.showerror("Ошибка", "Требуются права администратора.")
            sys.exit(1)

    if not check_windows_version():
        messagebox.showerror("Ошибка", "Только для Windows 10/11 Pro.")
        sys.exit(1)

    app = TelemetryApp()
    app.mainloop()

if __name__ == "__main__":
    main()