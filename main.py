import os, sys
import tkinter as tk
from tkinter import ttk
from tabs.tab1_template import Tab1Template
from tabs.tab2_entry import Tab2Entry


def resource_path(relative_path: str) -> str:
    # Mendukung PyInstaller (sys._MEIPASS) dan run dari source
    base_path = getattr(sys, "_MEIPASS", None) or os.path.dirname(
        os.path.abspath(__file__)
    )
    return os.path.join(base_path, relative_path)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("不良品データ 管理システム")
        self.geometry("1500x800")

        # Pasang icon dengan aman (fallback jika tidak ada)
        try:
            ico_path = resource_path("icon.ico")
            if os.path.exists(ico_path):
                self.iconbitmap(ico_path)
            else:
                # fallback ke default; jangan crash
                self.iconbitmap("")  # boleh juga di-skip
        except Exception:
            # Jika platform non-Windows atau format tidak sesuai, diamkan saja
            pass

        self.base_dir = None
        self.creator = ""

        tabControl = ttk.Notebook(self)
        tab1 = ttk.Frame(tabControl)
        tab2 = ttk.Frame(tabControl)
        tabControl.add(tab1, text="テンプレート生成")
        tabControl.add(tab2, text="不良品データ入力")
        tabControl.pack(expand=1, fill="both")

        self.tab1_ui = Tab1Template(tab1, self)
        self.tab2_ui = Tab2Entry(tab2, self)


if __name__ == "__main__":
    app = App()
    app.mainloop()
