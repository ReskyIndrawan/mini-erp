import tkinter as tk
from tkinter import ttk
from tabs.tab1_template import Tab1Template
from tabs.tab2_entry import Tab2Entry


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("不良品データ 管理システム")
        self.geometry("1500x800")

        # pasang icon (icon.ico ada di root folder)
        self.iconbitmap("icon.ico")

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
