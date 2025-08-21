import tkinter as tk
from tkinter import filedialog, messagebox
import os, datetime
from excel_utils import create_excel_if_not_exists, append_excel


class Tab1Template:
    def __init__(self, parent, app):
        self.app = app
        frame = tk.Frame(parent)
        self.frame = frame

        tk.Label(frame, text="作成者").grid(
            row=0, column=0, sticky="w", pady=10, padx=10
        )
        self.entry_creator = tk.Entry(frame, width=30)
        self.entry_creator.grid(row=0, column=1, padx=10)

        tk.Label(frame, text="保存先ディレクトリ").grid(
            row=1, column=0, sticky="w", pady=10, padx=10
        )
        self.dir_label = tk.Label(
            frame, text="（未選択）", width=40, anchor="w", relief="sunken"
        )
        self.dir_label.grid(row=1, column=1, padx=10)
        tk.Button(frame, text="参照", command=self.select_dir).grid(row=1, column=2)

        self.btn_save = tk.Button(frame, text="Excel作成", command=self.generate_excel)
        self.btn_save.grid(row=2, column=0, columnspan=3, pady=20)

        self.frame.pack(fill="both", expand=True)

    def select_dir(self):
        folder = filedialog.askdirectory(title="保存先ディレクトリを選択")
        if folder:
            self.app.base_dir = folder
            self.dir_label.config(text=folder)

    def generate_excel(self):
        creator = self.entry_creator.get().strip()
        if not creator:
            messagebox.showerror("エラー", "作成者は必須項目です。")
            return
        if not self.app.base_dir:
            messagebox.showerror("エラー", "保存先ディレクトリを選択してください。")
            return
        self.app.creator = creator

        today = datetime.date.today()
        foldername = f"{today.year}-{today.month}-不良品データ"
        excel_folder = os.path.join(self.app.base_dir, foldername)
        os.makedirs(excel_folder, exist_ok=True)
        os.makedirs(os.path.join(excel_folder, "不良発生連絡書発行"), exist_ok=True)
        filepath = create_excel_if_not_exists(excel_folder, creator)

        # ダミー行を追加
        append_excel(
            excel_folder,
            [
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
            ],
            creator,
        )
        messagebox.showinfo("成功", f"Excelファイルが作成されました：\n{filepath}")
