import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from fontTools.ttLib import TTFont, TTLibFileIsCollectionError


class FontRenamer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("字体文件重命名工具")
        self.geometry("1000x600")
        self.file_paths = []
        self.checkbox_vars = []
        self.new_name_options = []
        self.details_vars = []
        self.progress_var = tk.DoubleVar()
        self.create_widgets()

    def create_widgets(self):
        title_frame = tk.Frame(self)
        title_frame.pack(side=tk.TOP, fill=tk.X, pady=5)

        self.select_all_var = tk.IntVar()
        select_all_checkbox = tk.Checkbutton(
            title_frame, text="全选", variable=self.select_all_var, command=self.toggle_all_checkboxes
        )
        select_all_checkbox.pack(side=tk.LEFT, padx=5)

        current_name_label = tk.Label(title_frame, text="现有名字")
        current_name_label.pack(side=tk.LEFT, padx=5)

        rename_name_label = tk.Label(title_frame, text="重命名的名字")
        rename_name_label.pack(side=tk.LEFT, padx=5)

        details_label = tk.Label(title_frame, text="详情")
        details_label.pack(side=tk.LEFT, padx=5)

        table_frame = tk.Frame(self)
        table_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(table_frame)
        self.scrollbar = tk.Scrollbar(
            table_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window(
            (0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 支持鼠标滚轮滚动
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        button_frame = tk.Frame(self)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)

        self.load_files_btn = tk.Button(
            button_frame, text="添加文件", command=self.load_files
        )
        self.load_files_btn.pack(side=tk.LEFT, padx=10)

        self.load_folder_btn = tk.Button(
            button_frame, text="添加文件夹", command=self.load_folder
        )
        self.load_folder_btn.pack(side=tk.LEFT, padx=10)

        self.rename_btn = tk.Button(
            button_frame, text="重命名选中的文件", command=self.rename_selected_files, state=tk.DISABLED
        )
        self.rename_btn.pack(side=tk.LEFT, padx=10)

        self.clear_btn = tk.Button(
            button_frame, text="清空列表", command=self.clear_table
        )
        self.clear_btn.pack(side=tk.LEFT, padx=10)

        self.progress = ttk.Progressbar(
            button_frame, orient="horizontal", length=200, mode="determinate", variable=self.progress_var, maximum=100
        )
        self.progress.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

        # 设置进度条颜色
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("TProgressbar",
                        thickness=10,
                        troughcolor='gray',
                        bordercolor='gray',
                        background='green',
                        lightcolor='green',
                        darkcolor='green')

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def toggle_all_checkboxes(self):
        for checkbox_var in self.checkbox_vars:
            checkbox_var.set(self.select_all_var.get())
        self.update_rename_button_state()

    def update_rename_button_state(self):
        if any(checkbox_var.get() for checkbox_var in self.checkbox_vars):
            self.rename_btn.config(state=tk.NORMAL)
        else:
            self.rename_btn.config(state=tk.DISABLED)

    def load_files(self):
        file_paths = filedialog.askopenfilenames(
            title="选择字体文件", filetypes=[("字体文件", "*.ttf;*.otf;*.ttc")]
        )
        self.process_files(file_paths)

    def load_folder(self):
        folder_path = filedialog.askdirectory(title="选择文件夹")
        if folder_path:
            file_paths = []
            for root, _, files in os.walk(folder_path):
                for file in files:
                    if file.lower().endswith(('.ttf', '.otf', '.ttc')):
                        file_paths.append(os.path.join(root, file))
            self.process_files(file_paths)

    def process_files(self, file_paths):
        total_files = len(file_paths)
        for idx, file_path in enumerate(file_paths):
            if file_path in self.file_paths:
                continue  # 忽略已经加载的文件
            font_names = self.get_font_names(file_path)
            if font_names:
                self.file_paths.append(file_path)
                self.add_file_to_table(file_path, font_names)
            self.update_progress_bar(idx + 1, total_files)
        self.update_rename_button_state()

    def update_progress_bar(self, current, total):
        self.progress_var.set((current / total) * 100)
        self.update_idletasks()


    def get_font_names(self, file_path):
        try:
            font = TTFont(file_path)
            chinese_names = []
            other_names = []
            for record in font['name'].names:
                name = record.toUnicode()
                if name and record.isUnicode():
                    if self.is_chinese(name):
                        chinese_names.append(name)
                    else:
                        other_names.append(name)
            font.close()
            if chinese_names:
                return chinese_names
            elif other_names:
                return other_names
            else:
                return [os.path.splitext(os.path.basename(file_path))[0]]
        except TTLibFileIsCollectionError:
            return self.get_font_names_from_collection(file_path)


    def is_chinese(self, string):
        """
        检查字符串是否包含中文字符
        """
        for char in string:
            if '\u4e00' <= char <= '\u9fff':
                return True
        return False

    def get_font_names_from_collection(self, file_path):
        font_names = []
        ttc = TTFont(file_path)
        for i in range(ttc.reader.numFonts):
            try:
                font = TTFont(file_path, fontNumber=i)
                full_name = ""
                family_name = ""
                for record in font['name'].names:
                    if record.nameID == 4:  # Full name
                        full_name = record.toUnicode()
                    elif record.nameID == 1:  # Family name
                        family_name = record.toUnicode()
                font.close()
                if full_name:
                    font_names.append(full_name)
                elif family_name:
                    font_names.append(family_name)
            except Exception as e:
                messagebox.showerror("错误", f"无法读取字体集合中的字体：\n{e}")
        return font_names

    def add_file_to_table(self, file_path, font_names):
        row = len(self.file_paths) - 1
        checkbox_var = tk.IntVar()
        self.checkbox_vars.append(checkbox_var)
        checkbox = tk.Checkbutton(
            self.scrollable_frame, variable=checkbox_var, command=self.update_rename_button_state)
        checkbox.grid(row=row, column=0, sticky='w', padx=5, pady=2)

        current_name_label = tk.Label(
            self.scrollable_frame, text=os.path.basename(file_path))
        current_name_label.grid(row=row, column=1, sticky='w', padx=5, pady=2)

        new_name_var = tk.StringVar(value=font_names[0])
        self.new_name_options.append(new_name_var)
        new_name_dropdown = ttk.Combobox(
            self.scrollable_frame, textvariable=new_name_var, values=font_names)
        new_name_dropdown.grid(row=row, column=2, sticky='ew', padx=5, pady=2)
        new_name_dropdown.bind("<<ComboboxSelected>>", lambda e,
                               nv=new_name_var, p=file_path: self.update_details(nv, p))

        details_var = tk.StringVar()
        self.details_vars.append(details_var)
        details_label = tk.Label(
            self.scrollable_frame, textvariable=details_var, wraplength=300)
        details_label.grid(row=row, column=3, sticky='w', padx=5, pady=2)

        self.update_details(new_name_var, file_path)

        # 添加行间隔线
        if row > 0:
            tk.Frame(self.scrollable_frame, height=1, bg='gray').grid(
                row=row*2, column=0, columnspan=4, sticky='ew')

    def update_details(self, new_name_var, file_path):
        selected_name = new_name_var.get()
        details = self.get_font_details(file_path, selected_name)
        details_index = self.new_name_options.index(new_name_var)
        self.details_vars[details_index].set(details)

    def get_font_details(self, file_path, font_name):
        try:
            font = TTFont(file_path)
            for record in font['name'].names:
                if font_name == record.toUnicode():
                    family = self.get_name_record(font, 1)
                    style = self.get_name_record(font, 2)
                    version = self.get_name_record(font, 5)
                    return f"家族: {family}, 风格: {style}, 版本: {version}"
            font.close()
        except Exception as e:
            messagebox.showerror("错误", f"无法获取字体详情 {file_path}:\n{e}")
        return "无详情"

    def get_name_record(self, font, nameID):
        for record in font['name'].names:
            if record.nameID == nameID:
                return record.toUnicode()
        return "无"

    def rename_selected_files(self):
        selected_files = []
        renamed_files = []
        for checkbox_var, file_path, new_name_var in zip(self.checkbox_vars, self.file_paths, self.new_name_options):
            if checkbox_var.get():
                new_name = new_name_var.get().strip()
                if not new_name:
                    messagebox.showwarning("警告", "新文件名不能为空！")
                    continue
                selected_files.append((file_path, new_name))

        if not selected_files:
            messagebox.showinfo("提示", "请选择要重命名的字体文件！")
            return

        total_files = len(selected_files)
        for idx, (file_path, new_name) in enumerate(selected_files):
            dir_path, old_name = os.path.split(file_path)
            _, ext = os.path.splitext(old_name)
            new_file_path = os.path.join(dir_path, f"{new_name}{ext}")

            try:
                counter = 1
                while os.path.exists(new_file_path):
                    new_file_path = os.path.join(
                        dir_path, f"{new_name}_{counter}{ext}")
                    counter += 1
                os.rename(file_path, new_file_path)
                renamed_files.append(
                    (old_name, os.path.basename(new_file_path)))
            except Exception as e:
                messagebox.showerror("错误", f"无法重命名文件 {file_path}:\n{e}")

            self.update_progress_bar(idx + 1, total_files)

        if renamed_files:
            message = "以下文件已成功重命名：\n"
            for old_name, new_name in renamed_files:
                message += f"{old_name} -> {new_name}\n"
            messagebox.showinfo("重命名成功", message)
        self.clear_table()

    def clear_table(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.file_paths = []
        self.checkbox_vars = []
        self.new_name_options = []
        self.update_rename_button_state()
        self.progress_var.set(0)


if __name__ == "__main__":
    app = FontRenamer()
    app.mainloop()
