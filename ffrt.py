import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from fontTools.ttLib import TTFont, TTLibFileIsCollectionError
import threading
from queue import Queue, Empty

class FontRenamer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("字体文件重命名工具")
        self.geometry("1200x600")
        self.file_paths = []  # 保存字体文件路径的列表
        self.queue = Queue()
        self.create_widgets()
        self.after(100, self.process_queue)
        self.tree.tag_configure("error", background="red")

    def create_widgets(self):
        # 创建菜单栏
        menubar = tk.Menu(self)
        self.config(menu=menubar)

        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="使用说明", command=self.show_instructions)
        help_menu.add_command(label="版本信息", command=self.show_version_info)

        # 创建表格框架
        table_frame = tk.Frame(self)
        table_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # 创建 Treeview 和滚动条
        self.tree = ttk.Treeview(table_frame, columns=("Index", "CurrentName", "NewName", "Family", "Subfamily", "FullName", "Version", "PostScriptName", "Manufacturer", "Designer", "Description", "URLVendor", "URLDesigner", "License", "LicenseURL", "SampleText"), show='headings')

        # 添加竖向滚动条
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)

        # 配置表头
        self.tree.heading("Index", text="序号", anchor='center')
        self.tree.heading("CurrentName", text="现有名字", anchor='center')
        self.tree.heading("NewName", text="重命名的名字", anchor='center')
        self.tree.heading("Family", text="家族", anchor='center')
        self.tree.heading("Subfamily", text="子家族", anchor='center')
        self.tree.heading("FullName", text="全名", anchor='center')
        self.tree.heading("Version", text="版本", anchor='center')
        self.tree.heading("PostScriptName", text="PostScript名称", anchor='center')
        self.tree.heading("Manufacturer", text="制造商", anchor='center')
        self.tree.heading("Designer", text="设计师", anchor='center')
        self.tree.heading("Description", text="描述", anchor='center')
        self.tree.heading("URLVendor", text="供应商URL", anchor='center')
        self.tree.heading("URLDesigner", text="设计师URL", anchor='center')
        self.tree.heading("License", text="许可", anchor='center')
        self.tree.heading("LicenseURL", text="许可URL", anchor='center')
        self.tree.heading("SampleText", text="样本文本", anchor='center')

        # 配置列宽
        self.tree.column("Index", width=50, anchor='center')
        self.tree.column("CurrentName", width=200, anchor='w')
        self.tree.column("NewName", width=300, anchor='w')
        self.tree.column("Family", width=150, anchor='w')
        self.tree.column("Subfamily", width=150, anchor='w')
        self.tree.column("FullName", width=200, anchor='w')
        self.tree.column("Version", width=100, anchor='w')
        self.tree.column("PostScriptName", width=200, anchor='w')
        self.tree.column("Manufacturer", width=150, anchor='w')
        self.tree.column("Designer", width=150, anchor='w')
        self.tree.column("Description", width=250, anchor='w')
        self.tree.column("URLVendor", width=200, anchor='w')
        self.tree.column("URLDesigner", width=200, anchor='w')
        self.tree.column("License", width=150, anchor='w')
        self.tree.column("LicenseURL", width=200, anchor='w')
        self.tree.column("SampleText", width=250, anchor='w')

        # 将 Treeview 和滚动条放置在同一个框架中
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # 创建按钮框架
        button_frame = tk.Frame(self)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)

        self.load_files_btn = tk.Button(button_frame, text="添加文件", command=self.load_files)
        self.load_files_btn.pack(side=tk.LEFT, padx=10)

        self.load_folder_btn = tk.Button(button_frame, text="添加文件夹", command=self.load_folder)
        self.load_folder_btn.pack(side=tk.LEFT, padx=10)

        self.rename_btn = tk.Button(button_frame, text="重命名选中的文件", command=self.rename_selected_files, state=tk.DISABLED)
        self.rename_btn.pack(side=tk.LEFT, padx=10)

        self.clear_btn = tk.Button(button_frame, text="清空列表", command=self.clear_table)
        self.clear_btn.pack(side=tk.LEFT, padx=10)

        self.info_btn = tk.Button(button_frame, text="信息", command=self.show_instructions)
        self.info_btn.pack(side=tk.RIGHT, padx=10)

        self.progress_var = tk.DoubleVar()
        self.progress = ttk.Progressbar(button_frame, orient="horizontal", length=200, mode="determinate", variable=self.progress_var)
        self.progress.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

        # 初始化包含列状态的字典
        self.column_visibility = {
            "Family": False, "Subfamily": False, "FullName": False, "Version": False, 
            "PostScriptName": False, "Manufacturer": False, "Designer": False, 
            "Description": False, "URLVendor": False, "URLDesigner": False, 
            "License": False, "LicenseURL": False, "SampleText": False
        }

        self.combobox = ttk.Combobox(self.tree)
        self.combobox.bind("<<ComboboxSelected>>", self.update_new_name)
        self.combobox.bind("<FocusOut>", self.on_focus_out)
        self.tree.bind("<Double-1>", self.on_double_click)

    def show_instructions(self):
        messagebox.showinfo("使用说明", "这是一个字体文件重命名工具，您可以添加字体文件或文件夹，然后选择要重命名的文件并输入新名称。\n\n使用方法：\n1. 点击'添加文件'或'添加文件夹'按钮加载字体文件。\n2. 双击'重命名的名字'列，选择新的文件名。\n3. 点击'重命名选中的文件'按钮完成重命名。\n4. 点击'清空列表'按钮清空所有加载的字体文件。\5. 双击'重命名的名字'列，可以选择和编辑新的文件名。\n6. 点击'信息'按钮查看使用说明和版本信息。")

    def show_version_info(self):
        messagebox.showinfo("版本信息", "字体文件重命名工具\n版本 1.0.0")

    def load_files(self):
        file_paths = filedialog.askopenfilenames(title="选择字体文件", filetypes=[("字体文件", "*.ttf;*.otf;*.ttc")])
        self.start_background_task(file_paths)

    def load_folder(self):
        folder_path = filedialog.askdirectory(title="选择文件夹")
        if folder_path:
            file_paths = []
            for root, _, files in os.walk(folder_path):
                for file in files:
                    if file.lower().endswith(('.ttf', '.otf', '.ttc')):
                        file_paths.append(os.path.join(root, file))
            self.start_background_task(file_paths)

    def start_background_task(self, file_paths):
        threading.Thread(target=self.process_files, args=(file_paths,)).start()

    def process_files(self, file_paths):
        total_files = len(file_paths)
        if total_files == 0:
            self.queue.put("DONE")
            return
        
        for idx, file_path in enumerate(file_paths):
            try:
                font_names = self.get_font_names(file_path)
                if font_names:
                    self.queue.put((file_path, font_names))
                else:
                    self.queue.put((file_path, []))  # 添加一个空的字体名称列表到队列中
            except Exception as e:
                messagebox.showerror("错误", f"无法处理文件 {file_path}：\n{e}")
                self.queue.put((file_path, []))  # 添加一个空的字体名称列表到队列中
            
            self.update_progress_bar(idx + 1, total_files)
        
        self.queue.put("DONE")

    def process_queue(self):
        try:
            while True:
                task = self.queue.get_nowait()
                if task == "DONE":
                    self.update_rename_button_state()
                    self.hide_unused_columns()
                    return
                elif isinstance(task, tuple) and task[0] == 'progress':
                    self.progress_var.set(task[1])
                else:
                    file_path, font_names = task
                    self.file_paths.append(file_path)
                    self.add_file_to_table(file_path, font_names)
        except Empty:
            pass
        self.after(100, self.process_queue)

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
            return chinese_names + other_names
        except TTLibFileIsCollectionError:
            return self.get_font_names_from_collection(file_path)

    def is_chinese(self, string):
        for char in string:
            if '\u4e00' <= char <= '\u9fff':
                return True
        return False

    def get_font_names_from_collection(self, file_path):
        font_names = []
        try:
            ttc = TTFont(file_path, fontNumber=0)  # 打开字体集合中的第一个字体
            for i in range(ttc.reader.numFonts):
                try:
                    font = TTFont(file_path, fontNumber=i)
                    full_name = ""
                    family_name = ""
                    for record in font['name'].names:
                        if record.nameID == 4:
                            full_name = record.toUnicode()
                        elif record.nameID == 1:
                            family_name = record.toUnicode()
                    font.close()
                    if full_name:
                        font_names.append(full_name)
                    elif family_name:
                        font_names.append(family_name)
                except Exception as e:
                    messagebox.showerror("错误", f"无法读取字体集合中的字体：\n{e}")
        except Exception as e:
            messagebox.showerror("错误", f"无法打开字体集合：\n{e}")
        return font_names

    def add_file_to_table(self, file_path, font_names):
        current_name = os.path.basename(file_path)
        
        # 获取所有可能的字体名称
        all_chinese_names = [name.strip() for name in font_names if self.is_chinese(name.strip())]
        all_other_names = [name.strip() for name in font_names if not self.is_chinese(name.strip())]
        all_names = all_chinese_names + all_other_names

        # 默认展示第一个中文字体名称，如果没有中文字体名称则展示第一个字体名称
        new_name = all_chinese_names[0] if all_chinese_names else all_other_names[0] if all_other_names else ""
        details = self.get_font_details(file_path, new_name)

        # 更新列可见性状态
        for i, key in enumerate(self.column_visibility.keys(), start=3):
            if details[i - 3]:
                self.column_visibility[key] = True

        values = (len(self.file_paths), current_name, new_name, *details)
        item_id = self.tree.insert("", "end", values=values)

        # 如果存在无法处理的字体文件或者出现错误的字体文件，则设置该行为红色背景
        if not all_names or not details:
            self.tree.item(item_id, tags=("error",))  # 在这里添加 "error" 标签

    def on_double_click(self, event):
        item = self.tree.selection()[0]
        column = self.tree.identify_column(event.x)
        if column == "#3":  # Check if the column is the "NewName" column
            values = self.tree.item(item, "values")[2:]
            all_names = [name for name in values if name.strip()]
            all_names = sorted(all_names, key=self.is_chinese, reverse=True)
            self.combobox['values'] = all_names
            x, y, width, height = self.tree.bbox(item, column)
            self.combobox.place(x=x, y=y, width=width, height=height)
            self.combobox.focus_set()
            current_value = self.tree.set(item, "NewName")
            self.combobox.set(current_value)

    def update_new_name(self, event):
        combobox = event.widget
        new_name = combobox.get()
        selected_item = self.tree.selection()[0]
        values = list(self.tree.item(selected_item, 'values'))
        values[2] = new_name  # Update the "NewName" value
        self.tree.item(selected_item, values=values)
        combobox.place_forget()

    def on_focus_out(self, event):
        combobox = event.widget
        selected_item = self.tree.selection()[0]
        values = list(self.tree.item(selected_item, 'values'))
        new_name = combobox.get()
        values[2] = new_name
        self.tree.item(selected_item, values=values)
        combobox.place_forget()

    def get_font_details(self, file_path, full_name):
        try:
            font = TTFont(file_path)
            family = self.get_name_record(font, 1)
            subfamily = self.get_name_record(font, 2)
            version = self.get_name_record(font, 5)
            postscript_name = self.get_name_record(font, 6)
            manufacturer = self.get_name_record(font, 8)
            designer = self.get_name_record(font, 9)
            description = self.get_name_record(font, 10)
            url_vendor = self.get_name_record(font, 11)
            url_designer = self.get_name_record(font, 12)
            license = self.get_name_record(font, 13)
            license_url = self.get_name_record(font, 14)
            sample_text = self.get_name_record(font, 19)
            font.close()
            return [family, subfamily, full_name, version, postscript_name, manufacturer, designer, description, url_vendor, url_designer, license, license_url, sample_text]
        except Exception as e:
            messagebox.showerror("错误", f"无法读取字体信息：\n{e}")
            return [""] * 13

    def get_name_record(self, font, nameID):
        for record in font['name'].names:
            if record.nameID == nameID and record.isUnicode():
                return record.toUnicode()
        return ""

    def update_progress_bar(self, current, total):
        self.queue.put(('progress', (current / total) * 100))

    def rename_selected_files(self):
        selected_items = self.tree.selection()
        new_names = [self.tree.item(item, "values")[2] for item in selected_items]
        if len(new_names) != len(set(new_names)):
            messagebox.showerror("错误", "选中的文件名有重复，请修改后再试。")
            return

        for item in selected_items:
            values = self.tree.item(item, "values")
            current_path = next((fp for fp in self.file_paths if os.path.basename(fp) == values[1]), None)
            if current_path:
                new_name = values[2]
                new_path = os.path.join(os.path.dirname(current_path), new_name)
                os.rename(current_path, new_path)
                self.file_paths.remove(current_path)
                self.file_paths.append(new_path)
                self.tree.item(item, values=(values[0], new_name, new_name, *values[3:]))

    def clear_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.file_paths = []
        self.update_rename_button_state()
        self.progress_var.set(0)

    def update_rename_button_state(self):
        if self.tree.get_children():
            self.rename_btn.config(state=tk.NORMAL)
        else:
            self.rename_btn.config(state=tk.DISABLED)

    def hide_unused_columns(self):
        for column, visible in self.column_visibility.items():
            if not visible:
                self.tree.column(column, width=0, stretch=tk.NO)

    def process_queue(self):
        while not self.queue.empty():
            try:
                message = self.queue.get_nowait()
                if message == "DONE":
                    self.update_rename_button_state()
                    self.hide_unused_columns()
                elif isinstance(message, tuple) and message[0] == 'progress':
                    self.progress_var.set(message[1])
                else:
                    file_path, font_names = message
                    self.file_paths.append(file_path)
                    self.add_file_to_table(file_path, font_names)
            except Empty:
                pass
        self.after(100, self.process_queue)

if __name__ == "__main__":
    app = FontRenamer()
    app.mainloop()
