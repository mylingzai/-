import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import random
import json
import os
from datetime import datetime
import threading
import time
from functools import lru_cache

class CheckboxTreeview(ttk.Treeview):
    def __init__(self, master=None, **kwargs):
        if 'height' in kwargs:
            kwargs.pop('height')
        super().__init__(master, **kwargs)
        self.checked_items = set()
        self.bind('<Button-1>', self.on_click)
        
        self.checkbox_images = {
            "checked": self.create_checkbox_image(True),
            "unchecked": self.create_checkbox_image(False)
        }
    
    def create_checkbox_image(self, checked):
        image = tk.PhotoImage(width=16, height=16)
        
        if checked:
            image.put("green", to=(0, 0, 15, 15))
            image.put("white", to=(2, 2, 13, 13))
            image.put("green", to=(4, 4, 11, 11))
        else:
            image.put("black", to=(0, 0, 15, 15))
            image.put("white", to=(1, 1, 14, 14))
        
        return image
    
    def on_click(self, event):
        item = self.identify_row(event.y)
        column = self.identify_column(event.x)
        
        if item and column == "#0":
            try:
                if item in self.checked_items:
                    self.change_state(item, "unchecked")
                else:
                    self.change_state(item, "checked")
            except tk.TclError:
                pass
    
    def change_state(self, item, state):
        try:
            if state == "checked":
                self.checked_items.add(item)
                self.item(item, image=self.checkbox_images["checked"])
            else:
                if item in self.checked_items:
                    self.checked_items.remove(item)
                self.item(item, image=self.checkbox_images["unchecked"])
        except tk.TclError:
            if item in self.checked_items:
                self.checked_items.remove(item)
    
    def get_checked_items(self):
        checked_names = []
        items_to_check = self.checked_items.copy()
        
        for item in items_to_check:
            try:
                self.item(item)
                checked_names.append(self.item(item, "text"))
            except tk.TclError:
                if item in self.checked_items:
                    self.checked_items.remove(item)
        
        return checked_names
    
    def insert(self, parent, index, iid=None, **kw):
        item = super().insert(parent, index, iid, **kw)
        
        try:
            self.item(item, image=self.checkbox_images["unchecked"])
        except tk.TclError:
            pass
        
        return item
    
    def clear_all_checks(self):
        self.checked_items.clear()

class EnhancedLotterySystem:
    def __init__(self, root):
        self.root = root
        self.root.title("增强版终极抽签系统")
        self.root.geometry("1200x900")
        self.root.minsize(1000, 700)
        
        self.setup_styles()
        
        self.show_animation = tk.BooleanVar(value=True)
        self.animation_speed = tk.StringVar(value="中速")
        self.lottery_mode = tk.StringVar(value="常规")
        self.num_to_select = tk.StringVar(value="3")
        
        self.all_students = []
        self.selected_students = []
        self.current_round = 1
        self.last_round_unselected = []
        self.lottery_history = []
        self.is_animating = False
        self.student_groups = {}
        self.student_weights = {}
        self.import_history = []
        self.animation_window = None
        self.auto_backup = True
        self.data_version = "2.1"
        self.backup_count = 0
        self.max_backups = 10
        
        self.config_file = "lottery_config.json"
        
        self.create_widgets()
        
        self.load_config()
        
        self.load_unselected()
        
        self.update_statistics()
        
        self.setup_autosave()
    
    def setup_styles(self):
        style = ttk.Style()
        style.configure("Custom.TFrame", background="#f5f5f5")
        style.configure("Title.TLabel", font=("Arial", 16, "bold"), background="#f5f5f5")
        style.configure("Subtitle.TLabel", font=("Arial", 10, "bold"), background="#f5f5f5")
        
    def setup_autosave(self):
        def autosave():
            if hasattr(self, 'auto_save') and self.auto_save:
                try:
                    self.save_unselected()
                except:
                    pass
            self.root.after(30000, autosave)
        
        self.root.after(30000, autosave)
    
    def load_config(self):
        default_config = {
            'auto_backup': True,
            'auto_save': True,
            'show_animation': True,
            'animation_speed': '中速',
            'default_lottery_mode': '常规',
            'max_backups': 10,
            'theme': 'default',
            'recent_files': []
        }
        
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as file:
                    self.config = {**default_config, **json.load(file)}
            else:
                self.config = default_config
        except:
            self.config = default_config
        
        self.apply_config()
    
    def save_config(self):
        try:
            with open(self.config_file, 'w', encoding='utf-8') as file:
                json.dump(self.config, file, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置时出错: {str(e)}")
    
    def apply_config(self):
        self.auto_backup = self.config.get('auto_backup', True)
        self.auto_save = self.config.get('auto_save', True)
        self.show_animation.set(self.config.get('show_animation', True))
        self.animation_speed.set(self.config.get('animation_speed', '中速'))
        self.lottery_mode.set(self.config.get('default_lottery_mode', '常规'))
    
    def create_widgets(self):
        self.create_menu()
        
        main_frame = ttk.Frame(self.root, padding="15", style="Custom.TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        paned_window = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True)
        
        left_frame = ttk.Frame(paned_window, padding="10")
        paned_window.add(left_frame, weight=1)
        
        right_frame = ttk.Frame(paned_window, padding="10")
        paned_window.add(right_frame, weight=2)
        
        title_label = ttk.Label(left_frame, text="清辞版班级抽签系统", style="Title.TLabel")
        title_label.pack(pady=(0, 15))
        
        list_management_frame = ttk.LabelFrame(left_frame, text="名单管理", padding="10")
        list_management_frame.pack(fill=tk.X, pady=(0, 10))
        
        history_frame = ttk.Frame(list_management_frame)
        history_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(history_frame, text="历史名单:").pack(side=tk.LEFT)
        self.history_var = tk.StringVar()
        self.history_combo = ttk.Combobox(history_frame, textvariable=self.history_var, state="readonly")
        self.history_combo.pack(side=tk.LEFT, padx=(5, 10), fill=tk.X, expand=True)
        self.history_combo.bind("<<ComboboxSelected>>", self.on_history_selected)
        
        ttk.Button(history_frame, text="导入名单", command=self.import_students).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(history_frame, text="清空名单", command=self.clear_students).pack(side=tk.LEFT)
        
        search_frame = ttk.Frame(list_management_frame)
        search_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(search_frame, text="搜索:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=20)
        search_entry.pack(side=tk.LEFT, padx=(5, 5), fill=tk.X, expand=True)
        search_entry.bind('<KeyRelease>', self.on_search)
        
        ttk.Button(search_frame, text="清空", command=self.clear_search).pack(side=tk.LEFT)
        
        ttk.Label(list_management_frame, text="学生名单:", style="Subtitle.TLabel").pack(anchor=tk.W, pady=(0, 5))
        
        self.students_text = scrolledtext.ScrolledText(list_management_frame, height=8, wrap=tk.WORD, state='disabled')
        self.students_text.pack(fill=tk.X, pady=(0, 10))
        
        list_btn_frame = ttk.Frame(list_management_frame)
        list_btn_frame.pack(fill=tk.X)
        
        ttk.Button(list_btn_frame, text="手动添加", command=self.add_student_dialog).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(list_btn_frame, text="批量添加", command=self.batch_add_dialog).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(list_btn_frame, text="随机打乱", command=self.shuffle_students).pack(side=tk.LEFT)
        
        quick_frame = ttk.LabelFrame(left_frame, text="快捷操作", padding="5")
        quick_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(quick_frame, text="抽1人", 
                  command=lambda: self.quick_draw_single()).grid(row=0, column=0, padx=2, pady=2, sticky=tk.W+tk.E)
        ttk.Button(quick_frame, text="抽3人", 
                  command=lambda: self.quick_draw_multiple(3)).grid(row=0, column=1, padx=2, pady=2, sticky=tk.W+tk.E)
        ttk.Button(quick_frame, text="抽5人", 
                  command=lambda: self.quick_draw_multiple(5)).grid(row=0, column=2, padx=2, pady=2, sticky=tk.W+tk.E)
        ttk.Button(quick_frame, text="抽10人", 
                  command=lambda: self.quick_draw_multiple(10)).grid(row=0, column=3, padx=2, pady=2, sticky=tk.W+tk.E)
        
        for i in range(4):
            quick_frame.columnconfigure(i, weight=1)
        
        weights_frame = ttk.LabelFrame(left_frame, text="权重设置", padding="10")
        weights_frame.pack(fill=tk.X, pady=(0, 10))
        
        weights_btn_frame = ttk.Frame(weights_frame)
        weights_btn_frame.pack(fill=tk.X)
        
        ttk.Button(weights_btn_frame, text="设置权重", command=self.set_weights_dialog).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(weights_btn_frame, text="重置权重", command=self.reset_weights).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(weights_btn_frame, text="智能平衡", command=self.enable_smart_balance).pack(side=tk.LEFT)
        
        settings_frame = ttk.LabelFrame(left_frame, text="抽签设置", padding="10")
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        
        settings_row1 = ttk.Frame(settings_frame)
        settings_row1.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(settings_row1, text="每轮抽取人数:").pack(side=tk.LEFT)
        ttk.Entry(settings_row1, textvariable=self.num_to_select, width=8).pack(side=tk.LEFT, padx=(5, 20))
        
        ttk.Label(settings_row1, text="当前轮次:").pack(side=tk.LEFT)
        self.round_label = ttk.Label(settings_row1, text="1", font=("Arial", 10, "bold"), foreground="blue")
        self.round_label.pack(side=tk.LEFT, padx=(5, 20))
        
        settings_row2 = ttk.Frame(settings_frame)
        settings_row2.pack(fill=tk.X)
        
        ttk.Label(settings_row2, text="动画速度:").pack(side=tk.LEFT)
        speed_combo = ttk.Combobox(settings_row2, textvariable=self.animation_speed, 
                                  values=["慢速", "中速", "快速"], width=8, state="readonly")
        speed_combo.pack(side=tk.LEFT, padx=(5, 20))
        
        ttk.Label(settings_row2, text="抽签模式:").pack(side=tk.LEFT)
        mode_combo = ttk.Combobox(settings_row2, textvariable=self.lottery_mode,
                                 values=["常规", "权重模式", "公平模式"], width=10, state="readonly")
        mode_combo.pack(side=tk.LEFT, padx=(5, 0))
        
        animation_frame = ttk.Frame(settings_frame)
        animation_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Checkbutton(animation_frame, text="显示独立抽签动画", variable=self.show_animation).pack(side=tk.LEFT)
        
        action_frame = ttk.LabelFrame(left_frame, text="操作", padding="10")
        action_frame.pack(fill=tk.X, pady=(0, 10))
        
        action_btn_frame = ttk.Frame(action_frame)
        action_btn_frame.pack(fill=tk.X)
        
        ttk.Button(action_btn_frame, text="开始抽签", command=self.start_lottery).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(action_btn_frame, text="重置系统", command=self.reset_system).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(action_btn_frame, text="跳过本轮", command=self.skip_round).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(action_btn_frame, text="一键抽取", command=self.quick_draw).pack(side=tk.LEFT)
        
        data_btn_frame = ttk.Frame(action_frame)
        data_btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(data_btn_frame, text="保存结果", command=self.save_results).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(data_btn_frame, text="导出名单", command=self.export_list).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(data_btn_frame, text="查看历史", command=self.show_history).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(data_btn_frame, text="备份管理", command=self.show_backup_manager).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(data_btn_frame, text="系统设置", command=self.show_settings).pack(side=tk.LEFT)
        
        stats_frame = ttk.LabelFrame(left_frame, text="统计信息", padding="10")
        stats_frame.pack(fill=tk.X)
        
        self.stats_label = ttk.Label(stats_frame, text="总人数: 0 | 已抽中: 0 | 未抽中: 0", font=("Arial", 10))
        self.stats_label.pack(anchor=tk.W)
        
        self.progress = ttk.Progressbar(stats_frame, mode='indeterminate')
        
        selected_frame = ttk.LabelFrame(right_frame, text="已抽中学生名单", padding="10")
        selected_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        selected_toolbar = ttk.Frame(selected_frame)
        selected_toolbar.pack(fill=tk.X, pady=(0, 5))
        
        self.selected_count_label = ttk.Label(selected_toolbar, text="已抽中: 0人", font=("Arial", 10, "bold"))
        self.selected_count_label.pack(side=tk.LEFT)
        
        selected_search_frame = ttk.Frame(selected_toolbar)
        selected_search_frame.pack(side=tk.RIGHT, padx=(10, 0))
        
        ttk.Label(selected_search_frame, text="搜索:").pack(side=tk.LEFT)
        self.selected_search_var = tk.StringVar()
        selected_search_entry = ttk.Entry(selected_search_frame, textvariable=self.selected_search_var, width=15)
        selected_search_entry.pack(side=tk.LEFT, padx=(5, 5))
        selected_search_entry.bind('<KeyRelease>', self.on_selected_search)
        
        ttk.Button(selected_search_frame, text="清空", command=self.clear_selected_search).pack(side=tk.LEFT)
        
        batch_btn_frame = ttk.Frame(selected_toolbar)
        batch_btn_frame.pack(side=tk.RIGHT, padx=(10, 0))
        
        ttk.Button(batch_btn_frame, text="移回未抽中", command=self.batch_move_to_unselected).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(batch_btn_frame, text="全选", command=self.select_all_selected).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(batch_btn_frame, text="取消选择", command=self.deselect_all_selected).pack(side=tk.LEFT)
        
        selected_tree_frame = ttk.Frame(selected_frame)
        selected_tree_frame.pack(fill=tk.BOTH, expand=True)
        
        self.selected_tree = CheckboxTreeview(selected_tree_frame, 
                                             columns=("round", "timestamp", "weight"), 
                                             show="tree headings")
        self.selected_tree.heading("#0", text="✓ 姓名")
        self.selected_tree.heading("round", text="抽中轮次")
        self.selected_tree.heading("timestamp", text="抽中时间")
        self.selected_tree.heading("weight", text="权重")
        self.selected_tree.column("#0", width=150, minwidth=120)
        self.selected_tree.column("round", width=80, minwidth=70)
        self.selected_tree.column("timestamp", width=120, minwidth=100)
        self.selected_tree.column("weight", width=60, minwidth=50)
        
        selected_scrollbar = ttk.Scrollbar(selected_tree_frame, orient=tk.VERTICAL, command=self.selected_tree.yview)
        self.selected_tree.configure(yscrollcommand=selected_scrollbar.set)
        
        self.selected_tree.grid(row=0, column=0, sticky='nsew')
        selected_scrollbar.grid(row=0, column=1, sticky='ns')
        
        selected_tree_frame.grid_rowconfigure(0, weight=1)
        selected_tree_frame.grid_columnconfigure(0, weight=1)
        
        unselected_frame = ttk.LabelFrame(right_frame, text="未抽中学生名单", padding="10")
        unselected_frame.pack(fill=tk.BOTH, expand=True)
        
        unselected_toolbar = ttk.Frame(unselected_frame)
        unselected_toolbar.pack(fill=tk.X, pady=(0, 5))
        
        self.unselected_count_label = ttk.Label(unselected_toolbar, text="未抽中: 0人", font=("Arial", 10, "bold"))
        self.unselected_count_label.pack(side=tk.LEFT)
        
        unselected_search_frame = ttk.Frame(unselected_toolbar)
        unselected_search_frame.pack(side=tk.RIGHT, padx=(10, 0))
        
        ttk.Label(unselected_search_frame, text="搜索:").pack(side=tk.LEFT)
        self.unselected_search_var = tk.StringVar()
        unselected_search_entry = ttk.Entry(unselected_search_frame, textvariable=self.unselected_search_var, width=15)
        unselected_search_entry.pack(side=tk.LEFT, padx=(5, 5))
        unselected_search_entry.bind('<KeyRelease>', self.on_unselected_search)
        
        ttk.Button(unselected_search_frame, text="清空", command=self.clear_unselected_search).pack(side=tk.LEFT)
        
        batch_btn_frame2 = ttk.Frame(unselected_toolbar)
        batch_btn_frame2.pack(side=tk.RIGHT, padx=(10, 0))
        
        ttk.Button(batch_btn_frame2, text="标记为已抽中", command=self.batch_move_to_selected).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(batch_btn_frame2, text="全选", command=self.select_all_unselected).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(batch_btn_frame2, text="取消选择", command=self.deselect_all_unselected).pack(side=tk.LEFT)
        
        unselected_tree_frame = ttk.Frame(unselected_frame)
        unselected_tree_frame.pack(fill=tk.BOTH, expand=True)
        
        self.unselected_tree = CheckboxTreeview(unselected_tree_frame, 
                                               columns=("status", "weight"), 
                                               show="tree headings")
        self.unselected_tree.heading("#0", text="✓ 姓名")
        self.unselected_tree.heading("status", text="状态")
        self.unselected_tree.heading("weight", text="权重")
        self.unselected_tree.column("#0", width=150, minwidth=120)
        self.unselected_tree.column("status", width=100, minwidth=80)
        self.unselected_tree.column("weight", width=60, minwidth=50)
        
        unselected_scrollbar = ttk.Scrollbar(unselected_tree_frame, orient=tk.VERTICAL, command=self.unselected_tree.yview)
        self.unselected_tree.configure(yscrollcommand=unselected_scrollbar.set)
        
        self.unselected_tree.grid(row=0, column=0, sticky='nsew')
        unselected_scrollbar.grid(row=0, column=1, sticky='ns')
        
        unselected_tree_frame.grid_rowconfigure(0, weight=1)
        unselected_tree_frame.grid_columnconfigure(0, weight=1)
        
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.status_label = ttk.Label(status_frame, text="系统已就绪，请导入学生名单")
        self.status_label.pack(side=tk.LEFT)
        
        self.selected_tree.bind("<Double-1>", self.on_selected_double_click)
        self.unselected_tree.bind("<Double-1>", self.on_unselected_double_click)
        
        self.init_groups()
        
        self.root.bind('<F1>', lambda e: self.start_lottery())
        self.root.bind('<F5>', lambda e: self.reset_system())
        self.root.bind('<Control-a>', lambda e: self.select_all_in_focus())
        self.root.bind('<Control-s>', lambda e: self.save_results())
        self.root.bind('<Control-o>', lambda e: self.import_students())
        
        self.update_history_combo()
    
    def timing_decorator(func):
        def wrapper(*args, **kwargs):
            start_time = time.time()
            result = func(*args, **kwargs)
            end_time = time.time()
            if end_time - start_time > 0.1:
                print(f"{func.__name__} 执行时间: {end_time - start_time:.3f}秒")
            return result
        return wrapper
    
    @lru_cache(maxsize=128)
    def get_student_weight(self, student_name):
        return self.student_weights.get(student_name, 5)
    
    def batch_update_treeview(self, tree, data, columns):
        tree.configure(yscrollcommand=None)
        
        tree.delete(*tree.get_children())
        
        for item_data in data:
            tree.insert("", tk.END, **item_data)
        
        scrollbar = tree.master.winfo_children()[-1]
        tree.configure(yscrollcommand=scrollbar.set)
    
    def select_all_in_focus(self):
        focus_widget = self.root.focus_get()
        if focus_widget == self.selected_tree:
            self.select_all_selected()
        elif focus_widget == self.unselected_tree:
            self.select_all_unselected()
    
    def select_all_selected(self):
        for item in self.selected_tree.get_children():
            self.selected_tree.change_state(item, "checked")
    
    def deselect_all_selected(self):
        for item in self.selected_tree.get_children():
            self.selected_tree.change_state(item, "unchecked")
    
    def select_all_unselected(self):
        for item in self.unselected_tree.get_children():
            self.unselected_tree.change_state(item, "checked")
    
    def deselect_all_unselected(self):
        for item in self.unselected_tree.get_children():
            self.unselected_tree.change_state(item, "unchecked")
    
    def batch_move_to_unselected(self):
        checked_items = self.selected_tree.get_checked_items()
        if not checked_items:
            messagebox.showinfo("提示", "请先勾选要移回的学生")
            return
        
        if self.auto_backup:
            self.auto_backup_data()
            
        moved_count = 0
        for student_name in checked_items:
            if student_name in self.selected_students:
                self.selected_students.remove(student_name)
                if student_name not in self.last_round_unselected:
                    self.last_round_unselected.append(student_name)
                moved_count += 1
        
        self.selected_tree.clear_all_checks()
        
        self.update_selected_tree()
        self.update_unselected_tree()
        self.update_statistics()
        self.save_unselected()
        self.status_label.config(text=f"已将 {moved_count} 名学生移回未抽中名单")
    
    def batch_move_to_selected(self):
        checked_items = self.unselected_tree.get_checked_items()
        if not checked_items:
            messagebox.showinfo("提示", "请先勾选要标记为已抽中的学生")
            return
        
        if self.auto_backup:
            self.auto_backup_data()
            
        moved_count = 0
        for student_name in checked_items:
            if student_name in self.last_round_unselected:
                self.last_round_unselected.remove(student_name)
                if student_name not in self.selected_students:
                    self.selected_students.append(student_name)
                moved_count += 1
        
        self.unselected_tree.clear_all_checks()
        
        self.update_selected_tree()
        self.update_unselected_tree()
        self.update_statistics()
        self.save_unselected()
        self.status_label.config(text=f"已将 {moved_count} 名学生标记为已抽中")
    
    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="导入名单", command=self.import_students, accelerator="Ctrl+O")
        file_menu.add_command(label="导出结果", command=self.save_results, accelerator="Ctrl+S")
        file_menu.add_command(label="备份管理", command=self.show_backup_manager)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit, accelerator="Ctrl+Q")
        
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="编辑", menu=edit_menu)
        edit_menu.add_command(label="添加学生", command=self.add_student_dialog)
        edit_menu.add_command(label="批量添加", command=self.batch_add_dialog)
        edit_menu.add_separator()
        edit_menu.add_command(label="设置权重", command=self.set_weights_dialog)
        edit_menu.add_separator()
        edit_menu.add_command(label="全选已抽中", command=self.select_all_selected, accelerator="Ctrl+Shift+A")
        edit_menu.add_command(label="全选未抽中", command=self.select_all_unselected, accelerator="Ctrl+Shift+U")
        
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="工具", menu=tools_menu)
        tools_menu.add_command(label="随机打乱", command=self.shuffle_students)
        tools_menu.add_command(label="智能平衡", command=self.enable_smart_balance)
        tools_menu.add_separator()
        tools_menu.add_command(label="查看历史", command=self.show_history)
        tools_menu.add_command(label="统计报告", command=self.generate_report)
        tools_menu.add_command(label="数据清理", command=self.data_cleanup)
        
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="设置", menu=settings_menu)
        settings_menu.add_command(label="系统设置", command=self.show_settings)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="使用说明", command=self.show_help)
        help_menu.add_command(label="快捷键", command=self.show_shortcuts)
        help_menu.add_command(label="关于", command=self.show_about)
    
    def init_groups(self):
        self.student_groups = {
            "默认分组": []
        }
    
    def update_history_combo(self):
        self.history_combo['values'] = [record['name'] for record in self.import_history]
    
    def on_history_selected(self, event):
        selected_name = self.history_var.get()
        for record in self.import_history:
            if record['name'] == selected_name:
                self.last_round_unselected = record['students'].copy()
                self.update_unselected_tree()
                self.update_statistics()
                self.update_students_text(self.last_round_unselected)
                
                self.status_label.config(text=f"已加载历史名单: {selected_name}")
                break
    
    def quick_draw_single(self):
        self.num_to_select.set("1")
        self.start_lottery()
    
    def quick_draw_multiple(self, count):
        self.num_to_select.set(str(count))
        self.start_lottery()
    
    def on_search(self, event):
        search_text = self.search_var.get().lower()
        if not search_text:
            self.update_unselected_tree()
            return
        
        filtered_students = [s for s in self.last_round_unselected if search_text in s.lower()]
        
        self.unselected_tree.delete(*self.unselected_tree.get_children())
        for student in filtered_students:
            self.unselected_tree.insert("", tk.END, text=student, 
                                      values=("未抽中", self.get_student_weight(student)))
    
    def clear_search(self):
        self.search_var.set("")
        self.update_unselected_tree()
    
    def on_selected_search(self, event):
        search_text = self.selected_search_var.get().lower()
        if not search_text:
            self.update_selected_tree()
            return
        
        filtered_students = [s for s in self.selected_students if search_text in s.lower()]
        
        self.selected_tree.delete(*self.selected_tree.get_children())
        for student in filtered_students:
            round_num = "未知"
            timestamp = "未知"
            for record in self.lottery_history:
                if student in record["selected"]:
                    round_num = str(record["round"])
                    timestamp = record["timestamp"]
                    break
            
            self.selected_tree.insert("", tk.END, text=student, 
                                   values=(round_num, timestamp, self.get_student_weight(student)))
    
    def clear_selected_search(self):
        self.selected_search_var.set("")
        self.update_selected_tree()
    
    def on_unselected_search(self, event):
        search_text = self.unselected_search_var.get().lower()
        if not search_text:
            self.update_unselected_tree()
            return
        
        filtered_students = [s for s in self.last_round_unselected if search_text in s.lower()]
        
        self.unselected_tree.delete(*self.unselected_tree.get_children())
        for student in filtered_students:
            self.unselected_tree.insert("", tk.END, text=student, 
                                      values=("未抽中", self.get_student_weight(student)))
    
    def clear_unselected_search(self):
        self.unselected_search_var.set("")
        self.update_unselected_tree()
    
    def quick_draw(self):
        if not self.last_round_unselected:
            messagebox.showinfo("提示", "没有学生可以抽取")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("一键抽取")
        dialog.geometry("300x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="抽取人数:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        num_var = tk.StringVar(value=str(min(5, len(self.last_round_unselected))))
        num_entry = ttk.Entry(dialog, textvariable=num_var, width=10)
        num_entry.grid(row=0, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))
        num_entry.focus()
        num_entry.select_range(0, tk.END)
        
        def do_quick_draw():
            try:
                num = int(num_var.get())
                if num <= 0:
                    messagebox.showwarning("警告", "抽取人数必须大于0")
                    return
                
                if num > len(self.last_round_unselected):
                    messagebox.showwarning("警告", f"抽取人数不能超过未抽中人数 {len(self.last_round_unselected)}")
                    return
                
                if self.auto_backup:
                    self.auto_backup_data()
                
                selected = random.sample(self.last_round_unselected, num)
                self.selected_students.extend(selected)
                
                self.last_round_unselected = [student for student in self.last_round_unselected if student not in selected]
                
                self.lottery_history.append({
                    "round": self.current_round,
                    "selected": selected,
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "mode": "一键抽取"
                })
                
                self.update_selected_tree()
                self.update_unselected_tree()
                
                self.current_round += 1
                self.round_label.config(text=str(self.current_round))
                
                self.update_statistics()
                
                self.show_enhanced_results(selected)
                
                self.status_label.config(text=f"一键抽取完成，已抽取 {len(selected)} 人")
                
                self.save_unselected()
                
                dialog.destroy()
            except ValueError:
                messagebox.showwarning("警告", "请输入有效的抽取人数")
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=10)
        
        ttk.Button(btn_frame, text="抽取", command=do_quick_draw).pack(side=tk.LEFT)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=(5, 0))
        
        dialog.bind('<Return>', lambda e: do_quick_draw())
    
    def update_students_text(self, students):
        self.students_text.config(state='normal')
        self.students_text.delete(1.0, tk.END)
        self.students_text.insert(1.0, "\n".join(students))
        self.students_text.config(state='disabled')
    
    def set_weights_dialog(self):
        if not self.last_round_unselected:
            messagebox.showinfo("提示", "没有学生可以设置权重")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("设置权重")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        help_frame = ttk.Frame(dialog)
        help_frame.pack(fill=tk.X, padx=10, pady=10)
        ttk.Label(help_frame, text="权重越高，被抽中的概率越大。权重范围: 1-10").pack(anchor=tk.W)
        
        tree_frame = ttk.Frame(dialog)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        weight_tree = ttk.Treeview(tree_frame, columns=("weight",), show="tree headings", height=15)
        weight_tree.heading("#0", text="学生姓名")
        weight_tree.heading("weight", text="权重")
        weight_tree.column("#0", width=200)
        weight_tree.column("weight", width=100)
        
        for student in self.last_round_unselected:
            weight = self.get_student_weight(student)
            weight_tree.insert("", tk.END, text=student, values=(weight,))
        
        def on_weight_edit(event):
            item = weight_tree.selection()[0]
            column = weight_tree.identify_column(event.x)
            
            if column == "#2":
                x, y, width, height = weight_tree.bbox(item, column)
                value = weight_tree.item(item, "values")[0]
                
                entry = ttk.Entry(tree_frame)
                entry.place(x=x, y=y, width=width, height=height)
                entry.insert(0, value)
                entry.select_range(0, tk.END)
                entry.focus()
                
                def save_edit():
                    try:
                        new_weight = int(entry.get())
                        if 1 <= new_weight <= 10:
                            weight_tree.set(item, column="weight", value=new_weight)
                            student_name = weight_tree.item(item, "text")
                            self.student_weights[student_name] = new_weight
                            entry.destroy()
                        else:
                            messagebox.showwarning("警告", "权重必须在1-10之间")
                    except ValueError:
                        messagebox.showwarning("警告", "请输入有效的数字")
                
                entry.bind("<Return>", lambda e: save_edit())
                entry.bind("<FocusOut>", lambda e: save_edit())
        
        weight_tree.bind("<Double-1>", on_weight_edit)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=weight_tree.yview)
        weight_tree.configure(yscrollcommand=scrollbar.set)
        
        weight_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(btn_frame, text="确定", command=dialog.destroy).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(btn_frame, text="应用默认权重", command=lambda: self.apply_default_weights(weight_tree)).pack(side=tk.RIGHT, padx=(5, 0))
    
    def apply_default_weights(self, weight_tree):
        for item in weight_tree.get_children():
            student_name = weight_tree.item(item, "text")
            self.student_weights[student_name] = 5
            weight_tree.set(item, column="weight", value=5)
    
    def reset_weights(self):
        if messagebox.askyesno("确认", "确定要重置所有权重吗？"):
            self.student_weights = {}
            self.update_unselected_tree()
            self.status_label.config(text="已重置所有权重")
    
    def enable_smart_balance(self):
        if not self.last_round_unselected:
            messagebox.showinfo("提示", "没有学生可以启用智能平衡")
            return
        
        selection_count = {}
        all_students = set(self.last_round_unselected + self.selected_students)
        
        for student in all_students:
            selection_count[student] = 0
        
        for record in self.lottery_history:
            for student in record["selected"]:
                if student in selection_count:
                    selection_count[student] += 1
        
        for student in self.last_round_unselected:
            count = selection_count.get(student, 0)
            new_weight = max(1, min(10, 10 - count))
            self.student_weights[student] = new_weight
        
        self.update_unselected_tree()
        self.status_label.config(text="已启用智能平衡，根据历史抽签记录调整了权重")
    
    def batch_add_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("批量添加学生")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="请输入学生姓名(每行一个):").pack(anchor=tk.W, padx=10, pady=10)
        
        text_area = scrolledtext.ScrolledText(dialog, height=10, wrap=tk.WORD)
        text_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        text_area.focus()
        
        def add_students():
            content = text_area.get(1.0, tk.END).strip()
            if content:
                students = [name.strip() for name in content.split('\n') if name.strip()]
                
                current_students = set(self.last_round_unselected + self.selected_students)
                new_students = []
                duplicates = 0
                
                for student in students:
                    if student not in current_students:
                        new_students.append(student)
                        current_students.add(student)
                    else:
                        duplicates += 1
                
                if self.auto_backup:
                    self.auto_backup_data()
                
                self.last_round_unselected.extend(new_students)
                
                self.update_unselected_tree()
                self.update_statistics()
                self.update_students_text(self.last_round_unselected)
                
                dialog.destroy()
                
                if duplicates > 0:
                    self.status_label.config(text=f"成功添加 {len(new_students)} 名学生，跳过 {duplicates} 个重复项")
                else:
                    self.status_label.config(text=f"成功添加 {len(new_students)} 名学生")
            else:
                messagebox.showwarning("警告", "请输入学生姓名")
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(btn_frame, text="添加", command=add_students).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT)
    
    def auto_backup_data(self):
        backup = {
            'selected': self.selected_students.copy(),
            'unselected': self.last_round_unselected.copy(),
            'round': self.current_round,
            'history': self.lottery_history.copy(),
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        try:
            with open("auto_backup.json", 'w', encoding='utf-8') as file:
                json.dump(backup, file, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"自动备份时出错: {str(e)}")
    
    def enhanced_auto_backup(self):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_data = {
            'version': self.data_version,
            'selected': self.selected_students.copy(),
            'unselected': self.last_round_unselected.copy(),
            'round': self.current_round,
            'history': self.lottery_history.copy(),
            'weights': self.student_weights.copy(),
            'import_history': self.import_history.copy(),
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'backup_id': f"backup_{timestamp}"
        }
        
        try:
            if not os.path.exists("backups"):
                os.makedirs("backups")
            
            backup_file = f"backups/backup_{timestamp}.json"
            with open(backup_file, 'w', encoding='utf-8') as file:
                json.dump(backup_data, file, ensure_ascii=False, indent=2)
            
            self.rotate_backups()
            
            self.backup_count += 1
        except Exception as e:
            print(f"增强备份时出错: {str(e)}")
    
    def rotate_backups(self):
        try:
            if not os.path.exists("backups"):
                return
            
            backup_files = []
            for file in os.listdir("backups"):
                if file.startswith("backup_") and file.endswith(".json"):
                    file_path = os.path.join("backups", file)
                    backup_files.append((file_path, os.path.getctime(file_path)))
            
            backup_files.sort(key=lambda x: x[1])
            
            while len(backup_files) > self.max_backups:
                oldest_file = backup_files.pop(0)[0]
                os.remove(oldest_file)
        except Exception as e:
            print(f"备份轮转时出错: {str(e)}")
    
    def show_backup_manager(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("备份管理")
        dialog.geometry("600x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        tree_frame = ttk.Frame(dialog)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        backup_tree = ttk.Treeview(tree_frame, columns=("timestamp", "students", "round"), show="headings")
        backup_tree.heading("timestamp", text="备份时间")
        backup_tree.heading("students", text="学生数量")
        backup_tree.heading("round", text="轮次")
        backup_tree.column("timestamp", width=150)
        backup_tree.column("students", width=100)
        backup_tree.column("round", width=80)
        
        backup_files = []
        if os.path.exists("backups"):
            for file in os.listdir("backups"):
                if file.startswith("backup_") and file.endswith(".json"):
                    file_path = os.path.join("backups", file)
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                            backup_files.append((file_path, data))
                    except:
                        continue
        
        backup_files.sort(key=lambda x: x[1].get('timestamp', ''), reverse=True)
        
        for file_path, data in backup_files:
            timestamp = data.get('timestamp', '未知')
            total_students = len(data.get('selected', [])) + len(data.get('unselected', []))
            round_num = data.get('round', 1)
            
            backup_tree.insert("", tk.END, values=(timestamp, total_students, round_num), tags=(file_path,))
        
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=backup_tree.yview)
        backup_tree.configure(yscrollcommand=scrollbar.set)
        
        backup_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        def restore_backup():
            selection = backup_tree.selection()
            if selection:
                file_path = backup_tree.item(selection[0], "tags")[0]
                if messagebox.askyesno("确认", "确定要恢复此备份吗？当前数据将会被覆盖"):
                    self.restore_from_backup(file_path)
                    dialog.destroy()
        
        def delete_backup():
            selection = backup_tree.selection()
            if selection:
                file_path = backup_tree.item(selection[0], "tags")[0]
                if messagebox.askyesno("确认", "确定要删除此备份吗？"):
                    try:
                        os.remove(file_path)
                        backup_tree.delete(selection[0])
                    except Exception as e:
                        messagebox.showerror("错误", f"删除备份时出错: {str(e)}")
        
        ttk.Button(btn_frame, text="恢复备份", command=restore_backup).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="删除备份", command=delete_backup).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="关闭", command=dialog.destroy).pack(side=tk.RIGHT)
    
    def restore_from_backup(self, backup_file):
        try:
            with open(backup_file, 'r', encoding='utf-8') as file:
                data = json.load(file)
                
            self.selected_students = data.get('selected', [])
            self.last_round_unselected = data.get('unselected', [])
            self.current_round = data.get('round', 1)
            self.lottery_history = data.get('history', [])
            self.student_weights = data.get('weights', {})
            self.import_history = data.get('import_history', [])
            
            self.update_students_text(self.last_round_unselected)
            self.update_selected_tree()
            self.update_unselected_tree()
            self.round_label.config(text=str(self.current_round))
            self.update_history_combo()
            self.update_statistics()
            
            self.status_label.config(text=f"已从备份恢复数据 (备份时间: {data.get('timestamp', '未知')})")
        except Exception as e:
            messagebox.showerror("错误", f"恢复备份时出错: {str(e)}")
    
    def show_settings(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("系统设置")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        auto_save_frame = ttk.LabelFrame(dialog, text="自动保存", padding="10")
        auto_save_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.auto_save = tk.BooleanVar(value=self.config.get('auto_save', True))
        ttk.Checkbutton(auto_save_frame, text="自动保存抽签进度", variable=self.auto_save).pack(anchor=tk.W)
        
        backup_frame = ttk.LabelFrame(dialog, text="备份设置", padding="10")
        backup_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.auto_backup_var = tk.BooleanVar(value=self.config.get('auto_backup', True))
        ttk.Checkbutton(backup_frame, text="重要操作前自动备份", variable=self.auto_backup_var).pack(anchor=tk.W)
        
        export_frame = ttk.LabelFrame(dialog, text="导出设置", padding="10")
        export_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.include_timestamp = tk.BooleanVar(value=self.config.get('include_timestamp', True))
        ttk.Checkbutton(export_frame, text="导出时包含时间戳", variable=self.include_timestamp).pack(anchor=tk.W)
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        def save_settings():
            self.config['auto_save'] = self.auto_save.get()
            self.config['auto_backup'] = self.auto_backup_var.get()
            self.config['include_timestamp'] = self.include_timestamp.get()
            self.auto_backup = self.auto_backup_var.get()
            self.save_config()
            dialog.destroy()
        
        ttk.Button(btn_frame, text="保存设置", command=save_settings).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.RIGHT, padx=(5, 0))
    
    def generate_report(self):
        if not self.selected_students and not self.last_round_unselected:
            messagebox.showinfo("提示", "没有数据可以生成报告")
            return
        
        total = len(self.selected_students) + len(self.last_round_unselected)
        selected_count = len(self.selected_students)
        unselected_count = len(self.last_round_unselected)
        
        selection_frequency = {}
        for record in self.lottery_history:
            for student in record["selected"]:
                selection_frequency[student] = selection_frequency.get(student, 0) + 1
        
        report = "抽签系统统计报告\n"
        report += "=" * 40 + "\n"
        report += f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        report += f"总抽签轮次: {self.current_round - 1}\n"
        report += f"总人数: {total}\n"
        report += f"已抽中人数: {selected_count}\n"
        report += f"未抽中人数: {unselected_count}\n\n"
        
        report += "抽签频率统计:\n"
        report += "-" * 20 + "\n"
        for student, count in sorted(selection_frequency.items(), key=lambda x: x[1], reverse=True):
            report += f"{student}: {count}次\n"
        
        report_window = tk.Toplevel(self.root)
        report_window.title("统计报告")
        report_window.geometry("500x400")
        
        text_area = scrolledtext.ScrolledText(report_window, wrap=tk.WORD, width=60, height=20)
        text_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        text_area.insert(1.0, report)
        text_area.config(state=tk.DISABLED)
        
        export_frame = ttk.Frame(report_window)
        export_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(export_frame, text="导出报告", command=lambda: self.export_report(report)).pack(side=tk.RIGHT)
    
    def export_report(self, report):
        file_path = filedialog.asksaveasfilename(
            title="导出报告",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(report)
                messagebox.showinfo("成功", f"报告已导出到: {file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"导出报告时出错: {str(e)}")
    
    def data_cleanup(self):
        if not self.selected_students and not self.last_round_unselected:
            messagebox.showinfo("提示", "没有数据可以清理")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("数据清理")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="选择清理选项:", font=("Arial", 10, "bold")).pack(anchor=tk.W, padx=10, pady=10)
        
        cleanup_var = tk.StringVar(value="none")
        
        ttk.Radiobutton(dialog, text="仅清理重复学生", variable=cleanup_var, value="duplicates").pack(anchor=tk.W, padx=20, pady=5)
        ttk.Radiobutton(dialog, text="清理空记录", variable=cleanup_var, value="empty").pack(anchor=tk.W, padx=20, pady=5)
        ttk.Radiobutton(dialog, text="重置所有数据", variable=cleanup_var, value="reset").pack(anchor=tk.W, padx=20, pady=5)
        
        stats_frame = ttk.Frame(dialog)
        stats_frame.pack(fill=tk.X, padx=10, pady=10)
        
        total_count = len(self.selected_students) + len(self.last_round_unselected)
        selected_count = len(self.selected_students)
        unselected_count = len(self.last_round_unselected)
        
        ttk.Label(stats_frame, text=f"总人数: {total_count}").grid(row=0, column=0, sticky=tk.W)
        ttk.Label(stats_frame, text=f"已抽中: {selected_count}").grid(row=0, column=1, sticky=tk.W, padx=(20, 0))
        ttk.Label(stats_frame, text=f"未抽中: {unselected_count}").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        ttk.Label(stats_frame, text=f"历史记录: {len(self.lottery_history)}").grid(row=1, column=1, sticky=tk.W, padx=(20, 0), pady=(5, 0))
        
        def perform_cleanup():
            option = cleanup_var.get()
            if option == "none":
                messagebox.showwarning("警告", "请选择清理选项")
                return
            
            if option == "duplicates":
                original_total = total_count
                self.selected_students = list(dict.fromkeys(self.selected_students))
                self.last_round_unselected = list(dict.fromkeys(self.last_round_unselected))
                
                new_total = len(self.selected_students) + len(self.last_round_unselected)
                removed = original_total - new_total
                
                self.update_selected_tree()
                self.update_unselected_tree()
                self.update_statistics()
                
                messagebox.showinfo("完成", f"已清理 {removed} 个重复学生")
                
            elif option == "empty":
                self.selected_students = [s for s in self.selected_students if s.strip()]
                self.last_round_unselected = [s for s in self.last_round_unselected if s.strip()]
                
                self.update_selected_tree()
                self.update_unselected_tree()
                self.update_statistics()
                
                messagebox.showinfo("完成", "已清理所有空记录")
                
            elif option == "reset":
                if messagebox.askyesno("确认", "确定要重置所有数据吗？此操作不可撤销"):
                    self.selected_students = []
                    self.last_round_unselected = []
                    self.lottery_history = []
                    self.current_round = 1
                    self.student_weights = {}
                    
                    self.update_selected_tree()
                    self.update_unselected_tree()
                    self.update_statistics()
                    self.round_label.config(text="1")
                    
                    messagebox.showinfo("完成", "已重置所有数据")
            
            dialog.destroy()
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(btn_frame, text="执行清理", command=perform_cleanup).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.RIGHT)
    
    def show_help(self):
        help_text = """
增强版终极抽签系统使用说明

1. 名单管理
   - 导入名单: 从文本文件导入学生名单
   - 历史名单: 选择之前导入过的名单
   - 手动添加: 逐个添加学生
   - 批量添加: 一次性添加多个学生
   - 随机打乱: 随机打乱未抽中学生名单

2. 权重设置
   - 设置权重: 为每个学生设置抽中概率(1-10)
   - 智能平衡: 根据历史记录自动调整权重
   - 重置权重: 恢复默认权重

3. 抽签操作
   - 开始抽签: 开始新一轮抽签
   - 跳过本轮: 跳过当前轮次
   - 一键抽取: 快速抽取指定数量的学生
   - 重置系统: 重置所有抽签记录

4. 批量操作
   - 勾选操作: 点击姓名前的复选框选择学生
   - 全选/取消选择: 快速选择或取消选择所有学生
   - 批量移动: 将勾选的学生批量移动到另一侧

5. 数据管理
   - 保存结果: 保存抽签结果到文件
   - 导出名单: 导出未抽中名单
   - 查看历史: 查看抽签历史记录
   - 统计报告: 生成统计报告
   - 数据清理: 清理重复或空数据

6. 其他功能
   - 系统设置: 配置系统参数
   - 独立动画: 显示独立的抽签动画窗口
   - 搜索功能: 在已抽中和未抽中名单中搜索学生

快捷键:
   - F1: 开始抽签
   - F5: 重置系统
   - Ctrl+A: 全选当前列表
   - Ctrl+S: 保存结果
   - Ctrl+O: 导入名单
"""
        messagebox.showinfo("使用说明", help_text)
    
    def show_shortcuts(self):
        shortcuts = """
快捷键列表:

F1          - 开始抽签
F5          - 重置系统
Ctrl+A      - 全选当前列表
Ctrl+S      - 保存结果
Ctrl+O      - 导入名单
Ctrl+Shift+A - 全选已抽中学生
Ctrl+Shift+U - 全选未抽中学生
"""
        messagebox.showinfo("快捷键", shortcuts)
    
    def show_about(self):
        about_text = """
清辞版抽签系统 v2.0

功能特点:
- 支持权重抽签和智能平衡
- 支持历史名单记录
- 支持批量操作
- 支持一键抽取
- 独立的抽签动画窗口
- 自动备份功能
- 增强的交互界面，支持复选框选择
- 搜索功能，支持已抽中和未抽中名单搜索
- 数据清理功能
- 自动保存功能

开发者: 被榨干的王振宇
版本: 2.0
更新日期: 2025
"""
        messagebox.showinfo("关于", about_text)
    
    def on_selected_double_click(self, event):
        item = self.selected_tree.identify_row(event.y)
        if item:
            student_name = self.selected_tree.item(item, "text")
            if messagebox.askyesno("确认", f"确定要将 {student_name} 移回未抽中名单吗？"):
                if self.auto_backup:
                    self.auto_backup_data()
                
                self.selected_students = [s for s in self.selected_students if s != student_name]
                if student_name not in self.last_round_unselected:
                    self.last_round_unselected.append(student_name)
                
                self.update_selected_tree()
                self.update_unselected_tree()
                self.update_statistics()
                self.save_unselected()
                self.status_label.config(text=f"已将 {student_name} 移回未抽中名单")
    
    def on_unselected_double_click(self, event):
        item = self.unselected_tree.identify_row(event.y)
        if item:
            student_name = self.unselected_tree.item(item, "text")
            if messagebox.askyesno("确认", f"确定要手动将 {student_name} 标记为已抽中吗？"):
                if self.auto_backup:
                    self.auto_backup_data()
                
                self.last_round_unselected = [s for s in self.last_round_unselected if s != student_name]
                if student_name not in self.selected_students:
                    self.selected_students.append(student_name)
                
                self.update_selected_tree()
                self.update_unselected_tree()
                self.update_statistics()
                self.save_unselected()
                self.status_label.config(text=f"已将 {student_name} 标记为已抽中")
    
    def add_student_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("添加学生")
        dialog.geometry("300x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="学生姓名:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        name_var = tk.StringVar()
        name_entry = ttk.Entry(dialog, textvariable=name_var, width=20)
        name_entry.grid(row=0, column=1, padx=10, pady=10, sticky=(tk.W, tk.E))
        name_entry.focus()
        
        def add_student():
            name = name_var.get().strip()
            if name:
                if name in self.last_round_unselected or name in self.selected_students:
                    messagebox.showwarning("警告", f"学生 '{name}' 已存在")
                    return
                
                if self.auto_backup:
                    self.auto_backup_data()
                
                if name not in self.last_round_unselected and name not in self.selected_students:
                    self.last_round_unselected.append(name)
                
                self.update_unselected_tree()
                self.update_statistics()
                self.update_students_text(self.last_round_unselected)
                dialog.destroy()
                self.status_label.config(text=f"已添加学生: {name}")
            else:
                messagebox.showwarning("警告", "请输入学生姓名")
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=10)
        
        ttk.Button(btn_frame, text="添加", command=add_student).grid(row=0, column=0, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).grid(row=0, column=1, padx=5)
        
        dialog.bind('<Return>', lambda e: add_student())
    
    def shuffle_students(self):
        if not self.last_round_unselected:
            messagebox.showinfo("提示", "没有需要打乱的学生名单")
            return
        
        if self.auto_backup:
            self.auto_backup_data()
        
        random.shuffle(self.last_round_unselected)
        self.update_unselected_tree()
        self.update_students_text(self.last_round_unselected)
        self.status_label.config(text="学生名单已随机打乱")
    
    def update_statistics(self):
        total = len(self.selected_students) + len(self.last_round_unselected)
        selected_count = len(self.selected_students)
        unselected_count = len(self.last_round_unselected)
        
        self.stats_label.config(
            text=f"总人数: {total} | 已抽中: {selected_count} | 未抽中: {unselected_count}"
        )
        
        self.selected_count_label.config(text=f"已抽中: {selected_count}人")
        self.unselected_count_label.config(text=f"未抽中: {unselected_count}人")
    
    def import_students(self):
        file_path = filedialog.askopenfilename(
            title="选择学生名单文件",
            filetypes=[("文本文件", "*.txt"), ("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as file:
                    students = [line.strip() for line in file if line.strip()]
                
                if students:
                    if self.auto_backup:
                        self.auto_backup_data()
                    
                    current_students = set(self.last_round_unselected + self.selected_students)
                    new_students = []
                    duplicates = 0
                    
                    for student in students:
                        if student not in current_students:
                            new_students.append(student)
                            current_students.add(student)
                        else:
                            duplicates += 1
                    
                    self.last_round_unselected.extend(new_students)
                    self.update_unselected_tree()
                    self.update_statistics()
                    self.update_students_text(self.last_round_unselected)
                    
                    file_name = os.path.basename(file_path)
                    self.import_history.append({
                        'name': file_name,
                        'path': file_path,
                        'students': new_students.copy(),
                        'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    })
                    
                    self.update_history_combo()
                    
                    if duplicates > 0:
                        self.status_label.config(text=f"成功导入 {len(new_students)} 名学生，跳过 {duplicates} 个重复项")
                    else:
                        self.status_label.config(text=f"成功导入 {len(new_students)} 名学生")
                else:
                    messagebox.showwarning("警告", "文件为空或格式不正确")
            except Exception as e:
                messagebox.showerror("错误", f"读取文件时出错: {str(e)}")
    
    def clear_students(self):
        if messagebox.askyesno("确认", "确定要清空所有学生名单吗？"):
            if self.auto_backup:
                self.auto_backup_data()
            
            self.all_students = []
            self.selected_students = []
            self.current_round = 1
            self.last_round_unselected = []
            self.lottery_history = []
            self.update_selected_tree()
            self.update_unselected_tree()
            self.update_statistics()
            self.round_label.config(text="1")
            self.update_students_text([])
            self.status_label.config(text="已清空所有名单")
    
    def start_lottery(self):
        if self.is_animating:
            return
            
        if not self.last_round_unselected:
            messagebox.showwarning("警告", "没有学生可以抽取")
            return
        
        try:
            num_to_select = int(self.num_to_select.get())
            if num_to_select <= 0:
                messagebox.showwarning("警告", "抽取人数必须大于0")
                return
        except ValueError:
            messagebox.showwarning("警告", "请输入有效的抽取人数")
            return
        
        if num_to_select > len(self.last_round_unselected):
            messagebox.showwarning("警告", f"抽取人数不能超过未抽中人数 {len(self.last_round_unselected)}")
            return
        
        if self.auto_backup:
            self.auto_backup_data()
        
        self.progress.pack(fill=tk.X, pady=(5, 0))
        self.progress.start()
        self.is_animating = True
        
        threading.Thread(target=self.animate_lottery, args=(num_to_select,), daemon=True).start()
    
    def animate_lottery(self, num_to_select):
        speed_map = {"慢速": 0.2, "中速": 0.1, "快速": 0.05}
        delay = speed_map.get(self.animation_speed.get(), 0.1)
        iterations = 20
        
        temp_students = self.last_round_unselected.copy()
        
        if self.show_animation.get():
            self.root.after(0, self.create_enhanced_animation)
        
        for i in range(iterations):
            if not self.is_animating:
                return
                
            random.shuffle(temp_students)
            display_text = " >>> ".join(temp_students[:5])
            self.root.after(0, lambda: self.status_label.config(text=f"抽签中... {display_text}"))
            
            if self.show_animation.get() and self.animation_window:
                self.root.after(0, lambda: self.enhanced_update_animation(temp_students[:3]))
            
            time.sleep(delay)
        
        if self.lottery_mode.get() == "权重模式":
            selected = self.weighted_lottery(num_to_select)
        elif self.lottery_mode.get() == "公平模式":
            selected = self.fair_lottery(num_to_select)
        else:
            selected = random.sample(self.last_round_unselected, num_to_select)
        
        self.selected_students.extend(selected)
        
        self.last_round_unselected = [student for student in self.last_round_unselected if student not in selected]
        
        self.lottery_history.append({
            "round": self.current_round,
            "selected": selected,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "mode": self.lottery_mode.get()
        })
        
        self.root.after(0, self.finish_lottery, selected)
    
    def create_enhanced_animation(self):
        if not self.show_animation.get():
            return
        
        if self.animation_window and self.animation_window.winfo_exists():
            self.animation_window.destroy()
        
        self.animation_window = tk.Toplevel(self.root)
        self.animation_window.title("🎯 抽签动画")
        self.animation_window.geometry("500x400")
        self.animation_window.transient(self.root)
        self.animation_window.resizable(False, False)
        self.animation_window.configure(bg='#2c3e50')
        
        self.animation_window.geometry("+%d+%d" % (
            self.root.winfo_x() + (self.root.winfo_width() - 500) // 2,
            self.root.winfo_y() + (self.root.winfo_height() - 400) // 2
        ))
        
        title_label = tk.Label(self.animation_window, text="抽签进行中...", 
                              font=("Arial", 18, "bold"), fg="white", bg='#2c3e50')
        title_label.pack(pady=20)
        
        self.animation_display = tk.Label(self.animation_window, text="", 
                                         font=("Arial", 16), fg="#ecf0f1", bg='#34495e',
                                         width=30, height=8, relief=tk.SUNKEN)
        self.animation_display.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)
        
        self.animation_progress = ttk.Progressbar(self.animation_window, mode='indeterminate')
        self.animation_progress.pack(fill=tk.X, padx=20, pady=20)
        self.animation_progress.start()
        
        close_btn = ttk.Button(self.animation_window, text="停止动画", 
                              command=self.cancel_animation)
        close_btn.pack(pady=10)
    
    def cancel_animation(self):
        self.is_animating = False
        self.close_animation()
        self.status_label.config(text="抽签动画已取消")
    
    def enhanced_update_animation(self, students):
        if self.animation_window and self.animation_window.winfo_exists():
            display_lines = []
            for i, student in enumerate(students[:5]):
                emoji = "🎲" if i == 0 else "✨"
                display_lines.append(f"{emoji} {student}")
            
            display_text = "\n\n".join(display_lines)
            self.animation_display.config(text=display_text)
            
            if hasattr(self, 'flash_counter'):
                self.flash_counter += 1
                if self.flash_counter % 2 == 0:
                    self.animation_display.config(bg='#2c3e50')
                else:
                    self.animation_display.config(bg='#34495e')
            else:
                self.flash_counter = 0
    
    def close_animation(self):
        if self.animation_window and self.animation_window.winfo_exists():
            self.animation_window.destroy()
            self.animation_window = None
    
    def weighted_lottery(self, num_to_select):
        weights = []
        for student in self.last_round_unselected:
            weight = self.get_student_weight(student)
            weights.append(weight)
        
        selected = []
        available_students = self.last_round_unselected.copy()
        available_weights = weights.copy()
        
        for _ in range(num_to_select):
            if not available_students:
                break
            
            chosen = random.choices(available_students, weights=available_weights, k=1)[0]
            selected.append(chosen)
            
            index = available_students.index(chosen)
            available_students.pop(index)
            available_weights.pop(index)
        
        return selected
    
    def fair_lottery(self, num_to_select):
        selection_count = {}
        for student in self.last_round_unselected:
            selection_count[student] = 0
        
        for record in self.lottery_history:
            for student in record["selected"]:
                if student in selection_count:
                    selection_count[student] += 1
        
        max_count = max(selection_count.values()) if selection_count else 1
        weights = []
        for student in self.last_round_unselected:
            count = selection_count.get(student, 0)
            weight = max_count - count + 1
            weights.append(weight)
        
        selected = []
        available_students = self.last_round_unselected.copy()
        available_weights = weights.copy()
        
        for _ in range(num_to_select):
            if not available_students:
                break
            
            chosen = random.choices(available_students, weights=available_weights, k=1)[0]
            selected.append(chosen)
            
            index = available_students.index(chosen)
            available_students.pop(index)
            available_weights.pop(index)
        
        return selected
    
    def finish_lottery(self, selected):
        self.progress.stop()
        self.progress.pack_forget()
        self.is_animating = False
        
        self.close_animation()
        
        self.update_selected_tree()
        self.update_unselected_tree()
        
        self.current_round += 1
        self.round_label.config(text=str(self.current_round))
        
        self.update_statistics()
        
        self.show_enhanced_results(selected)
        
        self.status_label.config(text=f"第 {self.current_round-1} 轮抽签完成，已抽取 {len(selected)} 人")
        
        self.save_unselected()
    
    def show_enhanced_results(self, selected):
        result_window = tk.Toplevel(self.root)
        result_window.title("🎉 抽签结果")
        result_window.geometry("400x900")
        result_window.transient(self.root)
        result_window.configure(bg='#f8f9fa')
        
        title_label = tk.Label(result_window, 
                              text=f"第 {self.current_round-1} 轮抽签结果",
                              font=("Arial", 16, "bold"), 
                              bg='#f8f9fa', fg='#2c3e50')
        title_label.pack(pady=15)
        
        mode_label = tk.Label(result_window,
                             text=f"模式: {self.lottery_mode.get()}",
                             font=("Arial", 10),
                             bg='#f8f9fa', fg='#7f8c8d')
        mode_label.pack(pady=5)
        
        result_frame = tk.Frame(result_window, bg='#f8f9fa')
        result_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        for i, student in enumerate(selected, 1):
            student_frame = tk.Frame(result_frame, bg='#ffffff', relief=tk.RAISED, bd=1)
            student_frame.pack(fill=tk.X, pady=2)
            
            number_label = tk.Label(student_frame, text=f"{i}", 
                                   font=("Arial", 12, "bold"), 
                                   bg='#3498db', fg='white', width=3)
            number_label.pack(side=tk.LEFT, padx=5, pady=5)
            
            name_label = tk.Label(student_frame, text=student, 
                                 font=("Arial", 12),
                                 bg='#ffffff', anchor=tk.W)
            name_label.pack(side=tk.LEFT, padx=10, pady=5, fill=tk.X, expand=True)
        
        btn_frame = tk.Frame(result_window, bg='#f8f9fa')
        btn_frame.pack(fill=tk.X, padx=20, pady=15)
        
        ttk.Button(btn_frame, text="确定", 
                  command=result_window.destroy).pack(side=tk.RIGHT)
        
        result_window.after(10000, result_window.destroy)
    
    def reset_system(self):
        if messagebox.askyesno("确认", "确定要重置系统吗？这将清除所有抽签记录"):
            if self.auto_backup:
                self.auto_backup_data()
            
            self.selected_students = []
            self.current_round = 1
            self.last_round_unselected = []
            self.lottery_history = []
            self.update_selected_tree()
            self.update_unselected_tree()
            self.update_statistics()
            self.round_label.config(text="1")
            self.update_students_text([])
            self.status_label.config(text="系统已重置")
    
    def save_results(self):
        if not self.selected_students:
            messagebox.showwarning("警告", "没有抽签结果可保存")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="保存抽签结果",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write("抽签结果报告\n")
                    file.write("=" * 40 + "\n")
                    file.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    file.write(f"总抽签轮次: {self.current_round - 1}\n")
                    file.write(f"已抽中总人数: {len(self.selected_students)}\n")
                    file.write(f"剩余未抽中人数: {len(self.last_round_unselected)}\n\n")
                    
                    file.write("已抽中学生名单:\n")
                    file.write("-" * 20 + "\n")
                    for i, student in enumerate(self.selected_students, 1):
                        file.write(f"{i}. {student}\n")
                    
                    file.write("\n抽签历史:\n")
                    file.write("-" * 20 + "\n")
                    for record in self.lottery_history:
                        file.write(f"第{record['round']}轮 ({record['timestamp']}) [{record.get('mode', '常规')}]: {', '.join(record['selected'])}\n")
                
                messagebox.showinfo("成功", f"抽签结果已保存到: {file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"保存文件时出错: {str(e)}")
    
    def export_list(self):
        if not self.last_round_unselected:
            messagebox.showwarning("警告", "没有未抽中名单可导出")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="导出未抽中名单",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write("未抽中学生名单\n")
                    file.write("=" * 20 + "\n")
                    file.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    file.write(f"总人数: {len(self.last_round_unselected)}\n\n")
                    for student in self.last_round_unselected:
                        file.write(f"{student}\n")
                
                messagebox.showinfo("成功", f"未抽中名单已导出到: {file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"导出文件时出错: {str(e)}")
    
    def show_history(self):
        if not self.lottery_history:
            messagebox.showinfo("抽签历史", "暂无抽签历史记录")
            return
        
        history_text = "抽签历史记录:\n\n"
        for record in self.lottery_history:
            history_text += f"第{record['round']}轮 ({record['timestamp']}) [{record.get('mode', '常规')}]:\n"
            history_text += f"  抽中: {', '.join(record['selected'])}\n\n"
        
        history_window = tk.Toplevel(self.root)
        history_window.title("抽签历史记录")
        history_window.geometry("500x400")
        
        text_area = scrolledtext.ScrolledText(history_window, wrap=tk.WORD, width=60, height=20)
        text_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        text_area.insert(1.0, history_text)
        text_area.config(state=tk.DISABLED)
    
    def update_selected_tree(self):
        for item in self.selected_tree.get_children():
            self.selected_tree.delete(item)
        
        for student in self.selected_students:
            round_num = "未知"
            timestamp = "未知"
            for record in self.lottery_history:
                if student in record["selected"]:
                    round_num = str(record["round"])
                    timestamp = record["timestamp"]
                    break
            
            self.selected_tree.insert("", tk.END, text=student, values=(round_num, timestamp, self.get_student_weight(student)))
    
    def update_unselected_tree(self):
        for item in self.unselected_tree.get_children():
            self.unselected_tree.delete(item)
        
        for student in self.last_round_unselected:
            self.unselected_tree.insert("", tk.END, text=student, values=("未抽中", self.get_student_weight(student)))
    
    def save_unselected(self):
        try:
            with open("unselected_students.json", 'w', encoding='utf-8') as file:
                json.dump({
                    'unselected': self.last_round_unselected,
                    'selected': self.selected_students,
                    'round': self.current_round,
                    'history': self.lottery_history,
                    'weights': self.student_weights,
                    'import_history': self.import_history,
                    'settings': {
                        'auto_backup': self.auto_backup,
                        'auto_save': self.auto_save
                    },
                    'last_updated': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }, file, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存未抽中名单时出错: {str(e)}")
    
    def load_unselected(self):
        try:
            if os.path.exists("unselected_students.json"):
                with open("unselected_students.json", 'r', encoding='utf-8') as file:
                    data = json.load(file)
                    self.last_round_unselected = data.get('unselected', [])
                    self.selected_students = data.get('selected', [])
                    self.current_round = data.get('round', 1)
                    self.lottery_history = data.get('history', [])
                    self.student_weights = data.get('weights', {})
                    self.import_history = data.get('import_history', [])
                    
                    settings = data.get('settings', {})
                    self.auto_backup = settings.get('auto_backup', True)
                    self.auto_save = settings.get('auto_save', True)
                    
                    self.update_students_text(self.last_round_unselected)
                    self.update_selected_tree()
                    self.update_unselected_tree()
                    self.round_label.config(text=str(self.current_round))
                    self.update_history_combo()
                    
                    last_updated = data.get('last_updated', '未知')
                    self.status_label.config(text=f"已加载上次的抽签记录(最后更新: {last_updated})")
        except Exception as e:
            print(f"加载未抽中名单时出错: {str(e)}")
    
    def skip_round(self):
        if not self.last_round_unselected:
            messagebox.showinfo("提示", "没有未抽中学生，无法跳过")
            return
            
        if messagebox.askyesno("确认", "确定要跳过本轮抽签吗？"):
            self.current_round += 1
            self.round_label.config(text=str(self.current_round))
            self.status_label.config(text=f"已跳过第 {self.current_round-1} 轮抽签")


def main():
    root = tk.Tk()
    app = EnhancedLotterySystem(root)
    root.mainloop()

if __name__ == "__main__":
    main()