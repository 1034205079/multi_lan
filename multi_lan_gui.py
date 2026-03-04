"""
多语言对比工具 - GUI界面版本
使用multi_lan_core和apk_decompiler，避免代码重复
"""
import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import queue
import subprocess
from multi_lan_core import MultiLanguageCore
from apk_decompiler import APKDecompiler

# 尝试导入拖放支持
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False


class MultiLanguageGUI:
    """多语言对比工具 - GUI版本"""

    def __init__(self, root):
        self.root = root
        self.root.title("多语言对比工具")
        self.root.geometry("1100x750")
        self.root.minsize(1000, 700)

        # 配色方案
        self.colors = {
            'bg': '#F5F5F5',
            'card_bg': '#FFFFFF',
            'primary': '#2196F3',
            'success': '#4CAF50',
            'warning': '#FF9800',
            'error': '#F44336',
            'text': '#212121',
            'text_secondary': '#757575',
            'border': '#E0E0E0'
        }

        self.root.configure(bg=self.colors['bg'])

        # 核心业务逻辑和工具
        self.core = MultiLanguageCore()
        self.decompiler = None

        # UI状态变量
        self.package_path = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.selected_sheet = tk.StringVar()
        self.res_base_path = tk.StringVar(value="res")
        self.is_processing = False
        self.log_queue = queue.Queue()

        # 构建UI
        self._build_ui()

        # 启用拖放功能
        self._setup_drag_drop()
        self._update_log()

    def _build_ui(self):
        """构建用户界面"""
        main_frame = tk.Frame(self.root, bg=self.colors['bg'])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # 标题
        self._build_header(main_frame)

        # 左右分栏
        content_frame = tk.Frame(main_frame, bg=self.colors['bg'])
        content_frame.pack(fill=tk.BOTH, expand=True)

        # 左侧操作面板
        left_frame = tk.Frame(content_frame, bg=self.colors['bg'])
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        self._build_step1_package(left_frame)
        self._build_step2_excel(left_frame)
        self._build_step3_actions(left_frame)
        self._build_results_panel(left_frame)

        # 右侧日志面板
        right_frame = tk.Frame(content_frame, bg=self.colors['bg'])
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self._build_log_panel(right_frame)

    def _setup_drag_drop(self):
        """设置拖放功能"""
        if not HAS_DND:
            self.log("⚠ 未安装 tkinterdnd2，拖放功能不可用", 'warning')
            self.log("  安装方法: pip install tkinterdnd2", 'info')
            return

        # 为 APK 输入框和整个卡片添加拖放支持
        self.package_entry.drop_target_register(DND_FILES)
        self.package_entry.dnd_bind('<<Drop>>', self._on_drop_package)

        self.package_card.drop_target_register(DND_FILES)
        self.package_card.dnd_bind('<<Drop>>', self._on_drop_package)

        # 为 Excel 输入框和整个卡片添加拖放支持
        self.excel_entry.drop_target_register(DND_FILES)
        self.excel_entry.dnd_bind('<<Drop>>', self._on_drop_excel)

        self.excel_card.drop_target_register(DND_FILES)
        self.excel_card.dnd_bind('<<Drop>>', self._on_drop_excel)



    def _on_drop_package(self, event):
        """处理 APK 文件拖放"""
        file_path = event.data
        # 去除可能的大括号
        if file_path.startswith('{') and file_path.endswith('}'):
            file_path = file_path[1:-1]

        if file_path.lower().endswith('.apk'):
            self.package_path.set(file_path)
            self.log(f"✓ 已选择包文件: {os.path.basename(file_path)}", 'success')
        else:
            messagebox.showwarning("文件类型错误", "请拖放 APK 文件！")

    def _on_drop_excel(self, event):
        """处理 Excel 文件拖放"""
        file_path = event.data
        # 去除可能的大括号
        if file_path.startswith('{') and file_path.endswith('}'):
            file_path = file_path[1:-1]

        if file_path.lower().endswith(('.xlsx', '.xls')):
            self.excel_path.set(file_path)
            self.log(f"✓ 已选择Excel: {os.path.basename(file_path)}", 'success')
            self.load_excel_sheets(file_path)
        else:
            messagebox.showwarning("文件类型错误", "请拖放 Excel 文件（.xlsx 或 .xls）！")

    def _build_header(self, parent):
        """标题区域"""
        header = tk.Frame(parent, bg=self.colors['bg'])
        header.pack(fill=tk.X, pady=(0, 20))

        tk.Label(
            header,
            text="🌍 多语言对比工具",
            font=('Microsoft YaHei UI', 24, 'bold'),
            fg=self.colors['text'],
            bg=self.colors['bg']
        ).pack(side=tk.LEFT)

        tk.Label(
            header,
            text="APK资源 ↔ Excel对比分析",
            font=('Microsoft YaHei UI', 10),
            fg=self.colors['text_secondary'],
            bg=self.colors['bg']
        ).pack(side=tk.LEFT, padx=(10, 0), pady=(8, 0))

    def _build_step1_package(self, parent):
        """步骤1：APK反编译"""
        self.package_card = self._create_card(parent)
        card = self.package_card

        tk.Label(
            card,
            text="步骤 1：APK反编译（可选）",
            font=('Microsoft YaHei UI', 11, 'bold'),
            fg=self.colors['text'],
            bg=self.colors['card_bg']
        ).pack(anchor=tk.W, padx=20, pady=(15, 10))

        # 文件路径
        path_frame = tk.Frame(card, bg=self.colors['card_bg'])
        path_frame.pack(fill=tk.X, padx=20, pady=(0, 10))

        self.package_entry = tk.Entry(
            path_frame,
            textvariable=self.package_path,
            font=('Microsoft YaHei UI', 9),
            bg='#FAFAFA',
            relief=tk.FLAT,
            bd=0
        )
        self.package_entry.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, ipady=6, padx=(0, 5))

        tk.Button(
            path_frame,
            text="选择文件",
            command=self.browse_package,
            font=('Microsoft YaHei UI', 9),
            fg=self.colors['primary'],
            bg=self.colors['card_bg'],
            relief=tk.FLAT,
            cursor='hand2',
            padx=15,
            pady=6,
            bd=1,
            highlightthickness=1,
            highlightbackground=self.colors['primary']
        ).pack(side=tk.LEFT, padx=(0, 5))

        self.decompile_btn = tk.Button(
            path_frame,
            text="反编译",
            command=self.start_decompile,
            font=('Microsoft YaHei UI', 9, 'bold'),
            fg='white',
            bg=self.colors['primary'],
            activebackground='#1976D2',
            relief=tk.FLAT,
            cursor='hand2',
            padx=15,
            pady=6
        )
        self.decompile_btn.pack(side=tk.LEFT)

        # 提示
        tk.Label(
            card,
            text="💡 支持 APK 格式文件（可拖到此卡片任意位置）",
            font=('Microsoft YaHei UI', 9),
            fg=self.colors['text_secondary'],
            bg=self.colors['card_bg']
        ).pack(anchor=tk.W, padx=20, pady=(0, 5))

        # Res路径
        res_frame = tk.Frame(card, bg=self.colors['card_bg'])
        res_frame.pack(fill=tk.X, padx=20, pady=(0, 15))

        tk.Label(
            res_frame,
            text="Res目录：",
            font=('Microsoft YaHei UI', 9),
            fg=self.colors['text_secondary'],
            bg=self.colors['card_bg']
        ).pack(side=tk.LEFT)

        tk.Entry(
            res_frame,
            textvariable=self.res_base_path,
            font=('Microsoft YaHei UI', 9),
            bg='#FAFAFA',
            relief=tk.FLAT,
            bd=0,
            width=30
        ).pack(side=tk.LEFT, ipady=4, padx=(5, 0))

    def _build_step2_excel(self, parent):
        """步骤2：Excel选择"""
        self.excel_card = self._create_card(parent)
        card = self.excel_card

        tk.Label(
            card,
            text="步骤 2：选择Excel文件和Sheet",
            font=('Microsoft YaHei UI', 11, 'bold'),
            fg=self.colors['text'],
            bg=self.colors['card_bg']
        ).pack(anchor=tk.W, padx=20, pady=(15, 10))

        # Excel路径
        excel_frame = tk.Frame(card, bg=self.colors['card_bg'])
        excel_frame.pack(fill=tk.X, padx=20, pady=(0, 10))

        self.excel_entry = tk.Entry(
            excel_frame,
            textvariable=self.excel_path,
            font=('Microsoft YaHei UI', 9),
            bg='#FAFAFA',
            relief=tk.FLAT,
            bd=0
        )
        self.excel_entry.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, ipady=6, padx=(0, 5))

        tk.Button(
            excel_frame,
            text="选择Excel",
            command=self.browse_excel,
            font=('Microsoft YaHei UI', 9),
            fg=self.colors['primary'],
            bg=self.colors['card_bg'],
            relief=tk.FLAT,
            cursor='hand2',
            padx=15,
            pady=6,
            bd=1,
            highlightthickness=1,
            highlightbackground=self.colors['primary']
        ).pack(side=tk.LEFT)

        # 提示
        tk.Label(
            card,
            text="💡 支持 .xlsx 和 .xls 格式（可拖到此卡片任意位置）",
            font=('Microsoft YaHei UI', 9),
            fg=self.colors['text_secondary'],
            bg=self.colors['card_bg']
        ).pack(anchor=tk.W, padx=20, pady=(5, 0))

        # Sheet选择
        sheet_frame = tk.Frame(card, bg=self.colors['card_bg'])
        sheet_frame.pack(fill=tk.X, padx=20, pady=(0, 15))

        tk.Label(
            sheet_frame,
            text="Sheet：",
            font=('Microsoft YaHei UI', 9),
            fg=self.colors['text_secondary'],
            bg=self.colors['card_bg']
        ).pack(side=tk.LEFT)

        self.sheet_combo = ttk.Combobox(
            sheet_frame,
            textvariable=self.selected_sheet,
            state='readonly',
            font=('Microsoft YaHei UI', 9)
        )
        self.sheet_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

    def _build_step3_actions(self, parent):
        """步骤3：操作按钮"""
        card = self._create_card(parent)

        tk.Label(
            card,
            text="步骤 3：执行对比",
            font=('Microsoft YaHei UI', 11, 'bold'),
            fg=self.colors['text'],
            bg=self.colors['card_bg']
        ).pack(anchor=tk.W, padx=20, pady=(15, 10))

        btn_frame = tk.Frame(card, bg=self.colors['card_bg'])
        btn_frame.pack(fill=tk.X, padx=20, pady=(0, 15))

        self.compare_btn = tk.Button(
            btn_frame,
            text="🔍 开始对比",
            command=self.start_compare,
            font=('Microsoft YaHei UI', 10, 'bold'),
            fg='white',
            bg=self.colors['success'],
            activebackground='#388E3C',
            relief=tk.FLAT,
            cursor='hand2',
            padx=30,
            pady=10
        )
        self.compare_btn.pack(side=tk.LEFT, padx=(0, 10))

        tk.Button(
            btn_frame,
            text="📂 打开结果",
            command=self.open_results,
            font=('Microsoft YaHei UI', 9),
            fg=self.colors['primary'],
            bg=self.colors['card_bg'],
            relief=tk.FLAT,
            cursor='hand2',
            padx=20,
            pady=8,
            bd=1,
            highlightthickness=1,
            highlightbackground=self.colors['primary']
        ).pack(side=tk.LEFT)

    def _build_results_panel(self, parent):
        """结果统计面板"""
        card = self._create_card(parent)

        tk.Label(
            card,
            text="📊 对比结果",
            font=('Microsoft YaHei UI', 11, 'bold'),
            fg=self.colors['text'],
            bg=self.colors['card_bg']
        ).pack(anchor=tk.W, padx=20, pady=(15, 10))

        stats_frame = tk.Frame(card, bg=self.colors['card_bg'])
        stats_frame.pack(fill=tk.X, padx=20, pady=(0, 15))

        # 差异数量
        diff_frame = tk.Frame(stats_frame, bg='#FFEBEE', relief=tk.FLAT, bd=0)
        diff_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        tk.Label(
            diff_frame,
            text="差异",
            font=('Microsoft YaHei UI', 9),
            fg=self.colors['error'],
            bg='#FFEBEE'
        ).pack(pady=(10, 0))

        self.diff_label = tk.Label(
            diff_frame,
            text="0",
            font=('Microsoft YaHei UI', 24, 'bold'),
            fg=self.colors['error'],
            bg='#FFEBEE'
        )
        self.diff_label.pack(pady=(0, 10))

        # 相同数量
        same_frame = tk.Frame(stats_frame, bg='#E8F5E9', relief=tk.FLAT, bd=0)
        same_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        tk.Label(
            same_frame,
            text="相同",
            font=('Microsoft YaHei UI', 9),
            fg=self.colors['success'],
            bg='#E8F5E9'
        ).pack(pady=(10, 0))

        self.same_label = tk.Label(
            same_frame,
            text="0",
            font=('Microsoft YaHei UI', 24, 'bold'),
            fg=self.colors['success'],
            bg='#E8F5E9'
        )
        self.same_label.pack(pady=(0, 10))

    def _build_log_panel(self, parent):
        """日志面板"""
        card = self._create_card(parent)

        tk.Label(
            card,
            text="📋 日志",
            font=('Microsoft YaHei UI', 11, 'bold'),
            fg=self.colors['text'],
            bg=self.colors['card_bg']
        ).pack(anchor=tk.W, padx=20, pady=(15, 10))

        log_frame = tk.Frame(card, bg=self.colors['card_bg'])
        log_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 15))

        self.log_text = tk.Text(
            log_frame,
            font=('Consolas', 9),
            bg='#FAFAFA',
            fg=self.colors['text'],
            relief=tk.FLAT,
            wrap=tk.WORD,
            height=60
        )
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)

        # 配置日志颜色
        self.log_text.tag_config('info', foreground=self.colors['text'])
        self.log_text.tag_config('success', foreground=self.colors['success'])
        self.log_text.tag_config('warning', foreground=self.colors['warning'])
        self.log_text.tag_config('error', foreground=self.colors['error'])
        self.log_text.tag_config('primary', foreground=self.colors['primary'])

    def _create_card(self, parent):
        """创建卡片容器"""
        card = tk.Frame(
            parent,
            bg=self.colors['card_bg'],
            relief=tk.FLAT,
            bd=0,
            highlightthickness=1,
            highlightbackground=self.colors['border']
        )
        card.pack(fill=tk.X, pady=(0, 15))
        return card

    def log(self, message, level='info'):
        """添加日志"""
        self.log_queue.put((message, level))

    def _update_log(self):
        """更新日志显示"""
        try:
            while True:
                message, level = self.log_queue.get_nowait()
                self.log_text.insert(tk.END, f"{message}\n", level)
                self.log_text.see(tk.END)
        except queue.Empty:
            pass
        self.root.after(100, self._update_log)

    def browse_package(self):
        """选择APK文件"""
        filename = filedialog.askopenfilename(
            title="选择APK文件",
            filetypes=[("APK files", "*.apk"), ("All files", "*.*")]
        )
        if filename:
            self.package_path.set(filename)

    def browse_excel(self):
        """选择Excel文件"""
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.excel_path.set(filename)
            self.load_excel_sheets(filename)

    def load_excel_sheets(self, excel_path):
        """加载Excel的Sheet列表"""
        try:
            sheets = self.core.get_excel_sheets(excel_path)
            self.sheet_combo['values'] = sheets
            if sheets:
                self.selected_sheet.set(sheets[0])
        except Exception as e:
            messagebox.showerror("错误", f"读取Excel失败: {e}")

    def start_decompile(self):
        """开始反编译APK"""
        if self.is_processing:
            messagebox.showwarning("警告", "正在处理中...")
            return

        apk_path = self.package_path.get()
        if not apk_path:
            messagebox.showerror("错误", "请选择APK文件")
            return

        self.is_processing = True
        self.decompile_btn.config(state=tk.DISABLED, bg='#BDBDBD')

        thread = threading.Thread(target=self._decompile_thread, args=(apk_path,))
        thread.daemon = True
        thread.start()

    def _decompile_thread(self, apk_path):
        """反编译线程"""
        try:
            self.log("=" * 60, 'primary')
            self.log("开始APK反编译", 'primary')
            self.log("=" * 60, 'primary')

            # 创建APKDecompiler，传递日志回调函数
            self.decompiler = APKDecompiler(log_callback=self.log)
            self.decompiler.apk_path = apk_path

            if self.decompiler.decompile():
                res_dir = self.decompiler.get_res_directory()
                if res_dir:
                    self.core.res_base_path = res_dir
                    self.root.after(0, lambda: self.res_base_path.set(res_dir))
                    self.log(f"\n✓ 反编译成功！res目录: {res_dir}", 'success')
                    self.root.after(0, lambda: messagebox.showinfo(
                        "成功", f"反编译完成！\nres目录: {res_dir}"
                    ))
                else:
                    self.log("✗ 未找到res目录", 'error')
            else:
                self.log("✗ 反编译失败", 'error')

        except Exception as e:
            self.log(f"✗ 反编译错误: {e}", 'error')
            import traceback
            self.log(traceback.format_exc(), 'error')
        finally:
            self.is_processing = False
            self.root.after(0, lambda: self.decompile_btn.config(
                state=tk.NORMAL, bg=self.colors['primary']
            ))

    def start_compare(self):
        """开始对比"""
        if self.is_processing:
            messagebox.showwarning("警告", "正在处理中...")
            return

        if not self.excel_path.get():
            messagebox.showerror("错误", "请选择Excel文件")
            return

        if not self.selected_sheet.get():
            messagebox.showerror("错误", "请选择Sheet")
            return

        self.is_processing = True
        self.compare_btn.config(state=tk.DISABLED, bg='#BDBDBD')

        thread = threading.Thread(target=self._compare_thread)
        thread.daemon = True
        thread.start()

    def _compare_thread(self):
        """对比线程"""
        try:
            self.log("=" * 60, 'primary')
            self.log("开始多语言对比", 'primary')
            self.log("=" * 60, 'primary')

            # 更新res路径
            self.core.res_base_path = self.res_base_path.get()

            # 加载Excel
            self.log("\n加载Excel文件...", 'primary')
            self.core.load_excel(self.excel_path.get(), self.selected_sheet.get())
            self.log("✓ Excel加载成功", 'success')

            # 获取keys
            keys = self.core.get_keys_from_excel()
            self.log(f"✓ 加载 {len(keys)} 个Key", 'success')

            # 获取国家列表
            countries = self.core.get_countries_from_excel()
            self.log(f"✓ 国家列表: {', '.join(countries)}", 'success')

            # 读取XML
            self.log("\n读取XML文件...", 'primary')
            xml_data, missing_keys = self.core.read_strings_from_xml()

            for country in countries:
                count = len(xml_data.get(country, {}))
                self.log(f"  {country}: {count} 项", 'info')

            if missing_keys:
                self.log(f"\n⚠ 警告: {len(missing_keys)} 个语言目录未找到", 'warning')
                for key in missing_keys:
                    self.log(f"  - {key}", 'warning')

            # 对比
            self.log("\n开始对比...", 'primary')
            diff_count, same_count = self.core.compare_and_generate_results()

            self.root.after(0, lambda: self.diff_label.config(text=str(diff_count)))
            self.root.after(0, lambda: self.same_label.config(text=str(same_count)))

            self.log("=" * 60, 'success')
            self.log(f"✓ 对比完成！差异: {diff_count}, 相同: {same_count}", 'success')
            self.log("=" * 60, 'success')

            self.root.after(0, lambda: messagebox.showinfo(
                "完成",
                f"对比完成！\n差异: {diff_count}\n相同: {same_count}\n\n结果已保存到Excel文件"
            ))

        except Exception as e:
            self.log(f"✗ 对比错误: {e}", 'error')
            import traceback
            self.log(traceback.format_exc(), 'error')
            self.root.after(0, lambda: messagebox.showerror("错误", f"对比失败: {e}"))
        finally:
            self.is_processing = False
            self.root.after(0, lambda: self.compare_btn.config(
                state=tk.NORMAL, bg=self.colors['success']
            ))

    def open_results(self):
        """打开结果文件"""
        if os.path.exists("对比差异结果.xlsx"):
            if sys.platform == 'win32':
                os.startfile("对比差异结果.xlsx")
            elif sys.platform == 'darwin':
                subprocess.run(['open', "对比差异结果.xlsx"])
            else:
                subprocess.run(['xdg-open', "对比差异结果.xlsx"])
        else:
            messagebox.showinfo("提示", "请先执行对比")


def main():
    # 如果有 tkinterdnd2，使用 TkinterDnD.Tk()，否则使用普通 tk.Tk()
    if HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()

    app = MultiLanguageGUI(root)

    # 居中窗口
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')

    root.mainloop()


if __name__ == '__main__':
    main()