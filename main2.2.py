import os
import pandas as pd
import json
import shutil
from difflib import SequenceMatcher
from typing import List, Tuple, Optional
from tkinter import (
    Tk,
    Label,
    Button,
    Entry,
    messagebox,
    Frame,
    Toplevel,
    Text,
    filedialog,
    StringVar,
)


CONFIG_FILE = "Project Classifier Config.json"
DEFAULT_RANGES = [
    (0, 500, "500万以内项目"),
    (500, 10000, "500万-1亿元项目"),
    (10000, 50000, "1-5亿元项目"),
    (50000, 1000000, "5-100亿元项目"),
]


class RangeConfigurator:
    """处理范围配置的逻辑"""

    def __init__(self, frame: Frame):
        self.entries = []
        self.create_widgets(frame)

    def create_widgets(self, frame: Frame):
        """创建范围输入组件"""
        for i, (min_val, max_val, dir_name) in enumerate(DEFAULT_RANGES):
            row = i
            min_var = StringVar(value=str(min_val))
            max_var = StringVar(value=str(max_val))
            dir_var = StringVar(value=dir_name)

            Label(frame, text=f"范围{i+1}:").grid(row=row, column=0, padx=5, pady=5)

            Entry(frame, textvariable=min_var, width=10).grid(row=row, column=1)
            Entry(frame, textvariable=max_var, width=10).grid(row=row, column=2)
            Entry(frame, textvariable=dir_var, width=30).grid(row=row, column=3)

            self.entries.append((min_var, max_var, dir_var))

    def get_ranges(self) -> Tuple[List[Tuple[float, float, str]], List[str]]:
        """获取验证后的范围配置"""
        ranges = []
        errors = []

        for min_var, max_var, dir_var in self.entries:
            min_val = min_var.get()
            max_val = max_var.get()
            dir_name = dir_var.get()

            # 验证逻辑
            if not min_val.isdigit():
                errors.append(f"范围最小值必须为数字: {min_val}")
                continue

            if max_val != "inf" and not max_val.isdigit():
                errors.append(f"范围最大值必须为数字或'inf': {max_val}")
                continue

            if not dir_name:
                errors.append("文件夹名称不能为空")
                continue

            min_num = int(min_val)
            max_num = float("inf") if max_val == "inf" else int(max_val)

            if min_num >= max_num:
                errors.append(f"范围设置错误: {min_num} >= {max_num}")
                continue

            ranges.append((min_num, max_num, dir_name))

        return ranges, errors


class FileClassifierApp(Tk):
    """主应用程序类"""

    def __init__(self):
        super().__init__()
        self.title("项目文件分类工具")
        self.geometry(self._center_geometry(400, 500))
        self.config = AppConfig()

        # 初始化UI组件
        self.file_frame = FileSelectorFrame(self)
        self.range_frame = RangeConfigFrame(self)
        self.control_frame = ControlFrame(self)

        self._load_config()

    def _center_geometry(self, width: int, height: int) -> str:
        """计算居中窗口位置"""
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        return f"+{x}+{y}"

    def _load_config(self):
        """加载保存的配置"""
        if self.config.excel_path:
            self.file_frame.set_path(self.config.excel_path)
        for i, (min_val, max_val, dir_name) in enumerate(self.config.ranges):
            if i < len(self.range_frame.ranges):
                min_var, max_var, dir_var = self.range_frame.ranges[i]
                min_var.set(min_val)
                max_var.set(max_val)
                dir_var.set(dir_name)

    def classify_files(self):
        """执行文件分类"""
        excel_path = self.file_frame.get_path()
        if not excel_path:
            messagebox.showerror("错误", "请先选择Excel文件")
            return

        ranges, errors = self.range_frame.get_validated_ranges()
        if errors:
            messagebox.showerror("输入错误", "\n".join(errors))
            return

        try:
            classifier = FileClassifier(
                excel_path=excel_path,
                base_dir=os.path.dirname(excel_path),
                ranges=ranges,
            )
            classifier.run()
            self.config.save(excel_path, ranges)
            messagebox.showinfo("成功", "文件分类完成！")
        except Exception as e:
            messagebox.showerror("运行错误", str(e))


class FileClassifier:
    """处理文件分类的核心逻辑"""

    def __init__(
        self, excel_path: str, base_dir: str, ranges: List[Tuple[float, float, str]]
    ):
        self.excel_path = excel_path
        self.base_dir = base_dir
        self.ranges = ranges
        self._validate_paths()

    def _validate_paths(self):
        """路径验证"""
        if not os.path.isfile(self.excel_path):
            raise FileNotFoundError(f"Excel文件不存在: {self.excel_path}")

        if not os.path.isdir(self.base_dir):
            raise NotADirectoryError(f"目录不存在: {self.base_dir}")

    def _find_project_folder(self, project_name: str) -> Optional[str]:
        """查找匹配的项目文件夹"""
        for folder in os.listdir(self.base_dir):
            folder_path = os.path.join(self.base_dir, folder)
            if os.path.isdir(folder_path):
                similarity = SequenceMatcher(None, project_name, folder).ratio()
                if similarity > 0.8:
                    return folder_path
        return None

    def _get_target_dir(self, project_size: float) -> Optional[str]:
        """获取目标目录"""
        for min_size, max_size, dir_name in self.ranges:
            if min_size <= project_size < max_size:
                return os.path.join(self.base_dir, dir_name)
        return None

    def run(self):
        """执行分类操作"""
        try:
            df = pd.read_excel(self.excel_path, header=None)
        except Exception as e:
            raise ValueError(f"Excel文件读取失败: {e}")

        for _, row in df.iterrows():
            project_name = str(row[0])
            try:
                project_size = float(row[1])
            except (ValueError, IndexError):
                continue

            if folder_path := self._find_project_folder(project_name):
                if target_dir := self._get_target_dir(project_size):
                    os.makedirs(target_dir, exist_ok=True)
                    shutil.move(folder_path, target_dir)


class AppConfig:
    """处理应用程序配置"""

    def __init__(self):
        self.excel_path = ""
        self.ranges = DEFAULT_RANGES.copy()
        self.load()

    def load(self):
        """加载配置文件"""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.excel_path = data.get("excel_path", "")
                    self.ranges = [
                        (item[0], item[1], item[2]) for item in data.get("ranges", [])
                    ]
            except Exception as e:
                print(f"配置加载失败: {e}")

    def save(self, excel_path: str, ranges: list):
        """保存配置文件"""
        data = {"excel_path": excel_path, "ranges": ranges}
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"配置保存失败: {e}")


# GUI组件 --------------------------------------------------


class FileSelectorFrame(Frame):
    """文件选择组件"""

    def __init__(self, master):
        super().__init__(master)
        self.path_var = StringVar()
        self._create_widgets()
        self.pack(pady=10)

    def _create_widgets(self):
        Label(self, text="选择包含项目信息的Excel文件").grid(
            row=0, column=0, columnspan=2, pady=(0, 5)
        )
        Entry(self, textvariable=self.path_var, width=50).grid(row=1, column=0, padx=5)
        Button(self, text="选择文件", command=self._select_file).grid(
            row=1, column=1, padx=5
        )

    def _select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx;*.xls")])
        if path:
            self.path_var.set(path)

    def get_path(self) -> str:
        return self.path_var.get()

    def set_path(self, path: str):
        self.path_var.set(path)


class RangeConfigFrame(Frame):
    """范围配置组件"""

    def __init__(self, master):
        super().__init__(master)
        self.ranges = []
        self._create_widgets()
        self.pack(pady=10)

    def _create_widgets(self):
        Label(self, text="项目范围配置").grid(row=0, columnspan=4)
        configurator = RangeConfigurator(self)
        self.ranges = configurator.entries

    def get_validated_ranges(self) -> Tuple[list, list]:
        return RangeConfigurator(self).get_ranges()


class ControlFrame(Frame):
    """控制按钮组件"""

    def __init__(self, master):
        super().__init__(master)
        self._create_widgets()
        self.pack(pady=10)

    def _create_widgets(self):
        Button(self, text="执行分类", command=self.master.classify_files).grid(
            row=0, column=0, padx=10
        )
        Button(self, text="重置配置", command=self._reset).grid(
            row=0, column=1, padx=10
        )
        Button(self, text="帮助", command=self._show_help).grid(
            row=0, column=2, padx=10
        )

    def _reset(self):
        self.master.config = AppConfig()
        self.master._load_config()

    def _show_help(self):
        help_win = Toplevel(self.master)
        help_win.title("帮助文档")

        # 设置帮助窗口的大小
        help_win.geometry("400x250")

        # 获取主窗口的位置和大小
        main_x = self.master.winfo_x()
        main_y = self.master.winfo_y()
        main_width = self.master.winfo_width()
        main_height = self.master.winfo_height()

        # 计算帮助窗口的位置，使其居中于主窗口
        help_x = main_x + (main_width // 2) - 200  # 200 是帮助窗口宽度的一半
        help_y = main_y + (main_height // 2) - 125  # 125 是帮助窗口高度的一半

        help_win.geometry(f"+{help_x}+{help_y}")

        # 添加帮助内容
        help_text = Text(help_win, wrap="word", width=60, height=16)
        help_text.pack(padx=10, pady=10)
        help_text.insert(
            "1.0",
            "帮助内容:\n\n"
            "   1. 点击“选择文件”按钮手动选择文件\n"
            "   2. 输入项目规模范围和对应的文件夹名称\n"
            "   3. 点击“执行项目分类”按钮进行分类\n"
            "   4. 点击“重置文本框”按钮重置输入框内容\n\n"
            "注意:\n\n"
            "   - 范围必须是数字\n"
            "   - 左区间不能大于右区间\n"
            "   - Excel文件第1、2列内容是项目名称、投资额（万元）"
            "\n\n\tv2.2 by HANJ 20250307\n"
            "\thanj-cn@qq.com",
        )
        help_text.config(state="disabled")


if __name__ == "__main__":
    app = FileClassifierApp()
    app.mainloop()
