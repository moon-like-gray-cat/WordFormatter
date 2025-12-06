import json
import os
import sys
import tkinter as tk
from tkinter import ttk
from pathlib import Path
from ..config import ConfigManager
from ..config import CONFIG_PATH


def resource_path(relative_path: str) -> str:
    """
    获取资源的绝对路径
    PyInstaller 打包后使用 sys._MEIPASS
    """
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)


class WordFormatterUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Word 格式化工具")
        self.root.geometry("1000x650")  # 增加窗口尺寸

        # 设置图标
        icon_path = resource_path("icon.ico")
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)
        else:
            print(f"图标文件不存在: {icon_path}")

        # 可识别标题格式
        self.title_formats = [
            "一", "一、", "（一）", "（一）、", "（一）.",
            "（1）", "（1）、", "（1）.", "1", "1.", "1、",
            "a", "a.", "A", "A.", "①", "I", "I.", "（I）"
        ]

        # 字号映射
        self.font_size_map = {
            "三号 (16pt)": 16,
            "小三号 (15pt)": 15,
            "四号 (14pt)": 14,
            "小四号 (12pt)": 12,
            "五号 (10.5pt)": 10.5,
            "小五号 (9pt)": 9,
        }

        # 存放控件
        self.title_widgets = {}
        self.fig_title = {}
        self.table_title = {}

        # 加载配置
        self.config_data = self._load_config()

        # 构建UI
        self._build_ui()
        self._apply_config_to_ui()

    # ---------------------- 配置读写 ----------------------
    def _load_config(self):
        """
        从 ConfigManager 加载配置。
        ConfigManager 保证一定返回有效配置。
        """
        return ConfigManager.load_config()

    def save_config(self, cfg):
        CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=4, ensure_ascii=False)

    # ---------------------- UI 构建 ----------------------
    def _build_ui(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill="both", expand=True)

        # 调整列权重，使两侧均匀分布
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)

        self._build_left(main_frame)
        self._build_right(main_frame)
        self._build_bottom(main_frame)

    # 左侧：标题/正文
    def _build_left(self, parent):
        lf = ttk.LabelFrame(parent, text="标题 / 正文设置", padding=10)
        lf.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        # 标题1-4
        def add_title_row(level_name, key):
            row = ttk.Frame(lf)
            row.pack(fill="x", pady=6)

            ttk.Label(row, text=level_name, width=8).pack(side="left")
            ttk.Label(row, text="格式").pack(side="left", padx=(5, 2))
            cb_format = ttk.Combobox(row, values=self.title_formats, width=8)
            cb_format.pack(side="left", padx=2)

            ttk.Label(row, text="字体").pack(side="left", padx=(5, 2))
            cb_font = ttk.Combobox(row, values=["宋体", "黑体", "微软雅黑", "楷体"], width=10)
            cb_font.pack(side="left", padx=2)

            ttk.Label(row, text="字号").pack(side="left", padx=(5, 2))
            cb_size = ttk.Combobox(row, values=list(self.font_size_map.keys()), width=12)
            cb_size.pack(side="left", padx=2)

            bold_var = tk.BooleanVar()
            ttk.Checkbutton(row, text="B", variable=bold_var).pack(side="left", padx=(5, 0))

            self.title_widgets[key] = {
                "format": cb_format,
                "font": cb_font,
                "size": cb_size,
                "bold": bold_var
            }

        for i in range(1, 5):
            add_title_row(f"{i}级标题:", f"title{i}")

        # 正文
        ttk.Label(lf, text="正文设置", font=("微软雅黑", 10, "bold")).pack(anchor="w", pady=(10, 5))
        bf = ttk.Frame(lf)
        bf.pack(fill="x", pady=5)

        ttk.Label(bf, text="字体").pack(side="left")
        self.body_font = ttk.Combobox(bf, values=["宋体", "黑体", "微软雅黑"], width=10)
        self.body_font.pack(side="left", padx=5)

        ttk.Label(bf, text="字号").pack(side="left", padx=(5, 2))
        self.body_size = ttk.Combobox(bf, values=list(self.font_size_map.keys()), width=12)
        self.body_size.pack(side="left", padx=5)

        self.body_bold = tk.BooleanVar()
        ttk.Checkbutton(bf, text="B", variable=self.body_bold).pack(side="left", padx=5)

        # 行距
        bf2 = ttk.Frame(lf)
        bf2.pack(fill="x", pady=5)
        ttk.Label(bf2, text="正文行距").pack(side="left")
        self.body_line_rule = ttk.Combobox(bf2, values=["固定值", "多倍行距"], width=10)
        self.body_line_rule.pack(side="left", padx=5)
        self.body_spacing = ttk.Entry(bf2, width=8)
        self.body_spacing.pack(side="left", padx=5)
        ttk.Label(bf2, text="默认多倍行距 1.25").pack(side="left", padx=5)

    # 右侧：图表标题设置（合并图题和表题）
    def _build_right(self, parent):
        rf = ttk.LabelFrame(parent, text="图表标题设置", padding=10)
        rf.grid(row=0, column=1, sticky="nsew")

        # 让右侧框架内部可以扩展
        rf.grid_columnconfigure(0, weight=1)

        # 使用网格布局替代pack布局，以便更好控制
        row_frame = ttk.Frame(rf)
        row_frame.grid(row=0, column=0, sticky="ew", pady=10)
        row_frame.grid_columnconfigure(1, weight=1)

        # 标签 - 使用网格布局确保对齐
        ttk.Label(row_frame, text="图表标题", width=10).grid(row=0, column=0, sticky="w", padx=(0, 5))

        # 字体设置
        ttk.Label(row_frame, text="字体").grid(row=0, column=1, sticky="w", padx=(0, 2))
        self.caption_font = ttk.Combobox(row_frame, values=["宋体", "黑体", "微软雅黑", "楷体"], width=10)
        self.caption_font.grid(row=0, column=2, sticky="w", padx=2)

        # 字号设置
        ttk.Label(row_frame, text="字号").grid(row=0, column=3, sticky="w", padx=(10, 2))
        self.caption_size = ttk.Combobox(row_frame, values=list(self.font_size_map.keys()), width=12)
        self.caption_size.grid(row=0, column=4, sticky="w", padx=2)

        # 加粗
        self.caption_bold = tk.BooleanVar()
        ttk.Checkbutton(row_frame, text="B", variable=self.caption_bold).grid(row=0, column=5, sticky="w", padx=(10, 2))

        # 第二行：行距设置
        row_frame2 = ttk.Frame(rf)
        row_frame2.grid(row=1, column=0, sticky="ew", pady=10)
        row_frame2.grid_columnconfigure(1, weight=1)

        ttk.Label(row_frame2, text="行距设置", width=10).grid(row=0, column=0, sticky="w", padx=(0, 5))

        ttk.Label(row_frame2, text="类型").grid(row=0, column=1, sticky="w", padx=(0, 2))
        self.caption_line_rule = ttk.Combobox(row_frame2, values=["固定值", "多倍行距"], width=10)
        self.caption_line_rule.grid(row=0, column=2, sticky="w", padx=2)

        ttk.Label(row_frame2,).grid(row=0, column=3, sticky="w", padx=(10, 2))
        self.caption_spacing = ttk.Entry(row_frame2, width=8)
        self.caption_spacing.grid(row=0, column=4, sticky="w", padx=2)
        ttk.Label(row_frame2, ).grid(row=0, column=5, sticky="w", padx=2)

        # 添加说明文本
        note_label = ttk.Label(rf, text="注：图表标题包括图题和表题，统一设置格式", font=("", 9))
        note_label.grid(row=2, column=0, sticky="w", pady=(10, 0))

    # 底部按钮
    def _build_bottom(self, parent):
        bottom = ttk.Frame(parent, padding=10)
        bottom.grid(row=1, column=0, columnspan=2, sticky="ew")

        # 配置底部框架的列权重
        bottom.grid_columnconfigure(0, weight=1)
        bottom.grid_columnconfigure(1, weight=1)

        # 左侧按钮
        left_btn_frame = ttk.Frame(bottom)
        left_btn_frame.grid(row=0, column=0, sticky="w")

        self.btn_choose = ttk.Button(left_btn_frame, text="选择文件")
        self.btn_choose.pack(side="left", padx=5)
        self.btn_output = ttk.Button(left_btn_frame, text="输出路径")
        self.btn_output.pack(side="left", padx=5)

        # 右侧按钮
        right_btn_frame = ttk.Frame(bottom)
        right_btn_frame.grid(row=0, column=1, sticky="e")

        self.btn_start = ttk.Button(right_btn_frame, text="开始格式化", width=20)
        self.btn_start.pack(side="right", padx=5)

    # ---------------------- 填充配置 ----------------------
    def _apply_config_to_ui(self):
        cfg = self.config_data

        # 标题
        for key, w in self.title_widgets.items():
            if key in cfg["titles"]:
                wcfg = cfg["titles"][key]
                w["format"].set(wcfg.get("format", ""))
                w["font"].set(wcfg.get("font", "宋体"))
                w["size"].set(wcfg.get("size", "小四号 (12pt)"))
                w["bold"].set(wcfg.get("bold", False))

        # 正文
        body_cfg = cfg["body"]
        self.body_font.set(body_cfg.get("font", "宋体"))
        self.body_size.set(body_cfg.get("size", "小四号 (12pt)"))
        self.body_bold.set(body_cfg.get("bold", False))
        self.body_line_rule.set(body_cfg.get("line_rule", "多倍行距"))
        self.body_spacing.delete(0, "end")
        self.body_spacing.insert(0, body_cfg.get("spacing", "1.25"))

        # 图表标题
        caption_cfg = cfg.get("caption", {})
        self.caption_font.set(caption_cfg.get("font", "宋体"))
        self.caption_size.set(caption_cfg.get("size", "小五号 (9pt)"))
        self.caption_bold.set(caption_cfg.get("bold", False))
        self.caption_line_rule.set(caption_cfg.get("line_rule", "多倍行距"))
        self.caption_spacing.delete(0, "end")
        self.caption_spacing.insert(0, caption_cfg.get("spacing", "1.25"))

    def get_config(self):
        cfg = {
            "titles": {},
            "body": {},
            "caption": {}
        }

        # 标题 1-4
        for key, w in self.title_widgets.items():
            cfg["titles"][key] = {
                "format": w["format"].get(),
                "font": w["font"].get(),
                "size": w["size"].get(),
                "bold": w["bold"].get()
            }

        # 正文
        cfg["body"] = {
            "font": self.body_font.get(),
            "size": self.body_size.get(),
            "bold": self.body_bold.get(),
            "line_rule": self.body_line_rule.get(),
            "spacing": self.body_spacing.get()
        }

        # 图表标题
        cfg["caption"] = {
            "font": self.caption_font.get(),
            "size": self.caption_size.get(),
            "bold": self.caption_bold.get(),
            "line_rule": self.caption_line_rule.get(),
            "spacing": self.caption_spacing.get()
        }

        return cfg

    def run(self):
        self.root.mainloop()