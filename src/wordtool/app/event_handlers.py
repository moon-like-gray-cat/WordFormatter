from tkinter import filedialog, messagebox
from pathlib import Path
from wordtool.config import ConfigManager
from wordtool.core.formatter import WordFormatter
import json
from tkinter import font as tkfont

class EventHandlers:
    def __init__(self, ui):
        self.ui = ui
        self.input_file = None
        self.output_dir = None
        self._bind_events()

    def _bind_events(self):
        self.ui.btn_choose.config(command=self.choose_file)
        self.ui.btn_output.config(command=self.choose_output_path)
        self.ui.btn_start.config(command=self.start_formatting)
        # 绑定导入/导出按钮 ---
        self.ui.btn_import.config(command=self.import_config)
        self.ui.btn_export.config(command=self.export_config)

    def check_fonts_installed(self, config):
        """检查配置中的字体在系统中是否存在"""
        # 获取系统已安装的所有字体名称
        installed_fonts = set(tkfont.families())

        # 收集配置中所有使用到的中文字体
        needed_fonts = set()

        # 1. 检查标题字体
        for title_cfg in config.get("titles", {}).values():
            needed_fonts.add(title_cfg.get("font"))

        # 2. 检查正文字体
        needed_fonts.add(config.get("body", {}).get("font"))

        # 3. 检查图表标题字体
        needed_fonts.add(config.get("caption", {}).get("font"))

        # 找出缺失的字体 (排除 None 或空字符串)
        missing_fonts = [f for f in needed_fonts if f and f not in installed_fonts]

        return missing_fonts

    def get_safe_config(self, config, missing_fonts):
        """
        核心逻辑：如果字体不存在，将其在配置副本中替换为'宋体'
        """
        import copy
        safe_cfg = copy.deepcopy(config)
        missing_set = set(missing_fonts)

        # 1. 修正标题字体
        for key in safe_cfg.get("titles", {}):
            if safe_cfg["titles"][key].get("font") in missing_set:
                safe_cfg["titles"][key]["font"] = "宋体"

        # 2. 修正正文
        if safe_cfg.get("body", {}).get("font") in missing_set:
            safe_cfg["body"]["font"] = "宋体"

        # 3. 修正图表标题
        if safe_cfg.get("caption", {}).get("font") in missing_set:
            safe_cfg["caption"]["font"] = "宋体"

        return safe_cfg
    def export_config(self):
        """将当前 UI 面板的配置导出为 JSON 文件"""
        file_path = filedialog.asksaveasfilename(
            title="导出配置文件",
            defaultextension=".json",
            filetypes=[("JSON Configuration", "*.json")]
        )
        if file_path:
            try:
                # 获取 UI 当前所有控件的值
                config = self.ui.get_config()
                with open(file_path, "w", encoding="utf-8") as f:
                    json.dump(config, f, indent=4, ensure_ascii=False)
                messagebox.showinfo("导出成功", f"配置已保存至：\n{file_path}")
            except Exception as e:
                messagebox.showerror("导出失败", f"错误详情：{str(e)}")

    def import_config(self):
        """从外部 JSON 文件加载配置到 UI 面板"""
        file_path = filedialog.askopenfilename(
            title="导入配置文件",
            filetypes=[("JSON Configuration", "*.json")]
        )
        if file_path:
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    config = json.load(f)

                # 更新 UI 类中的 config_data 并刷新界面
                self.ui.config_data = config
                self.ui._apply_config_to_ui()

                messagebox.showinfo("导入成功", "配置已加载到面板")
            except Exception as e:
                messagebox.showerror("导入失败", f"无效的 JSON 文件：{str(e)}")
    def choose_file(self):
        path = filedialog.askopenfilename(
            title="选择 Word 文件",
            filetypes=[("Word Document", "*.docx")]
        )
        if path:
            self.input_file = path
            messagebox.showinfo("选择成功", path)

    def choose_output_path(self):
        path = filedialog.askdirectory(title="选择输出路径")
        if path:
            self.output_dir = path
            messagebox.showinfo("输出路径已设置", path)

    def start_formatting(self):
        if not self.input_file:
            messagebox.showwarning("缺少文件", "请先选择 Word 文件")
            return

        if not self.output_dir:
            messagebox.showwarning("缺少路径", "请先选择输出路径")
            return

        # 从 UI 获取配置
        raw_config = self.ui.get_config()
        final_config = raw_config
        # --- 字体检测逻辑 ---
        missing = self.check_fonts_installed(raw_config)
        if missing:
            msg = "检测到以下字体未在系统中安装：\n\n"
            msg += "\n".join([f"· {f}" for f in missing])
            msg += "\n\n继续操作可能导致 Word 文档格式错乱（自动替换为宋体）。\n是否仍要继续？"

            # 使用 askyesno 询问用户，如果选“否”则中断操作
            if not messagebox.askyesno("字体缺失警告", msg):
                return
            final_config = self.get_safe_config(raw_config, missing)
            self.ui.config_data = final_config  # 更新 UI 对象的数据源
            self.ui._apply_config_to_ui()
        # ------------------------
        # 保存配置到 JSON
        ConfigManager.save_config(final_config)

        output_file = Path(self.output_dir) / ("格式化_" + Path(self.input_file).name)

        try:
            formatter = WordFormatter(self.input_file, final_config)
            formatter.save(str(output_file))
            messagebox.showinfo("完成", f"文件已保存：{output_file}")

        except Exception as e:
            messagebox.showerror("错误", str(e))
            raise e
