from tkinter import filedialog, messagebox
from pathlib import Path
from wordtool.config import ConfigManager
from wordtool.core.formatter import WordFormatter


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
        config = self.ui.get_config()

        # 保存配置到 JSON
        ConfigManager.save_config(config)

        output_file = Path(self.output_dir) / ("格式化_" + Path(self.input_file).name)

        try:
            formatter = WordFormatter(self.input_file, config)
            formatter.save(str(output_file))
            messagebox.showinfo("完成", f"文件已保存：{output_file}")

        except Exception as e:
            messagebox.showerror("错误", str(e))
            raise e
