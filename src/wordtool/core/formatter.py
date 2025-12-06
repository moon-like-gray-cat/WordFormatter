import re
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.shared import RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

TITLE_FORMATS = [
    "一", "一、", "（一）", "（一）、", "（一）.",
    "（1）", "（1）、", "（1）.", "1", "1.", "1、",
    "a", "a.", "A", "A.", "①", "I", "I.", "（I）"
]

_FORMAT_TO_REGEX = {
    # ------------------ 中文大写数字类 -------------------
    "一": r"^[一二三四五六七八九十]+\s*",  # 匹配 "一 " 或 "二 " (无标点)
    "一、": r"^[一二三四五六七八九十]+[、\.]\s*",  # 匹配 "一、", "二.", "三 " (带顿号或点号)
    "（一）": r"^（[一二三四五六七八九十]+）\s*",  # 匹配 "(一) ", "(二) "
    "（一）、": r"^（[一二三四五六七八九十]+）[、\.]?\s*",  # 匹配 "(一) 、", "(二) ", "(三)."
    "（一）\.": r"^（[一二三四五六七八九十]+）[、\.]\s*",  # 同上，匹配带点号或顿号

    # ------------------ 阿拉伯数字类 ----------------------
    "1": r"^\d+\s*",  # 匹配 "1 " 或 "2 " (无标点)
    "1.": r"^\d+[、\.]\s*",  # 匹配 "1.", "2、", "3 " (带点号或顿号)
    "1、": r"^\d+[、\.]\s*",  # 同上
    "（1）": r"^（\d+）\s*",  # 匹配 "(1) ", "(2) "
    "（1）、": r"^（\d+）[、\.]?\s*",  # 匹配 "(1) 、", "(2) ", "(3)."
    "（1）\.": r"^（\d+）[、\.]\s*",  # 同上

    # ------------------ 字母和罗马数字类 --------------------
    "a": r"^[a-z]{1,2}\s*",  # 匹配 "a " 或 "b "
    "a.": r"^[a-z]{1,2}[、\.]\s*",  # 匹配 "a.", "b、"
    "A": r"^[A-Z]{1,2}\s*",  # 匹配 "A " 或 "B "
    "A.": r"^[A-Z]{1,2}[、\.]\s*",  # 匹配 "A.", "B、"
    "I": r"^[IVXLCDM]+\s*",  # 匹配 "I " 或 "II "
    "I.": r"^[IVXLCDM]+[、\.]\s*",  # 匹配 "I.", "II、"
    "（I）": r"^（[IVXLCDM]+）\s*",  # 匹配 "(I) ", "(II) "

    # ------------------ 特殊符号类 -------------------------
    "①": r"^[①②③④⑤⑥⑦⑧⑨⑩]+\s*",  # 匹配带圈数字
}


class WordFormatter:
    def __init__(self, file_path, config: dict):
        self.file_path = file_path
        self.config = config
        self.titles = config.get("titles", {})
        self.body = config.get("body", {})
        self.figure = config.get("figure", {})
        self.table = config.get("table", {})



    # ----------------------------------------------------------------------
    # 设置文本 run 样式（图片 run 跳过）
    # ----------------------------------------------------------------------
    def _set_run_style(self, run, font_name, size, bold):
        if run._element.xpath(".//w:drawing"):
            return
        run.font.name = font_name
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        run.font.size = Pt(size)
        run.bold = bold
        run.font.color.rgb = RGBColor(0, 0, 0)

    # ----------------------------------------------------------------------
    # 标题层级检测
    # ----------------------------------------------------------------------
    def _detect_level(self, text):
        # 1. 预处理文本：标准化括号和去除隐藏字符
        normalized_text = self._normalize_brackets(text.strip())
        normalized_text = re.sub(r'^[\s\x00-\x1f]+', '', normalized_text)
        if not normalized_text:
            return 0

        for i in range(1, 5):
            key = f"title{i}"
            # 2. 从 JSON 配置中获取用户设定的标识 (例如 "（1）" 或 "1.")
            format_key = self.titles.get(key, {}).get("format", "")

            # 3. 查表获取对应的正则表达式字符串
            # 注意：这里我们使用全局的 _FORMAT_TO_REGEX 字典
            regex_pattern = _FORMAT_TO_REGEX.get(format_key)

            if regex_pattern:
                # 4. 使用 re.match 进行匹配（执行前缀匹配）
                try:
                    if re.match(regex_pattern, normalized_text):
                        return i
                except re.error as e:
                    # 提示配置中正则错误
                    print(f"Warning: Invalid regex pattern used for {key}: {e}")
                    continue
        return 0

    # ----------------------------------------------------------------------
    # 获取样式
    # ----------------------------------------------------------------------
    def _get_style(self, level):
        if level == 0:
            return self.body
        else:
            key = f"title{level}"
            return self.titles.get(key, self.body)

    # ----------------------------------------------------------------------
    # 应用样式到段落
    # ----------------------------------------------------------------------
    def _apply_style(self, paragraph, level, heading_style=True):
        style_cfg = self._get_style(level)
        font_name = style_cfg.get("font", "宋体")
        size_str = style_cfg.get("size", "12")
        m = re.search(r"\(([\d.]+)pt\)", size_str)
        size = float(m.group(1)) if m else 12
        bold = bool(style_cfg.get("bold", False))

        # 设置段落样式
        if heading_style and level > 0:
            # 使用 Word 内置 Heading 样式
            paragraph.style = f'Heading {level}'

        # 设置 run 样式
        for run in paragraph.runs:
            self._set_run_style(run, font_name, size, bold)

        # 设置正文行间距
        if level == 0 and self.body.get("line_rule"):
            spacing = float(self.body.get("spacing", "1.25"))
            if self.body.get("line_rule") == "多倍行距":
                paragraph.paragraph_format.line_spacing = spacing
            else:  # 固定值（磅）
                paragraph.paragraph_format.line_spacing_rule = None
                paragraph.paragraph_format.space_before = Pt(spacing)
                paragraph.paragraph_format.space_after = Pt(spacing)

    # ----------------------------------------------------------------------
    # 将英文括号转中文括号
    # ----------------------------------------------------------------------
    def _normalize_brackets(self, text):
        text = text.replace("(", "（").replace(")", "）")
        return text

    # ----------------------------------------------------------------------
    # 预处理图题和表题
    # ----------------------------------------------------------------------
    def _preprocess_captions(self, doc):
        """
        只处理已有图题和表题
        - 图片下方已有图题的段落，设置样式和居中
        - 表格上方已有表题的段落，设置样式和居中
        """
        paragraphs = doc.paragraphs
        for i, para in enumerate(paragraphs):
            # 图片下方图题
            if para._element.xpath(".//w:drawing"):
                # 检查下一段是否以“图”开头
                if i + 1 < len(paragraphs) and paragraphs[i + 1].text.strip().startswith("图"):
                    caption_para = paragraphs[i + 1]
                    self._apply_style(caption_para, level=0)
                    caption_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

            # 表格上方表题
            next_elem = para._element.getnext()
            if next_elem is not None and next_elem.tag.endswith("tbl"):
                # 检查当前段落是否以“表”开头
                if para.text.strip().startswith("表"):
                    caption_para = para
                    self._apply_style(caption_para, level=0)
                    caption_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # ----------------------------------------------------------------------
    # 保存文档
    # ----------------------------------------------------------------------
    def save(self, output_path):
        try:
            doc = Document(self.file_path)

            # 1. 预处理标题括号（只修改文字 run，不覆盖段落）
            for para in doc.paragraphs:
                for run in para.runs:
                    if not run._element.xpath(".//w:drawing"):
                        run.text = self._normalize_brackets(run.text)

            # 2. 处理已有图题和表题
            self._preprocess_captions(doc)

            # 3. 应用样式（标题和正文）
            for para in doc.paragraphs:
                level = self._detect_level(para.text)
                self._apply_style(para, level)

            doc.save(output_path)
            return True
        except Exception as e:
            print(f"Error saving document: {e}")
            return False
