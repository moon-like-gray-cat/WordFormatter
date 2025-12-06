import json
from pathlib import Path

# JSON 配置文件路径
CONFIG_PATH = Path(__file__).resolve().parent / "resources" / "ui_config.json"


class ConfigManager:
    @staticmethod
    def load_config():
        """
        从 ui_config.json 读取配置；
        若不存在或 JSON 错误，则返回默认配置。
        """
        if CONFIG_PATH.exists():
            try:
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass  # JSON 错误则重置为默认配置

        # ========== 默认配置结构 ==========
        return {
            "titles": {
                "title1": {"format": "一、", "font": "黑体", "size": "四号 (14pt)", "bold": False},
                "title2": {"format": "（一）", "font": "黑体", "size": "五号 (10.5pt)", "bold": False},
                "title3": {"format": "1.", "font": "宋体", "size": "五号 (10.5pt)", "bold": False},
                "title4": {"format": "（1）", "font": "宋体", "size": "五号 (10.5pt)", "bold": False},
            },
            "body": {
                "font": "宋体",
                "size": "五号 (10.5pt)",
                "bold": False,
                "line_rule": "多倍行距",
                "spacing": "1.25",  # 1.25倍行距
            },
            "caption": {
                "font": "宋体",
                "size": "小五号 (9pt)",
                "bold": False,
                "line_rule": "多倍行距",
                "spacing": "1.25"
            }
        }

    @staticmethod
    def save_config(config: dict):
        """
        将 UI 配置保存到 ui_config.json。
        """
        CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
