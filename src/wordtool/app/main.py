from wordtool.app.ui_components import WordFormatterUI
from wordtool.app.event_handlers import EventHandlers

import sys,os
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

def main():
    ui = WordFormatterUI()
    handlers = EventHandlers(ui)
    ui.run()


if __name__ == "__main__":
    main()
