# public_python
## package and project manager
* `uv` => [github](https://github.com/astral-sh/uv)
## NOTE
* When running GUI scripts on WSL2, the Japanese text in the window title is garbled. That is the way to deal with it.
```Python
import platform
from PySide6.QtGui import QFont, QFontDatabase
from PySide6.QtWidgets import QWidget
class SampleApp(QWidget):
    def __init__(self):
        super().__init__()
        # WSL-Ubuntuでフォント設定
        if platform.system() == "Linux":
            # install ipafont-gothic
            font_path = "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf"
            font_id = QFontDatabase.addApplicationFont(font_path)
            font_family = QFontDatabase.applicationFontFamilies(font_id)[0]
            font = QFont(font_family)
            self.setFont(font)
```
