# public_python
## package and project manager
* `uv` => [github](https://github.com/astral-sh/uv)
  * ***Initial settings***
    ```Shell
    sudo apt install direnv
    vim ~/.zshrc
      # direnv
      eval "$(direnv hook zsh)"
      # uv
      # https://docs.astral.sh/uv/configuration/environment/#pythonpath
      export PYTHONPATH=$PYTHONPATH:"directory_path_of_root"
      # https://docs.astral.sh/uv/configuration/environment/#uv_http_timeout
      export UV_HTTP_TIMEOUT=300
      # https://docs.astral.sh/uv/configuration/environment/#ssl_cert_file
      export SSL_CERT_FILE="absolute_file_path_of_ca"
    mkdir "directory_path_of_workspace"
    cd "directory_path_of_workspace"
    uv init --name "project_name"
    uv venv
    vim .envrc
      # direnv
      source ./.venv/bin/activate
    direnv allow
    uv tool install "tool_name"
    uv add --dev "tool_name"
    uv sync
    uv lock
    ```
  * ***If you want to use jupyter***
    ```Shell
    cd "directory_path_of_workspace"
    uv add --dev pip ipykernel pandas numpy jupyter
    uv run ipython kernel install --user --name="project_name"
    ```
  * ***If you want to use QT***
    ```Shell
    vim ~/.zshrc
      # QT
      export QT_QPA_PLATFORM=xcb
    apt search libxcb
    sudo apt install "package_name"
    ```
  * ***To do after opening workspace***
    ```Shell
    cd "directory_path_of_workspace"
    uv self update
    uv tool upgrade --all
    uv sync
    uv lock
    ```
  * ***If you want to run scripts***
    ```Shell
    cd "directory_path_of_scripts"
    uv run "file_name_of_script"
    ```
  * ***NOTE***
      * Create `__init__.py` in all directories relative to the test scripts.
## tool
### python
* `ruff` => [github](https://github.com/astral-sh/ruff)
* `pytest` => [github](https://github.com/pytest-dev/pytest/)
* `PySide6` => [url](https://doc.qt.io/qtforpython-6/)
* `PyInstaller` => [url](https://pyinstaller.org/en/stable/)
* `markitdown` => [github](https://github.com/microsoft/markitdown)
* `pypdf` => [github](https://github.com/py-pdf/pypdf)
### others
* `direnv` => [github](https://github.com/direnv/direnv)
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
