# public_python
## WSL2(Ubuntu)
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
    ```
  * ***To do when you want to use any tools***
    ```Shell
    cd "directory_path_of_workspace"
    uv tool install "tool_name"
    uv add --dev "tool_name"
    uv sync
    uv lock
    ```
  * ***To do after opening workspace***
    ```Shell
    cd "directory_path_of_workspace"
    uv self update
    uv tool upgrade --all
    uv sync
    uv lock
    ```
  * ***To do when you want to run scripts***
    ```Shell
    cd "directory_path_of_scripts"
    uv run "file_name_of_script"
    ```
  * ***To do when you want to use jupyter***
    ```Shell
    cd "directory_path_of_workspace"
    uv add --dev pip ipykernel pandas numpy jupyter
    uv run ipython kernel install --user --name="project_name"
    ```
  * ***To do when you want to use PySide6***
    ```Shell
    vim ~/.zshrc
      # QT
      export QT_QPA_PLATFORM=xcb
    ```
    * When you run a GUI program using PySide6, you get an error
      ```Shell
      uv run "file_name_of_script"
        # qt.qpa.plugin: From 6.5.0, xcb-cursor0 or libxcb-cursor0 is needed to load the Qt xcb platform plugin.
        # qt.qpa.plugin: Could not load the Qt platform plugin "xcb" in "" even though it was found.
        # This application failed to start because no Qt platform plugin could be initialized. Reinstalling the application may fix this problem.
        # Available platform plugins are: eglfs, wayland-egl, vnc, wayland, minimal, offscreen, vkkhrdisplay, linuxfb, minimalegl, xcb.
      ```
    * Please debug
      ```Shell
      QT_DEBUG_PLUGINS=1 uv run "file_name_of_script"
        # qt.core.library: "/home/b21k8002/dev/projects/public_python/.venv/lib/python3.13/site-packages/PySide6/Qt/plugins/platforms/libqxcb.so" cannot load: Cannot load library /home/b21k8002/dev/projects/public_python/.venv/lib/python3.13/site-packages/PySide6/Qt/plugins/platforms/libqxcb.so: libxkbcommon-x11.so.0: 共有オブジェクトファイルを開けません: そのようなファイルやディレクトリはありません
        # qt.core.plugin.loader: QLibraryPrivate::loadPlugin failed on "/home/b21k8002/dev/projects/public_python/.venv/lib/python3.13/site-packages/PySide6/Qt/plugins/platforms/libqxcb.so" : "Cannot load library /home/b21k8002/dev/projects/public_python/.venv/lib/python3.13/site-packages/PySide6/Qt/plugins/platforms/libqxcb.so: libxkbcommon-x11.so.0: 共有オブジェクトファイルを開けません: そのようなファイルやディレクトリはありません"
      ```
    * Please check the referenced libraries
      ```Shell
      ldd /home/b21k8002/dev/projects/public_python/.venv/lib/python3.13/site-packages/PySide6/Qt/plugins/platforms/libqxcb.so
        # linux-vdso.so.1 (0x00007ffe323d6000)
        # libQt6XcbQpa.so.6 => /home/b21k8002/dev/projects/public_python/.venv/lib/python3.13/site-packages/PySide6/Qt/plugins/platforms/../../lib/libQt6XcbQpa.so.6 (0x00007fb552ce4000)
        # libQt6Gui.so.6 => /home/b21k8002/dev/projects/public_python/.venv/lib/python3.13/site-packages/PySide6/Qt/plugins/platforms/../../lib/libQt6Gui.so.6 (0x00007fb552207000)
        # libGL.so.1 => /lib/x86_64-linux-gnu/libGL.so.1 (0x00007fb552178000)
        # libQt6Core.so.6 => /home/b21k8002/dev/projects/public_python/.venv/lib/python3.13/site-packages/PySide6/Qt/plugins/platforms/../../lib/libQt6Core.so.6 (0x00007fb551a6d000)
        # libxkbcommon.so.0 => /lib/x86_64-linux-gnu/libxkbcommon.so.0 (0x00007fb551a22000)
        # libxkbcommon-x11.so.0 => not found
        # libxcb-cursor.so.0 => /lib/x86_64-linux-gnu/libxcb-cursor.so.0 (0x00007fb551a1b000)
        # libxcb-icccm.so.4 => not found
        # libxcb-image.so.0 => /lib/x86_64-linux-gnu/libxcb-image.so.0 (0x00007fb551a15000)
        # libxcb-keysyms.so.1 => not found
        # libxcb-randr.so.0 => /lib/x86_64-linux-gnu/libxcb-randr.so.0 (0x00007fb551a02000)
        # libxcb-render-util.so.0 => /lib/x86_64-linux-gnu/libxcb-render-util.so.0 (0x00007fb5519fb000)
        # libxcb-shm.so.0 => /lib/x86_64-linux-gnu/libxcb-shm.so.0 (0x00007fb5519f6000)
        # libxcb-sync.so.1 => /lib/x86_64-linux-gnu/libxcb-sync.so.1 (0x00007fb5519ed000)
        # libxcb-xfixes.so.0 => /lib/x86_64-linux-gnu/libxcb-xfixes.so.0 (0x00007fb5519e3000)
        # libxcb-render.so.0 => /lib/x86_64-linux-gnu/libxcb-render.so.0 (0x00007fb5519d4000)
        # libxcb-shape.so.0 => not found
        # libxcb-xkb.so.1 => not found
        # libxcb.so.1 => /lib/x86_64-linux-gnu/libxcb.so.1 (0x00007fb5519a9000)
        # libX11-xcb.so.1 => /lib/x86_64-linux-gnu/libX11-xcb.so.1 (0x00007fb5519a4000)
        # libX11.so.6 => /lib/x86_64-linux-gnu/libX11.so.6 (0x00007fb551867000)
        # libdl.so.2 => /lib/x86_64-linux-gnu/libdl.so.2 (0x00007fb551860000)
        # libpthread.so.0 => /lib/x86_64-linux-gnu/libpthread.so.0 (0x00007fb55185b000)
        # libstdc++.so.6 => /lib/x86_64-linux-gnu/libstdc++.so.6 (0x00007fb5515dd000)
        # libm.so.6 => /lib/x86_64-linux-gnu/libm.so.6 (0x00007fb5514f4000)
        # libgcc_s.so.1 => /lib/x86_64-linux-gnu/libgcc_s.so.1 (0x00007fb5514c6000)
        # libc.so.6 => /lib/x86_64-linux-gnu/libc.so.6 (0x00007fb5512b4000)
        # libxcb-icccm.so.4 => not found
        # libxcb-keysyms.so.1 => not found
        # libxcb-shape.so.0 => not found
        # libxcb-xkb.so.1 => not found
        # libglib-2.0.so.0 => /lib/x86_64-linux-gnu/libglib-2.0.so.0 (0x00007fb551169000)
        # libxkbcommon-x11.so.0 => not found
        # libgthread-2.0.so.0 => /lib/x86_64-linux-gnu/libgthread-2.0.so.0 (0x00007fb551162000)
        # libEGL.so.1 => /lib/x86_64-linux-gnu/libEGL.so.1 (0x00007fb551150000)
        # libfontconfig.so.1 => /lib/x86_64-linux-gnu/libfontconfig.so.1 (0x00007fb5510ff000)
        # libQt6DBus.so.6 => /home/b21k8002/dev/projects/public_python/.venv/lib/python3.13/site-packages/PySide6/Qt/plugins/platforms/../../lib/libQt6DBus.so.6 (0x00007fb551028000)
        # libz.so.1 => /lib/x86_64-linux-gnu/libz.so.1 (0x00007fb55100a000)
        # libfreetype.so.6 => /lib/x86_64-linux-gnu/libfreetype.so.6 (0x00007fb550f3e000)
        # libGLdispatch.so.0 => /lib/x86_64-linux-gnu/libGLdispatch.so.0 (0x00007fb550e86000)
        # libGLX.so.0 => /lib/x86_64-linux-gnu/libGLX.so.0 (0x00007fb550e53000)
        # libicui18n.so.73 => /home/b21k8002/dev/projects/public_python/.venv/lib/python3.13/site-packages/PySide6/Qt/plugins/platforms/../../lib/libicui18n.so.73 (0x00007fb550b04000)
        # libicuuc.so.73 => /home/b21k8002/dev/projects/public_python/.venv/lib/python3.13/site-packages/PySide6/Qt/plugins/platforms/../../lib/libicuuc.so.73 (0x00007fb5508ea000)
        # libicudata.so.73 => /home/b21k8002/dev/projects/public_python/.venv/lib/python3.13/site-packages/PySide6/Qt/plugins/platforms/../../lib/libicudata.so.73 (0x00007fb54ea5c000)
        # libzstd.so.1 => /lib/x86_64-linux-gnu/libzstd.so.1 (0x00007fb54e9a2000)
        # librt.so.1 => /lib/x86_64-linux-gnu/librt.so.1 (0x00007fb54e99d000)
        # /lib64/ld-linux-x86-64.so.2 (0x00007fb552dae000)
        # libxcb-util.so.1 => /lib/x86_64-linux-gnu/libxcb-util.so.1 (0x00007fb54e995000)
        # libXau.so.6 => /lib/x86_64-linux-gnu/libXau.so.6 (0x00007fb54e98d000)
        # libXdmcp.so.6 => /lib/x86_64-linux-gnu/libXdmcp.so.6 (0x00007fb54e985000)
        # libpcre2-8.so.0 => /lib/x86_64-linux-gnu/libpcre2-8.so.0 (0x00007fb54e8eb000)
        # libexpat.so.1 => /lib/x86_64-linux-gnu/libexpat.so.1 (0x00007fb54e8bf000)
        # libdbus-1.so.3 => /lib/x86_64-linux-gnu/libdbus-1.so.3 (0x00007fb54e870000)
        # libbz2.so.1.0 => /lib/x86_64-linux-gnu/libbz2.so.1.0 (0x00007fb54e85a000)
        # libpng16.so.16 => /lib/x86_64-linux-gnu/libpng16.so.16 (0x00007fb54e822000)
        # libbrotlidec.so.1 => /lib/x86_64-linux-gnu/libbrotlidec.so.1 (0x00007fb54e814000)
        # libbsd.so.0 => /lib/x86_64-linux-gnu/libbsd.so.0 (0x00007fb54e7fe000)
        # libsystemd.so.0 => /lib/x86_64-linux-gnu/libsystemd.so.0 (0x00007fb54e71e000)
        # libbrotlicommon.so.1 => /lib/x86_64-linux-gnu/libbrotlicommon.so.1 (0x00007fb54e6f9000)
        # libmd.so.0 => /lib/x86_64-linux-gnu/libmd.so.0 (0x00007fb54e6ea000)
        # libcap.so.2 => /lib/x86_64-linux-gnu/libcap.so.2 (0x00007fb54e6dd000)
        # libgcrypt.so.20 => /lib/x86_64-linux-gnu/libgcrypt.so.20 (0x00007fb54e595000)
        # liblz4.so.1 => /lib/x86_64-linux-gnu/liblz4.so.1 (0x00007fb54e573000)
        # liblzma.so.5 => /lib/x86_64-linux-gnu/liblzma.so.5 (0x00007fb54e53f000)
        # libgpg-error.so.0 => /lib/x86_64-linux-gnu/libgpg-error.so.0 (0x00007fb54e51a000)
      ```
    * If you extract "not found" from the above, it will be as follows
      ```Shell
        # libxkbcommon-x11.so.0 => not found => libxkbcommon-x11-0
        # libxcb-icccm.so.4 => not found => libxcb-icccm4
        # libxcb-keysyms.so.1 => not found => libxcb-keysyms1
        # libxcb-shape.so.0 => not found => libxcb-shape0
        # libxcb-xkb.so.1 => not found => libxcb-xkb1
      ```
    * Please install the above library all at once
      ```Shell
      sudo apt install libxkbcommon-x11-0 libxcb-icccm4 libxcb-keysyms1 libxcb-shape0 libxcb-xkb1
      ```
    * Please install Japanese fonts
      ```Shell
      sudo apt install fonts-ipafont
      ```
  * ***To do When running GUI scripts, the Japanese text in the window title is garbled***
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
  * ***NOTE***
      * Create `__init__.py` in all directories relative to the test scripts.
## Windows
  * ***Initial settings***
    ```PowerShell
    # Windows
    winget search --name python
    winget install --id Python.Python.X.X
    # Add the folder path of Python to the system environment variable PATH.
    $NewPath = "C:\your\path\here"
    $OldPath = [Environment]::GetEnvironmentVariable("Path", [EnvironmentVariableTarget]::Machine)
    $UpdatedPath = "$OldPath;$NewPath"
    [Environment]::SetEnvironmentVariable("Path", $UpdatedPath, [EnvironmentVariableTarget]::Machine)
    [Environment]::GetEnvironmentVariable("Path", [EnvironmentVariableTarget]::Machine)
    python --version
    ```
  * ***To do when you want to use any tools***
    ```Shell
    cd "directory_path_of_workspace"
    pip install "tool_name"
    ```
  * ***To do when you want to run scripts***
    ```Shell
    cd "directory_path_of_scripts"
    python "file_name_of_script"
    ```
  * ***To do when you want to install pyinstaller to pip***
    ```Shell
    # WSL2
    cd ~
    mkdir pocket
    cd pocket
    git clone https://github.com/pyinstaller/pyinstaller.git
    cd pyinstaller
    git submodule update --init --recursive
    ```
    ```PowerShell
    # Windows
    cd \\wsl$\Ubuntu\home\"username"\pocket\pyinstaller
    pip install .
    pyinstaller --version
    ```
    ```Shell
    # WSL2
    rm -r ~/pocket/pyinstaller
    ```
  * ***To do when you want to create and use an executable file***
    ```PowerShell
    # Windows
    cd "folderpath"
    pyinstaller --onefile --noconsole "file_path_of_script"
    cd dist
    ls .
    # You can run the displayed exe file by double-clicking on it.
    ```
  * ***To do when you want to create and apply an icon for the executable file***
    ```PowerShell
    # Windows
    winget search --name inkscape
    winget install --id Inkscape.Inkscape
    # Create and save svg file
    winget search --name imagemagick
    winget install --id ImageMagick.ImageMagick
    # Convert svg to ico
    # Recommended size: Including 256x256 for support of high-resolution display. icon.ico will be the final icon for PyInstaller.
    magick "file_path_of_svg" -define icon:auto-resize=256,128,64,48,32,16 "file_path_of_ico"
    pyinstaller --onefile --noconsole "file_path_of_script" --icon="file_path_of_ico"
    ```
## tool
### python
* `ruff` => [github](https://github.com/astral-sh/ruff)
* `pytest` => [github](https://github.com/pytest-dev/pytest/)
* `PySide6` => [url](https://doc.qt.io/qtforpython-6/)
* `PyInstaller` => [url](https://pyinstaller.org/en/stable/)
* `markitdown` => [github](https://github.com/microsoft/markitdown)
* `pypdf` => [github](https://github.com/py-pdf/pypdf)
* `feedparser` => [url](https://feedparser.readthedocs.io/en/latest/)
### others
* `direnv` => [github](https://github.com/direnv/direnv)
