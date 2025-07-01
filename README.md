# public_python
## WSL2(Ubuntu)
  * `uv` => [github](https://github.com/astral-sh/uv)
    * ***Initial settings***
      ```Shell
      vim ~/.zshrc
        # uv
        # https://docs.astral.sh/uv/configuration/environment/#uv_http_timeout
        export UV_HTTP_TIMEOUT=300
        # https://docs.astral.sh/uv/configuration/environment/#ssl_cert_file
        export SSL_CERT_FILE="absolute_file_path_of_ca"
      mkdir "directory_path_of_workspace"
      cd "directory_path_of_workspace"
      uv init --name "project_name"
      ```
    * ***To do when you want to create venv***
      ```Shell
      sudo apt install direnv
      vim ~/.zshrc
        # direnv
        eval "$(direnv hook zsh)"
      cd "directory_path_of_workspace"
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
    * ***To do when you want to run scripts***
      * If you want to import scripts, create `__init__.py` file in each folder from the `PYTHONPATH` to that script
      ```Shell
      vim ~/.zshrc
        # uv
        # https://docs.astral.sh/uv/configuration/environment/#pythonpath
        export PYTHONPATH=$PYTHONPATH:"directory_path_of_root"
      cd "directory_path_of_scripts"
      uv run "file_name_of_script"
      ```
    * ***To do after opening workspace***
      ```Shell
      cd "directory_path_of_workspace"
      uv self update
      uv tool upgrade --all
      uv sync
      uv lock
      ```
## Windows
  * `Python` => [url](https://www.python.org/)
    * ***Initial settings***
      * Add the folder path of Python to the system environment variable `PATH`
      ```PowerShell
      # Windows
      winget search --name python
      winget install --id Python.Python.X.X
      python --version
      ```
    * ***To do when you want to create venv***
      ```PowerShell
      winget search --name direnv
      winget install --id direnv.direnv
      direnv --version
      notepad $PROFILE
        # direnv
        Invoke-Expression "$(direnv hook pwsh)"
      mkdir "directory_path_of_workspace"
      cd "directory_path_of_workspace"
      python -m venv "venv_name"
      notepad .envrc
        # direnv
        layout python "venv_name"
      direnv allow
      ```
    * ***To do when you want to use any tools***
      ```Shell
      cd "directory_path_of_workspace"
      pip install "tool_name"
      ```
    * ***To do when you want to run scripts***
      * Add the top folder path of Python script to the system environment variable `PYTHONPATH`
      * If you want to import scripts, create `__init__.py` file in each folder from the `PYTHONPATH` to that script
      ```Shell
      cd "directory_path_of_scripts"
      python "file_name_of_script"
      ```
    * ***To do when you want to set environment variables***
      * 環境変数の設定手順（GUI）
        1. システムの詳細設定を開きます
          「スタートメニュー」→「設定」を開きます
          「システム」→「バージョン情報」→「関連リンク」→「システムの詳細設定」
        2. 「システムのプロパティ」ダイアログを開きます
          「詳細設定」タブになっていることを確認します
          下部の「環境変数(N)...」ボタンをクリックします
        3. 「環境変数」ダイアログ
          上段：ユーザー環境変数
          下段：システム環境変数（全ユーザー共通）
          変数の追加・編集・削除
            * 新規...: 新しい変数を作成します
            * 編集...: 選択中の変数を編集します
            * 削除: 変数を削除（※慎重に）
        4. Path（パス）を編集します
          「Path」を選択して「編集」をクリック
          「環境変数名の編集」ウィンドウが開きます
          新規 をクリックして、パスを1行ずつ追加します
          OK を押してすべて閉じます
        5. 再起動・再ログインについて
          一部のアプリやターミナル（例：VS Code, PowerShellなど）は、再起動しないと新しい環境変数を反映しない場合があります。

