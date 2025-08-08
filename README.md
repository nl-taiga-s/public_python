# public_python
## WSL2(Ubuntu)
  * `uv` => [github](https://github.com/astral-sh/uv)
    * ***Initial settings***
      ```Shell
      mkdir "directory_path_of_workspace"
      cd "directory_path_of_workspace"
      uv init --name "project_name"
      vim ~/.zshrc
        # uv
        # https://docs.astral.sh/uv/configuration/environment/#uv_http_timeout
        export UV_HTTP_TIMEOUT=300
        # https://docs.astral.sh/uv/configuration/environment/#ssl_cert_file
        export SSL_CERT_FILE="file_path_of_cert"
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
        source .venv/bin/activate
        # https://docs.astral.sh/uv/configuration/environment/#pythonpath
        export PYTHONPATH=$PWD
      direnv allow
      ```
    * ***To do after opening workspace***
      ```Shell
      cd "directory_path_of_workspace"
      uv self update
      uv tool upgrade --all
      uv sync
      uv lock
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
      ```Shell
      cd "directory_path_of_workspace"
      uv run "relative_file_path_of_script"
      ```
      * Please place `__init__.py` in each folder leading from one level down in the workspace to the folder containing the script.
    * ***To do when you want to test scripts***
      ```Shell
      cd "directory_path_of_workspace"
      pytest
      ```
## Windows
  * `Python` => [url](https://www.python.org/)
    * ***Initial settings***
      ```PowerShell
      mkdir "directory_path_of_workspace"
      notepad $profile
        $env:PYTHONPATH = "directory_path_of_workspace"
      ```
    * ***To do when you want to create venv***
      ```PowerShell
      cd "directory_path_of_workspace"
      python -m venv "venv_name"
      ```
    * ***To do after opening workspace***
      ```PowerShell
      cd "directory_path_of_workspace"
      "venv_name"\Scripts\Activate.ps1
      python -m pip install --upgrade pip
      ```
    * ***To do when you want to use any tools***
      ```PowerShell
      cd "directory_path_of_workspace"
      pip install "tool_name"
      ```
    * ***To do when you want to run scripts***
      ```PowerShell
      cd "directory_path_of_workspace"
      python "relative_file_path_of_script"
      ```
      * Please place `__init__.py` in each folder leading from one level down in the workspace to the folder containing the script.
    * ***To do before closing workspace***
      ```PowerShell
      deactivate
      ```
