import os
import platform


def is_wsl(self) -> bool:
    """WSL環境かどうか判定します"""
    print(is_wsl.__doc__)
    if platform.system() != "Linux":
        return False
    try:
        with open("/proc/version", "r") as f:
            content = f.read().lower()
            return "microsoft" in content or "wsl" in content
    except Exception:
        return False


class FormatPath:
    """パスを整形します"""

    def __init__(self):
        """初期化します"""
        print(self.__class__.__doc__)

    def to_path_seaparator_for_os(self, target_path: str) -> str:
        """
        OSに応じて、パスの区切り文字を統一します
        * Windows: "/" => "\\"
        * WSL: "\\" => "/"
        """
        system_name = platform.system()
        try:
            if is_wsl():
                # WSL
                return os.path.normpath(target_path)
            elif system_name == "Windows":
                # Windows
                return target_path.replace("\\", "/")
            else:
                return target_path
        except Exception as e:
            print(e)

    def if_unc_path(self, target: str) -> str:
        """
        UNC(Universal Naming Convention)パスの条件分岐をします
        ex. \\\\ZZ.ZZZ.ZZZ.Z
        """
        if target.startswith(r"\\"):
            return target
        return os.path.abspath(target)
