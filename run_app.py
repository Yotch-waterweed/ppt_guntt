import streamlit.web.cli as stcli
import os, sys

def resolve_path(path):
    resolved_path = os.path.abspath(os.path.join(os.getcwd(), path))
    return resolved_path

if __name__ == "__main__":
    # app.pyのパスを取得
    script_path = resolve_path("app.py")
    
    # 実行環境がPyInstallerでパッケージ化されているかチェック
    if getattr(sys, 'frozen', False):
        # exeから実行されている場合
        sys.argv = ["streamlit", "run", script_path, "--global.developmentMode=false"]
        sys.exit(stcli.main())
    else:
        # 通常のPythonから実行されている場合
        sys.argv = ["streamlit", "run", script_path]
        sys.exit(stcli.main())
