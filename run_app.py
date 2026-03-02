import streamlit.web.cli as stcli
import os, sys

def get_base_path():
    """PyInstallerでパッケージ化されている場合は展開先、通常時はスクリプトのディレクトリを返す"""
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

if __name__ == "__main__":
    base = get_base_path()
    
    # app.pyのパスを取得（PyInstaller展開先 or スクリプトと同じディレクトリ）
    script_path = os.path.join(base, "app.py")
    
    # カレントディレクトリを展開先に変更（データファイル参照のため）
    os.chdir(base)
    
    if getattr(sys, 'frozen', False):
        sys.argv = ["streamlit", "run", script_path,
                     "--global.developmentMode=false",
                     "--server.headless=true",
                     "--browser.gatherUsageStats=false"]
    else:
        sys.argv = ["streamlit", "run", script_path]
    
    sys.exit(stcli.main())
