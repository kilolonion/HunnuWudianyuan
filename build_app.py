import os
import sys
import subprocess

def main():
    """
    使用Nuitka打包Streamlit应用为可执行文件
    """
    # 确保安装nuitka
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "nuitka"], check=True)
        print("✅ Nuitka已安装")
    except Exception as e:
        print(f"❌ 安装Nuitka失败: {str(e)}")
        return
    
    # 创建启动器脚本
    launcher_path = "WordFormatter_launcher.py"
    with open(launcher_path, "w", encoding="utf-8") as f:
        f.write("""
import os
import sys
import subprocess
import threading
import webbrowser
import time

def open_browser():
    # 等待服务器启动
    time.sleep(2)
    # 打开浏览器
    webbrowser.open('http://localhost:8501')

def main():
    # 设置环境变量，使Streamlit不显示菜单等
    os.environ['STREAMLIT_SERVER_HEADLESS'] = 'false'
    os.environ['STREAMLIT_SERVER_ADDRESS'] = 'localhost'
    os.environ['STREAMLIT_SERVER_PORT'] = '8501'
    os.environ['STREAMLIT_SERVER_ENABLE_TELEMETRY'] = 'false'
    os.environ['STREAMLIT_THEME_PRIMARYCOLOR'] = '#2E86C1'
    os.environ['STREAMLIT_BROWSER_GATHER_USAGE_STATS'] = 'false'
    
    # 启动浏览器线程
    threading.Thread(target=open_browser, daemon=True).start()
    
    # 导入并运行主应用
    import WordFormatter_GUI
    
if __name__ == "__main__":
    main()
""")
    print(f"✅ 创建启动器脚本: {launcher_path}")
    
    # 构建nuitka命令
    nuitka_cmd = [
        sys.executable, "-m", "nuitka",
        "--standalone",
        "--follow-imports",
        "--plugin-enable=numpy",
        "--plugin-enable=pylint-warnings",
        "--plugin-enable=tk-inter",
        "--include-package=streamlit",
        "--include-package=docx",
        "--include-package=openai",
        "--include-package=pandas",
        "--include-package=PIL",
        "--windows-disable-console",
        "--windows-icon-from-ico=favicon.ico" if os.path.exists("favicon.ico") else "",
        "--output-dir=dist",
        launcher_path
    ]
    
    # 移除空字符串
    nuitka_cmd = [item for item in nuitka_cmd if item]
    
    print("开始打包应用...")
    print(f"执行命令: {' '.join(nuitka_cmd)}")
    
    try:
        # 运行nuitka命令
        subprocess.run(nuitka_cmd, check=True)
        print("\n✅ 应用打包成功！可执行文件位于dist目录")
    except Exception as e:
        print(f"\n❌ 打包失败: {str(e)}")
        return

if __name__ == "__main__":
    main() 