import sys
import subprocess

if __name__ == '__main__':
    # 收集所有传入的文件路径
    files = sys.argv[1:]
    
    # 使用subprocess启动主程序，传递所有文件路径作为参数
    subprocess.run([sys.executable, "WinHeicToPng.py"] + files)