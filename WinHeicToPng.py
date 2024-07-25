import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image
import pillow_heif
import winreg
import threading
import queue
import ctypes

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def convert_heic_to_png(heic_path):
    try:
        heif_file = pillow_heif.read_heif(heic_path)
        image = Image.frombytes(
            heif_file.mode, 
            heif_file.size, 
            heif_file.data,
            "raw",
            heif_file.mode,
            heif_file.stride,
        )
        png_path = os.path.splitext(heic_path)[0] + '.png'
        image.save(png_path, 'PNG')
        return f"Converted {heic_path} to {png_path}"
    except Exception as e:
        return f"Error converting {heic_path}: {str(e)}"

class WinHeicToPngGUI:
    def __init__(self, master, initial_files=[]):
        self.master = master
        master.title("WinHeicToPng")
        master.geometry("400x350")

        self.label = tk.Label(master, text="点击添加文件按钮选择HEIC文件")
        self.label.pack(pady=10)

        self.file_listbox = tk.Listbox(master, width=50, height=10)
        self.file_listbox.pack(pady=10)

        button_frame = tk.Frame(master)
        button_frame.pack(fill=tk.X, padx=10, pady=5)

        self.add_button = tk.Button(button_frame, text="添加文件", command=self.add_files)
        self.add_button.pack(side=tk.LEFT, padx=5)

        self.convert_button = tk.Button(button_frame, text="转换", command=self.start_conversion)
        self.convert_button.pack(side=tk.LEFT, padx=5)

        self.register_button = tk.Button(button_frame, text="注册", command=self.register_context_menu)
        self.register_button.pack(side=tk.RIGHT, padx=5)

        self.progress_var = tk.StringVar()
        self.progress_label = tk.Label(master, textvariable=self.progress_var)
        self.progress_label.pack(pady=10)

        # 用于线程间通信的队列
        self.queue = queue.Queue()

        # 添加初始文件
        for file in initial_files:
            if file.lower().endswith('.heic'):
                self.file_listbox.insert(tk.END, file)

    def add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("HEIC files", "*.heic")])
        for file in files:
            self.file_listbox.insert(tk.END, file)

    def start_conversion(self):
        files = self.file_listbox.get(0, tk.END)
        if not files:
            messagebox.showinfo("提示", "请先添加HEIC文件")
            return
        
        self.convert_button.config(state=tk.DISABLED)
        self.progress_var.set("转换中...")
        
        threading.Thread(target=self.convert_files, args=(files,), daemon=True).start()
        self.master.after(100, self.check_queue)

    def convert_files(self, files):
        for file in files:
            result = convert_heic_to_png(file)
            self.queue.put(result)
        self.queue.put("DONE")

    def check_queue(self):
        try:
            message = self.queue.get_nowait()
            if message == "DONE":
                self.file_listbox.delete(0, tk.END)
                self.progress_var.set("转换完成")
                self.convert_button.config(state=tk.NORMAL)
                messagebox.showinfo("转换完成", "所有文件已成功转换")
            else:
                self.progress_var.set(message)
                self.master.after(100, self.check_queue)
        except queue.Empty:
            self.master.after(100, self.check_queue)

    def register_context_menu(self):
        if not is_admin():
            messagebox.showerror("错误", "注册右键菜单需要管理员权限。请以管理员身份运行此程序。")
            return

        try:
            key_path = r'*\shell\HeicToPng'
            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, key_path)
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path, 0, winreg.KEY_ALL_ACCESS)
            winreg.SetValueEx(key, '', 0, winreg.REG_SZ, 'Convert to PNG')
            winreg.SetValueEx(key, 'Icon', 0, winreg.REG_SZ, sys.executable)
            command_key = winreg.CreateKey(key, 'command')
            winreg.SetValueEx(command_key, '', 0, winreg.REG_SZ, f'"{sys.executable}" "{__file__}" "%V"')
            winreg.CloseKey(key)
            messagebox.showinfo("成功", "右键菜单已成功注册")
        except Exception as e:
            messagebox.showerror("错误", f"注册右键菜单时出错：{str(e)}")

def main():
    if len(sys.argv) > 1:
        # 如果有命令行参数，解析所有文件路径
        files = []
        for arg in sys.argv[1:]:
            # 移除引号
            arg = arg.strip('"')
            if os.path.isfile(arg) and arg.lower().endswith('.heic'):
                files.append(arg)
    else:
        files = []

    root = tk.Tk()
    gui = WinHeicToPngGUI(root, files)
    root.mainloop()

if __name__ == '__main__':
    main()