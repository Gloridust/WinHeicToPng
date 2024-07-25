import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from PIL import Image
import pillow_heif
import winreg
import threading
import queue

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
    def __init__(self, master):
        self.master = master
        master.title("WinHeicToPng")
        master.geometry("400x300")

        self.label = tk.Label(master, text="拖放 HEIC 文件到这里或点击添加文件按钮")
        self.label.pack(pady=10)

        self.file_listbox = tk.Listbox(master, width=50, height=10)
        self.file_listbox.pack(pady=10)

        self.add_button = tk.Button(master, text="添加文件", command=self.add_files)
        self.add_button.pack(side=tk.LEFT, padx=10)

        self.convert_button = tk.Button(master, text="转换", command=self.start_conversion)
        self.convert_button.pack(side=tk.RIGHT, padx=10)

        self.progress_var = tk.StringVar()
        self.progress_label = tk.Label(master, textvariable=self.progress_var)
        self.progress_label.pack(pady=10)

        # 设置拖放
        self.file_listbox.drop_target_register(DND_FILES)
        self.file_listbox.dnd_bind('<<Drop>>', self.drop)

        # 用于线程间通信的队列
        self.queue = queue.Queue()

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

    def drop(self, event):
        files = self.file_listbox.tk.splitlist(event.data)
        for file in files:
            if file.lower().endswith('.heic'):
                self.file_listbox.insert(tk.END, file)

def add_context_menu():
    try:
        key_path = r'*\shell\HeicToPng'
        winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, key_path)
        key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path, 0, winreg.KEY_ALL_ACCESS)
        winreg.SetValueEx(key, '', 0, winreg.REG_SZ, 'Convert to PNG')
        winreg.SetValueEx(key, 'Icon', 0, winreg.REG_SZ, sys.executable)
        command_key = winreg.CreateKey(key, 'command')
        winreg.SetValueEx(command_key, '', 0, winreg.REG_SZ, f'"{sys.executable}" "{__file__}" "%1"')
        winreg.CloseKey(key)
        print("Context menu added successfully")
    except Exception as e:
        print(f"Error adding context menu: {str(e)}")

if __name__ == '__main__':
    if len(sys.argv) > 1:
        # Convert file from context menu
        print(convert_heic_to_png(sys.argv[1]))
    else:
        # Run GUI
        root = TkinterDnD.Tk()
        gui = WinHeicToPngGUI(root)
        root.mainloop()

    # Add context menu on first run
    if not os.path.exists(os.path.join(os.getenv('APPDATA'), 'WinHeicToPng', 'installed')):
        add_context_menu()
        os.makedirs(os.path.join(os.getenv('APPDATA'), 'WinHeicToPng'), exist_ok=True)
        with open(os.path.join(os.getenv('APPDATA'), 'WinHeicToPng', 'installed'), 'w') as f:
            f.write('installed')