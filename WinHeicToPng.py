import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image
import winreg

def convert_heic_to_png(heic_path):
    try:
        image = Image.open(heic_path)
        png_path = os.path.splitext(heic_path)[0] + '.png'
        image.save(png_path, 'PNG')
        print(f"Converted {heic_path} to {png_path}")
    except Exception as e:
        print(f"Error converting {heic_path}: {str(e)}")

class WinHeicToPngGUI:
    def __init__(self, master):
        self.master = master
        master.title("WinHeicToPng")
        master.geometry("400x300")

        self.label = tk.Label(master, text="点击添加文件按钮选择HEIC文件")
        self.label.pack(pady=10)

        self.file_listbox = tk.Listbox(master, width=50, height=10)
        self.file_listbox.pack(pady=10)

        self.add_button = tk.Button(master, text="添加文件", command=self.add_files)
        self.add_button.pack(side=tk.LEFT, padx=10)

        self.convert_button = tk.Button(master, text="转换", command=self.convert_files)
        self.convert_button.pack(side=tk.RIGHT, padx=10)

    def add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("HEIC files", "*.heic")])
        for file in files:
            self.file_listbox.insert(tk.END, file)

    def convert_files(self):
        files = self.file_listbox.get(0, tk.END)
        for file in files:
            convert_heic_to_png(file)
        self.file_listbox.delete(0, tk.END)
        messagebox.showinfo("转换完成", "所有文件已成功转换")

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
        convert_heic_to_png(sys.argv[1])
    else:
        # Run GUI
        root = tk.Tk()
        gui = WinHeicToPngGUI(root)
        root.mainloop()

    # Add context menu on first run
    if not os.path.exists(os.path.join(os.getenv('APPDATA'), 'WinHeicToPng', 'installed')):
        add_context_menu()
        os.makedirs(os.path.join(os.getenv('APPDATA'), 'WinHeicToPng'), exist_ok=True)
        with open(os.path.join(os.getenv('APPDATA'), 'WinHeicToPng', 'installed'), 'w') as f:
            f.write('installed')