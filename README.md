
'''bash
pyinstaller --onefile --windowed --add-data "C:\Path\To\Your\Python\Lib\site-packages\tkinterdnd2;tkinterdnd2" --icon=path\to\your\icon.ico WinHeicToPng.py
'''

'''
pyinstaller --onefile --windowed --add-data "C:\Users\gloridust\AppData\Local\Programs\Python\Python312\Lib\site-packages\tkinterdnd2\__init__.py;tkinterdnd2" --add-binary "C:\Users\gloridust\AppData\Local\Programs\Python\Python312\Lib\site-packages\tkinterdnd2\__init__.py;tkdnd" --hidden-import tkinterdnd2 WinHeicToPng.py
'''