import tkinter as tk
from tkinterdnd2 import DND_FILES, TkinterDnD
import os
from docx2pdf import convert
from tkinter import messagebox

class DragDropWidget(tk.Frame):
    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self.create_widgets()

    def create_widgets(self):
        self.label = tk.Label(self, text='將Word文件拖放到此處', font=('Arial', 12), width=30, height=10)
        self.label.pack(pady=50)

        self.label.drop_target_register(DND_FILES)
        self.label.dnd_bind('<<Drop>>', self.on_drop)

    def on_drop(self, event):
        file_path = event.data
        if file_path:
            output_dir = os.path.dirname(file_path)
            convert_word_to_pdf(file_path, output_dir)
            messagebox.showinfo('轉換完成', 'Word 轉換為 PDF 文檔完成！')

def convert_word_to_pdf(word_path, output_dir):
    output_filename = os.path.splitext(os.path.basename(word_path))[0] + '.pdf'
    output_path = os.path.join(output_dir, output_filename)
    convert(word_path, output_path)

root = TkinterDnD.Tk()
root.title('Word 轉 PDF 工具')

app = DragDropWidget(root)
app.pack()

root.mainloop()