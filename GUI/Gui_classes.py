import os.path
import tkinter
import tkinter as tk

from tkinter import ttk
from PIL import Image, ImageTk, ImageEnhance


class Window(tk.Tk):

    def __init__(self, title_text: str, geometry: str = "800x400"):

        super().__init__()
        self.title(title_text)
        self.geometry(geometry)
        self.resizable(False, False)
        self.config(bg="#f0f0f0")

    def start(self):

        self.mainloop()


class ChildWindow(tk.Toplevel):

    def __init__(self, title: str, geometry: str = "200x100", parent=None):
        super().__init__(parent)

        self.geometry(geometry)
        #self.resizable(False, False)
        self.config(bg="#f0f0f0")
        self.title = title

    def back_to_main_window(self, parent):
        self.destroy()
        parent.deiconify()


class Button(tk.Button):

    def __init__(self, parent, command, path_img: str = "GUI/button_default.png", anchor=tk.CENTER, pady: int = 0, padx: int = 0):

        super().__init__(parent)

        self.path_img = path_img
        self.path_img_hover = None

        self.img = Image.open(path_img)
        self.image = ImageTk.PhotoImage(self.img)

        self.create_hover_image()

        self.config(relief=tkinter.FLAT, command=command, image=self.image)
        self.pack(anchor=anchor, pady=pady)

        self.bind("<Enter>", self.hover_enter)
        self.bind("<Leave>", self.hover_leave)

    def hover_enter(self, event):
        self.img = Image.open(self.path_img_hover)
        self.image = ImageTk.PhotoImage(self.img)
        self.config(image=self.image)

    def hover_leave(self, event):
        self.img = Image.open(self.path_img)
        self.image = ImageTk.PhotoImage(self.img)
        self.config(image=self.image)

    def create_hover_image(self):
        enhancer = ImageEnhance.Color(self.img)
        img_hover_output = enhancer.enhance(1.4)

        img_hover_output.save(f"img/button_img_hover/{os.path.basename(self.path_img)}")

        self.path_img_hover = f"img/button_img_hover/{os.path.basename(self.path_img)}"


class RadioButton(ttk.Radiobutton):
    def __init__(self, parent, text, value, value_variable, pady: int = 0, padx: int = 0):
        super().__init__(parent, value=value, text=text, variable=value_variable)
        self.pack(anchor=tk.W, pady=pady, padx=padx)


class Frame(tk.Frame):
    def __init__(self, parent, pad_x: int = 0, pad_y: int = 0, col:int = 0, row: int = 0, sticky=tk.NE, columnspan: int = 1):
        super().__init__(parent)
        self.config(padx=pad_x, pady=pad_y)
        self.grid(row=row, column=col, sticky=sticky, columnspan=columnspan)


class Label(tk.Label):

    def __init__(self, parent, path_img: str = None, text: str = "Default", pad_y: int = 0, pad_x: int = 0, anchor=tk.E):
        super().__init__(parent)

        if path_img:
            self.img = Image.open(path_img)
            self.image = ImageTk.PhotoImage(self.img)
            self.config(image=self.image)
        else:
            self.config(text=text, font=("Arial", 14))

        self.pack(pady=pad_y, padx=pad_x, anchor=anchor)


class Entry(tk.Entry):

    def __init__(self, parent, text: str = ""):
        super().__init__(parent)
        self.config(font=("Century", 14))
        self.insert(0, text)
        self.pack(pady=6)

    def clear_entry(self):
        self.delete(0, tk.END)


class Text(tk.Text):
    def __init__(self, parent, x: int, y: int, text: str = ""):
        super().__init__(parent)
        self.insert(tk.END, text)
        self.config(state="disabled")
        self.place(x=x, y=y)


class ProgressBar(ttk.Progressbar):
    def __init__(self, parent):
        self.parent = parent
        super().__init__(self.parent, value=0, mode="determinate", orient="horizontal", length=300)
        self.pack()
