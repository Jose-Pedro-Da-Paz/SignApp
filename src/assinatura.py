from tkinter import *
from tkinter import filedialog, filedialog, messagebox
from PIL import Image, ImageDraw, ImageTk
import os

class SignatureApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Capturar Assinatura")

        self.canvas_width = 400
        self.canvas_height = 200
        self.canvas = Canvas(self.master, width=self.canvas_width, height=self.canvas_height, bg='white')
        self.canvas.pack()

        # Variáveis para armazenar a imagem e a interface de desenho
        self.image = Image.new("RGB", (self.canvas_width, self.canvas_height), (255, 255, 255))
        self.draw = ImageDraw.Draw(self.image)

        self.last_x, self.last_y = None, None

        self.canvas.bind("<B1-Motion>", self.paint)
        self.canvas.bind("<ButtonRelease-1>", self.reset)
        
        # Botões para salvar e limpar
        Button(self.master, text="Salvar", command=self.save_signature).pack(side=LEFT, padx=10, pady=10)
        Button(self.master, text="Limpar", command=self.clear_canvas).pack(side=LEFT, padx=10, pady=10)
        
    def paint(self, event):
        if self.last_x and self.last_y:
            self.canvas.create_line((self.last_x, self.last_y, event.x, event.y), width=2, fill='black', capstyle=ROUND, smooth=True)
            self.draw.line((self.last_x, self.last_y, event.x, event.y), fill="black", width=2)
        self.last_x = event.x
        self.last_y = event.y

    def reset(self, event):
        self.last_x, self.last_y = None, None

    def save_signature(self):
        # Caminho para salvar a imagem da assinatura
        save_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png")])
        if save_path:
            self.image.save(save_path)
            messagebox.showinfo("Sucesso", "Assinatura salva com sucesso!")

    def clear_canvas(self):
        self.canvas.delete("all")
        self.image = Image.new("RGB", (self.canvas_width, self.canvas_height), (255, 255, 255))
        self.draw = ImageDraw.Draw(self.image)

if __name__ == "__main__":
    root = Tk()
    app = SignatureApp(root)
    root.mainloop()
