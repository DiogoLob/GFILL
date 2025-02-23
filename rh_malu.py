import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
from fpdf import FPDF

class ResumeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Currículo")
        self.root.geometry("600x700")
        
        self.logo_path = None
        self.candidate_name = tk.StringVar()
        
        self.create_widgets()

    def create_widgets(self):
        self.logo_btn = tk.Button(self.root, text="Adicionar Logo", command=self.add_logo)
        self.logo_btn.pack()
        
        self.logo_label = tk.Label(self.root)
        self.logo_label.pack()
        
        tk.Label(self.root, text="Nome do Candidato:").pack()
        self.name_entry = tk.Entry(self.root, textvariable=self.candidate_name, width=60)
        self.name_entry.pack()
        
        self.fields_frame = tk.Frame(self.root)
        self.fields_frame.pack(pady=10)
        
        self.entries = []
        self.add_field()
        
        self.add_field_btn = tk.Button(self.root, text="Adicionar Campo", command=self.add_field)
        self.add_field_btn.pack()
        
        self.generate_btn = tk.Button(self.root, text="Gerar Currículo", command=self.generate_resume)
        self.generate_btn.pack()
    
    def add_logo(self):
        file_path = filedialog.askopenfilename(filetypes=[("Imagens", "*.png;*.jpg;*.jpeg")])
        if file_path:
            self.logo_path = file_path
            img = Image.open(file_path)
            img = img.resize((600, 150))
            self.logo_img = ImageTk.PhotoImage(img)
            self.logo_label.config(image=self.logo_img)
    
    def add_field(self):
        frame = tk.Frame(self.fields_frame)
        frame.pack(fill='x', pady=3)
        
        tk.Label(frame, text="Título:").pack(side="left")
        title_entry = tk.Entry(frame, width=20)
        title_entry.pack(side="left", padx=3)
        
        tk.Label(frame, text="Conteúdo:").pack(side="left")
        content_entry = tk.Text(frame, width=40, height=3)
        content_entry.pack(side="left", padx=3)
        
        self.entries.append((title_entry, content_entry))
    
    def generate_resume(self):
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        
        if self.logo_path:
            pdf.image(self.logo_path, x=0, y=0, w=210)
            pdf.ln(38)
        
        pdf.set_font("Arial", style='B', size=25)
        pdf.cell(200, 10, self.candidate_name.get().encode('latin-1', 'ignore').decode('latin-1'), ln=True, align='C')
        pdf.ln(10)
        
        for title_entry, content_entry in self.entries:
            title = title_entry.get().strip()
            content = content_entry.get("1.0", tk.END).strip()
            if title and content:
                pdf.set_font("Arial", style='B', size=12)
                pdf.cell(200, 8, title.encode('latin-1', 'ignore').decode('latin-1'), ln=True)
                pdf.ln(1)  # Adiciona um pequeno espaço entre título e descrição
                pdf.set_font("Arial", size=12)
                pdf.multi_cell(0, 6, content.encode('latin-1', 'ignore').decode('latin-1'), align='J')
                pdf.ln(4)  # Espaçamento reduzido entre seções
        
        pdf.output("curriculo.pdf", "F")
        messagebox.showinfo("Sucesso", "Currículo gerado com sucesso! Verifique o arquivo curriculo.pdf.")
        
if __name__ == "__main__":
    root = tk.Tk()
    app = ResumeApp(root)
    root.mainloop()
