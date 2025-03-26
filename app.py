import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from PyPDF2 import PdfMerger
from PIL import Image, ImageTk

class PDFConcatenatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Concatenador de documentos")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        self.style = ttk.Style()
        self.style.configure("TButton", padding=6, relief="flat", background="#3498db")
        self.style.configure("TFrame", background="#f0f0f0")
        
        # Lista para armazenar caminhos dos PDFs
        self.pdf_files = []
        
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        title_label = ttk.Label(main_frame, text="Concatenador de documentos", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # Botões
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=10)
        
        select_button = ttk.Button(buttons_frame, text="Selecionar PDFs", command=self.select_pdfs)
        select_button.pack(side=tk.LEFT, padx=5)
        
        concat_button = ttk.Button(buttons_frame, text="Concatenar PDFs", command=self.concatenar_pdfs)
        concat_button.pack(side=tk.LEFT, padx=5)
        
        clear_button = ttk.Button(buttons_frame, text="Limpar Lista", command=self.clear_list)
        clear_button.pack(side=tk.LEFT, padx=5)
        
        list_frame = ttk.LabelFrame(main_frame, text="Arquivos PDF Selecionados")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.files_listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED, 
                                       font=("Arial", 10), 
                                       yscrollcommand=scrollbar.set)
        self.files_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        scrollbar.config(command=self.files_listbox.yview)
    
        order_frame = ttk.Frame(main_frame)
        order_frame.pack(fill=tk.X, pady=5)
        
        move_up_button = ttk.Button(order_frame, text="↑ Mover para Cima", command=self.move_up)
        move_up_button.pack(side=tk.LEFT, padx=5)
        
        move_down_button = ttk.Button(order_frame, text="↓ Mover para Baixo", command=self.move_down)
        move_down_button.pack(side=tk.LEFT, padx=5)
        
        remove_button = ttk.Button(order_frame, text="Remover Selecionado", command=self.remove_selected)
        remove_button.pack(side=tk.LEFT, padx=5)
    
        self.status_var = tk.StringVar()
        self.status_var.set("Pronto para selecionar arquivos PDF")
        status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def select_pdfs(self):
        """Abre diálogo para selecionar múltiplos arquivos PDF"""
        files = filedialog.askopenfilenames(
            title="Selecione os arquivos PDF",
            filetypes=[("Arquivos PDF", "*.pdf")],
            multiple=True
        )
        
        if files:
            for file in files:
                if file not in self.pdf_files:
                    self.pdf_files.append(file)
                    self.files_listbox.insert(tk.END, os.path.basename(file))
            
            self.status_var.set(f"{len(files)} arquivos adicionados à lista")
    
    def concatenar_pdfs(self):
        """Concatena os PDFs selecionados em um único arquivo"""
        if not self.pdf_files:
            messagebox.showwarning("Aviso", "Nenhum arquivo PDF selecionado!")
            return
        
        output_file = filedialog.asksaveasfilename(
            title="Salvar PDF concatenado como",
            defaultextension=".pdf",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        
        if not output_file:
            return
        
        try:
            merger = PdfMerger()
    
            progress_window = tk.Toplevel(self.root)
            progress_window.title("Progresso")
            progress_window.geometry("300x100")
            progress_window.transient(self.root)
            
            progress_label = ttk.Label(progress_window, text="Concatenando PDFs...")
            progress_label.pack(pady=10)
            
            progress_bar = ttk.Progressbar(progress_window, orient="horizontal", 
                                           length=250, mode="determinate", maximum=len(self.pdf_files))
            progress_bar.pack(pady=10)
            
            progress_window.update()
            
            # Concatenar PDFs na ordem da listbox
            for i in range(len(self.pdf_files)):
                file_path = self.pdf_files[i]
                merger.append(file_path)
                progress_bar["value"] = i + 1
                progress_window.update()
            
            # Salvar o arquivo resultante
            merger.write(output_file)
            merger.close()
            
            progress_window.destroy()
            
            self.status_var.set(f"PDF concatenado salvo com sucesso em: {output_file}")
            messagebox.showinfo("Sucesso", f"PDF concatenado com sucesso!\nSalvo em: {output_file}")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro durante a concatenação:\n{str(e)}")
            self.status_var.set("Erro ao concatenar PDFs")
    
    def clear_list(self):
        """Limpa a lista de arquivos selecionados"""
        self.pdf_files = []
        self.files_listbox.delete(0, tk.END)
        self.status_var.set("Lista de arquivos limpa")
    
    def move_up(self):
        """Move o item selecionado para cima na lista"""
        selected_indices = self.files_listbox.curselection()
        if not selected_indices or selected_indices[0] == 0:
            return
        
        for index in selected_indices:
            if index > 0:
        
                self.pdf_files[index], self.pdf_files[index-1] = self.pdf_files[index-1], self.pdf_files[index]
                
                file_name = self.files_listbox.get(index)
                self.files_listbox.delete(index)
                self.files_listbox.insert(index-1, file_name)
                self.files_listbox.selection_clear(0, tk.END)
                self.files_listbox.selection_set(index-1)
    
    def move_down(self):
        """Move o item selecionado para baixo na lista"""
        selected_indices = self.files_listbox.curselection()
        if not selected_indices or selected_indices[-1] == self.files_listbox.size() - 1:
            return
        
        for index in sorted(selected_indices, reverse=True):
            if index < self.files_listbox.size() - 1:
                
                self.pdf_files[index], self.pdf_files[index+1] = self.pdf_files[index+1], self.pdf_files[index]
                
                file_name = self.files_listbox.get(index)
                self.files_listbox.delete(index)
                self.files_listbox.insert(index+1, file_name)
                self.files_listbox.selection_clear(0, tk.END)
                self.files_listbox.selection_set(index+1)
    
    def remove_selected(self):
        """Remove os itens selecionados da lista"""
        selected_indices = self.files_listbox.curselection()
        if not selected_indices:
            return
        
        # Remover da lista em ordem reversa para evitar problemas com índices
        for index in sorted(selected_indices, reverse=True):
            del self.pdf_files[index]
            self.files_listbox.delete(index)
        
        self.status_var.set(f"{len(selected_indices)} arquivo(s) removido(s) da lista")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConcatenatorApp(root)
    root.mainloop()