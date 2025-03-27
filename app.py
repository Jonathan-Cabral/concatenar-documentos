import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import tempfile
import shutil
from PyPDF2 import PdfMerger
from PIL import Image, ImageTk
from docx2pdf import convert
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from docx import Document
import threading

class DocumentConcatenatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Concatenador de Documentos")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # Definindo estilo
        self.style = ttk.Style()
        self.style.configure("TButton", padding=6, relief="flat", background="#3498db")
        self.style.configure("TFrame", background="#f0f0f0")
        
        # Lista para armazenar caminhos dos documentos e seus tipos
        self.document_files = []  # Cada entrada será um tuple (caminho, tipo)
        self.temp_dir = tempfile.mkdtemp()  # Diretório temporário para conversões
        
        # Criando o frame principal
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        title_label = ttk.Label(main_frame, text="Concatenador de Documentos", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # Botões
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=10)
        
        select_button = ttk.Button(buttons_frame, text="Selecionar Documentos", command=self.select_documents)
        select_button.pack(side=tk.LEFT, padx=5)
        
        concat_button = ttk.Button(buttons_frame, text="Concatenar Documentos", command=self.concatenate_documents)
        concat_button.pack(side=tk.LEFT, padx=5)
        
        clear_button = ttk.Button(buttons_frame, text="Limpar Lista", command=self.clear_list)
        clear_button.pack(side=tk.LEFT, padx=5)
        
        # Lista de arquivos selecionados
        list_frame = ttk.LabelFrame(main_frame, text="Documentos Selecionados")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Scrollbar para a lista
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.files_listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED, 
                                       font=("Arial", 10), 
                                       yscrollcommand=scrollbar.set)
        self.files_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        scrollbar.config(command=self.files_listbox.yview)
        
        # Frame para botões de reordenação
        order_frame = ttk.Frame(main_frame)
        order_frame.pack(fill=tk.X, pady=5)
        
        move_up_button = ttk.Button(order_frame, text="↑ Mover para Cima", command=self.move_up)
        move_up_button.pack(side=tk.LEFT, padx=5)
        
        move_down_button = ttk.Button(order_frame, text="↓ Mover para Baixo", command=self.move_down)
        move_down_button.pack(side=tk.LEFT, padx=5)
        
        remove_button = ttk.Button(order_frame, text="Remover Selecionado", command=self.remove_selected)
        remove_button.pack(side=tk.LEFT, padx=5)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Pronto para selecionar documentos")
        status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Bind para fechamento da janela
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def select_documents(self):
        """Abre diálogo para selecionar múltiplos arquivos (PDF, DOCX, TXT)"""
        files = filedialog.askopenfilenames(
            title="Selecione os documentos",
            filetypes=[
                ("Todos os documentos suportados", "*.pdf *.docx *.txt"),
                ("Arquivos PDF", "*.pdf"),
                ("Documentos Word", "*.docx"),
                ("Arquivos de Texto", "*.txt")
            ],
            multiple=True
        )
        
        if files:
            added_count = 0
            for file in files:
                file_extension = os.path.splitext(file)[1].lower()
                
                # Determina o tipo do documento
                if file_extension == '.pdf':
                    doc_type = 'pdf'
                elif file_extension == '.docx':
                    doc_type = 'docx'
                elif file_extension == '.txt':
                    doc_type = 'txt'
                else:
                    continue  # Ignora tipos não suportados
                
                # Adiciona à lista se não estiver já
                file_info = (file, doc_type)
                if file_info not in self.document_files:
                    self.document_files.append(file_info)
                    filename = os.path.basename(file)
                    # Adiciona um ícone de tipo na listbox
                    self.files_listbox.insert(tk.END, f"[{doc_type.upper()}] {filename}")
                    added_count += 1
            
            self.status_var.set(f"{added_count} arquivos adicionados à lista")
    
    def convert_docx_to_pdf(self, docx_path):
        """Converte um arquivo DOCX para PDF"""
        try:
            filename = os.path.basename(docx_path)
            base_name = os.path.splitext(filename)[0]
            output_pdf = os.path.join(self.temp_dir, f"{base_name}.pdf")
            
            # Usando docx2pdf para converter
            convert(docx_path, output_pdf)
            return output_pdf
        except Exception as e:
            messagebox.showerror("Erro de Conversão", f"Erro ao converter arquivo DOCX: {str(e)}")
            return None
    
    def convert_txt_to_pdf(self, txt_path):
        """Converte um arquivo TXT para PDF com suporte a múltiplas páginas"""
        try:
            filename = os.path.basename(txt_path)
            base_name = os.path.splitext(filename)[0]
            output_pdf = os.path.join(self.temp_dir, f"{base_name}.pdf")
            
            # Configurações de página
            page_width, page_height = letter
            margin = 40
            font_name = "Helvetica"
            font_size = 12
            line_height = font_size * 1.2  # Espaçamento entre linhas
            
            # Cálculo de linhas por página (altura útil / altura da linha)
            lines_per_page = int((page_height - 2 * margin) / line_height)
            
            # Leitura do conteúdo do arquivo
            with open(txt_path, 'r', encoding='utf-8', errors='ignore') as file:
                all_lines = [line.strip() for line in file.readlines()]
            
            # Criação do PDF
            c = canvas.Canvas(output_pdf, pagesize=letter)
            
            page_num = 0
            start_line = 0
            
            # Loop de criação de páginas
            while start_line < len(all_lines):
                if page_num > 0:
                    c.showPage()  # Nova página
                
                y_position = page_height - margin
                
                # Adicionar cabeçalho com nome do arquivo na primeira página
                if page_num == 0:
                    c.setFont(font_name, font_size + 2)
                    c.drawString(margin, y_position, f"Arquivo: {filename}")
                    y_position -= line_height * 2
                
                c.setFont(font_name, font_size)
                
                # Processar linhas para esta página
                end_line = min(start_line + lines_per_page, len(all_lines))
                for i in range(start_line, end_line):
                    # Quebra linhas muito longas para ajustar à largura da página
                    line = all_lines[i]
                    max_width = page_width - 2 * margin
                    
                    if c.stringWidth(line, font_name, font_size) <= max_width:
                        c.drawString(margin, y_position, line)
                        y_position -= line_height
                    else:
                        # Quebra a linha em múltiplas linhas
                        words = line.split()
                        current_line = ""
                        
                        for word in words:
                            test_line = current_line + " " + word if current_line else word
                            if c.stringWidth(test_line, font_name, font_size) <= max_width:
                                current_line = test_line
                            else:
                                c.drawString(margin, y_position, current_line)
                                y_position -= line_height
                                current_line = word
                                
                                # Se não houver mais espaço na página atual
                                if y_position < margin:
                                    c.showPage()
                                    y_position = page_height - margin
                                    c.setFont(font_name, font_size)
                        
                        # Desenha a última linha se houver
                        if current_line:
                            c.drawString(margin, y_position, current_line)
                            y_position -= line_height
                
                start_line = end_line
                page_num += 1
            
            # Salva o PDF
            c.save()
            return output_pdf
        except Exception as e:
            messagebox.showerror("Erro de Conversão", f"Erro ao converter arquivo TXT: {str(e)}")
            return None

    def update_progress(self, progress_bar, progress_window, value, max_value, status_text):
        """Atualiza a barra de progresso e texto"""
        progress_bar["value"] = value
        progress_bar["maximum"] = max_value
        progress_window.title(f"Progresso - {value}/{max_value}")
        progress_window.update()
    
    def process_documents(self, output_file, progress_window, progress_bar, progress_label):
        """Processa os documentos em uma thread separada para evitar que a UI congele"""
        try:
            # Lista para armazenar caminhos dos PDFs convertidos
            pdf_files = []
            
            # Converter todos os documentos não-PDF para PDF
            total_files = len(self.document_files)
            for i, (file_path, doc_type) in enumerate(self.document_files):
                self.update_progress(progress_bar, progress_window, i, total_files, 
                                    f"Convertendo arquivo {i+1} de {total_files}")
                progress_label.config(text=f"Convertendo: {os.path.basename(file_path)}")
                
                # Converter de acordo com o tipo
                if doc_type == 'pdf':
                    pdf_files.append(file_path)
                elif doc_type == 'docx':
                    pdf_path = self.convert_docx_to_pdf(file_path)
                    if pdf_path:
                        pdf_files.append(pdf_path)
                elif doc_type == 'txt':
                    pdf_path = self.convert_txt_to_pdf(file_path)
                    if pdf_path:
                        pdf_files.append(pdf_path)
            
            # Mesclar todos os PDFs
            if pdf_files:
                progress_label.config(text="Mesclando PDFs...")
                merger = PdfMerger()
                
                for i, pdf_file in enumerate(pdf_files):
                    self.update_progress(progress_bar, progress_window, i, len(pdf_files), 
                                        f"Mesclando PDF {i+1} de {len(pdf_files)}")
                    merger.append(pdf_file)
                
                # Salvar o arquivo final
                merger.write(output_file)
                merger.close()
                
                # Atualizar a UI no thread principal
                self.root.after(0, lambda: self.status_var.set(f"Documento concatenado salvo em: {output_file}"))
                self.root.after(0, lambda: messagebox.showinfo("Sucesso", 
                                                            f"Documentos concatenados com sucesso!\nSalvo em: {output_file}"))
            else:
                self.root.after(0, lambda: messagebox.showwarning("Aviso", "Nenhum documento válido para concatenar!"))
            
            self.root.after(0, lambda: progress_window.destroy())
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Erro", f"Ocorreu um erro durante o processamento:\n{str(e)}"))
            self.root.after(0, lambda: progress_window.destroy())
    
    def concatenate_documents(self):
        """Concatena os documentos selecionados em um único arquivo PDF"""
        if not self.document_files:
            messagebox.showwarning("Aviso", "Nenhum documento selecionado!")
            return
        
        output_file = filedialog.asksaveasfilename(
            title="Salvar documento concatenado como",
            defaultextension=".pdf",
            filetypes=[("Arquivo PDF", "*.pdf")]
        )
        
        if not output_file:
            return
        
        # Criar janela de progresso
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Progresso")
        progress_window.geometry("400x120")
        progress_window.transient(self.root)
        progress_window.resizable(False, False)
        
        progress_label = ttk.Label(progress_window, text="Preparando documentos...")
        progress_label.pack(pady=10)
        
        progress_bar = ttk.Progressbar(progress_window, orient="horizontal", 
                                      length=350, mode="determinate")
        progress_bar.pack(pady=10)
        
        # Iniciar processamento em uma thread separada
        processing_thread = threading.Thread(
            target=self.process_documents,
            args=(output_file, progress_window, progress_bar, progress_label)
        )
        processing_thread.daemon = True  # Thread finaliza quando o programa fecha
        processing_thread.start()
    
    def clear_list(self):
        """Limpa a lista de arquivos selecionados"""
        self.document_files = []
        self.files_listbox.delete(0, tk.END)
        self.status_var.set("Lista de arquivos limpa")
    
    def move_up(self):
        """Move o item selecionado para cima na lista"""
        selected_indices = self.files_listbox.curselection()
        if not selected_indices or selected_indices[0] == 0:
            return
        
        for index in selected_indices:
            if index > 0:
                # Trocar na lista de arquivos
                self.document_files[index], self.document_files[index-1] = self.document_files[index-1], self.document_files[index]
                
                # Atualizar listbox
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
                # Trocar na lista de arquivos
                self.document_files[index], self.document_files[index+1] = self.document_files[index+1], self.document_files[index]
                
                # Atualizar listbox
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
            del self.document_files[index]
            self.files_listbox.delete(index)
        
        self.status_var.set(f"{len(selected_indices)} arquivo(s) removido(s) da lista")
    
    def on_closing(self):
        """Executa tarefas de limpeza antes de fechar a aplicação"""
        try:
            # Limpa os arquivos temporários
            if os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir)
        except Exception as e:
            print(f"Erro ao limpar arquivos temporários: {str(e)}")
        
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = DocumentConcatenatorApp(root)
    root.mainloop()