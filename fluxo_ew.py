import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
from docx import Document
import os

class FluxoEW:
    def __init__(self):
        self.file_path = ""

    def selec_planilha(self):
        try:
            # Cria uma janela “root” escondida
            root = tk.Tk()
            root.withdraw()  # Esconde a janela principal

            # Abre a janela para selecionar um arquivo
            self.file_path = filedialog.askopenfilename(
                title="Selecione um arquivo", filetypes=[("Todos os arquivos", "*.*")]
            )

            if self.file_path == "":
                print("Nenhum arquivo selecionado.")
            elif self.file_path.lower().endswith((".xlsx", ".xls")):
                print("Arquivo Excel selecionado:", self.file_path)
            else:
                print("O arquivo precisa ser do tipo Excel (.xlsx ou .xls).")
        except Exception as e:
            print(f"Erro ao selecionar o arquivo: {e}")

    def estrututar_doc(self, dados):
        # caminhos
        output_folder = "docs_gerados"
        os.makedirs(output_folder, exist_ok=True)  # cria a pasta se não existir

        doc = Document()

        # Cria uma tabela com 2 colunas (exemplo para título/valor)
        table = doc.add_table(rows=0, cols=2)

        # Aqui você define um estilo que já tem bordas visíveis
        table.style = "Table Grid"  # estilo padrão com bordas

        # Preenche as linhas
        for campo, valor in dados:
            row_cells = table.add_row().cells
            row_cells[0].text = campo
            row_cells[1].text = valor

        # Salva na pasta docs_gerados
        caminho_arquivo = os.path.join(output_folder, f"Documento sem os testes.docx")
        doc.save(caminho_arquivo)

    def documentar_casos(self):
        print("Documentando casos de teste...")