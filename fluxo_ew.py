import tkinter as tk
from tkinter import filedialog
from docx import Document
import pandas as pd
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

    def estruturar_doc(self, dados):
        # 1) Lê a planilha
        df = pd.read_excel(self.file_path, sheet_name="Planejamento", engine="openpyxl")
        df.columns = df.columns.str.strip()

        colunas_desejadas = [
            "ID",
            "Import Test Case Name",
            "Test Case Target",
            "Step Name",
            "Action",
            "Expected Result",
            "Card de Desenvolvimento",
            "Sistema",
            "Responsável",
            "Link do Card de Desenvolvimento",
        ]
        df = df[[c for c in colunas_desejadas if c in df.columns]]

        # Remove NaN e substitui por texto vazio
        df = df.astype(str).fillna("")

        output_folder = "docs_gerados"
        os.makedirs(output_folder, exist_ok=True)

        print("Documentos gerados com sucesso!")
