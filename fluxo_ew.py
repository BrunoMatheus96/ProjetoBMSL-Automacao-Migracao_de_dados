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

    def criar_tabela(self, dados):
        # Limpeza dos dados da planilha
        df = pd.read_excel(self.file_path, sheet_name="Planejamento")
        df.columns = (
            df.columns.str.strip()  # Remove espaços em branco dos nomes das colunas
        )
        # Lista das colunas que quero da planilha
        colunas_desejadas = [
            "Import Test Case Name",
            "Test Case Target",
            "Card de Desenvolvimento",
            "Sistema",
            "Responsável",
            "Link do Card de Desenvolvimento",
        ]

        df_tabela = df[colunas_desejadas].copy()  # Novo DataFrame só com essas colunas

        df_tabela = df_tabela.dropna(
            how="all"
        )  # Remove linhas onde todas as colunas estão vazias
