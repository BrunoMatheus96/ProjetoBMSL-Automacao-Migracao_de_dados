import os
import tkinter as tk
from tkinter import filedialog
from docx import Document
import pandas as pd


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
        # Passo 1 - Limpeza dos dados da planilha
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
        df_tabela = df_tabela.apply(
            lambda col: col.apply(lambda x: x.strip() if isinstance(x, str) else x)
        )  # Remove espaços em branco de cada célula

        # Passo 2 - Juntar os dados da planilha com os dados da main
        df_dados = pd.DataFrame([dados])
        df_final = pd.concat([df_tabela, df_dados], ignore_index=True)
        df_final.tail()

        # Passo 3 - Passar as informações do df_final para uma tabela no Word
        word = Document()

        # criar a tabela de duas colunas
        table = word.add_table(rows=0, cols=2)
        table.style = "Table Grid"

        # preencher com cada linha e cada coluna do df_final
        for _, row in df_final.iterrows():
            for nome_col, valor in zip(df_final.columns, row):
                # cada par (campo/valor) vira uma linha da tabela
                row_cells = table.add_row().cells
                row_cells[0].text = str(nome_col)
                row_cells[1].text = str(valor)

        # salva o documento
        output_folder = "docs_gerados"
        os.makedirs(output_folder, exist_ok=True)

        arquivo_saida = os.path.join(output_folder, "Documento_de_testes.docx")
        word.save(arquivo_saida)

        print("Documento gerado com sucesso!")
