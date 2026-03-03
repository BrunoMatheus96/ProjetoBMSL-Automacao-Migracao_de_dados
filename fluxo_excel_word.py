import os
import tkinter as tk
from tkinter import filedialog
from docx import Document
import pandas as pd


class excel_word:
    def __init__(self):
        self.file_path = ""

    def selec_planilha(self):
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

    def criar_doc(self, dados):
        try:
            # Passo 1 - Ler a planilha Excel e criar um novo DataFrame
            df = pd.read_excel(self.file_path, sheet_name="Planejamento")
            df.columns = (
                df.columns.str.strip()  # Remove espaços em branco dos nomes das colunas
            )
            colunas_desejadas = [
                "Import Test Case Name",
                "Test Case Target",
                "Card de Desenvolvimento",
                "Sistema",
                "Responsável",
                "Link do Card de Desenvolvimento",
            ]
            df_tabela = df[
                colunas_desejadas
            ].copy()  # Novo DataFrame só com essas colunas
            df_tabela = df_tabela.dropna(
                how="all"
            )  # Remove linhas onde todas as colunas estão vazias
            df_tabela = df_tabela.apply(
                lambda col: col.apply(lambda x: x.strip() if isinstance(x, str) else x)
            )  # Remove espaços em branco de cada célula

            # Passo 2 - Gerar os documentos Word
            # Para cada linha do planilha, criar um Word
            for i, row in df_tabela.iterrows():

                # pega o valor de ID da planilha original
                id_valor = df.loc[i, "ID"]

                # trata valores ausentes e converte para string
                if pd.isna(id_valor):
                    id_str = ""
                else:
                    id_str = str(int(id_valor))  # converte para inteiro antes de string

                # Cria um novo documento Word e as linhas da tabela
                doc = Document()
                table = doc.add_table(rows=0, cols=2)
                table.style = "Table Grid"

                # 🚧adiciona imagens no cabeçalho do documento🚧

                # 🚧adiciona o título no cabeçalho do documento🚧

                # adiciona dados da planilha na tabela
                for c in colunas_desejadas:
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(c)
                    row_cells[1].text = str(row.get(c, ""))

                # adiciona dados extras
                for chave, valor in dados.items():
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(chave)
                    row_cells[1].text = str(valor)

                output_folder = "docs_gerados"
                os.makedirs(output_folder, exist_ok=True)

                nome_arquivo = (
                    f"Documento de teste - {row['Sistema']} - TC{id_str}.docx"
                )
                caminho = os.path.join(output_folder, nome_arquivo)

                doc.save(caminho)
                print(f"Documento gerado: {nome_arquivo}")
        except:
            print("Erro ao criar o(s) documento(s), verifique se o arquivo selecionado é Excel (.xlsx ou .xls) e tente novamente.")

    def adicionar_steps(self):
        print("🚧")
