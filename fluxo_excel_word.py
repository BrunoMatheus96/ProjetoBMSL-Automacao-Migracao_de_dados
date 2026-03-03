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
            print("O arquivo precisa ser do tipo .xlsx ou .xls.")

    def criar_doc(self, dados):
        try:

            def preparar_dados():
                df = pd.read_excel(self.file_path, sheet_name="Planejamento")
                df.columns = df.columns.str.strip()

                colunas = [
                    "Import Test Case Name",
                    "Test Case Target",
                    "Card de Desenvolvimento",
                    "Sistema",
                    "Responsável",
                    "Link do Card de Desenvolvimento",
                ]

                df_tabela = df[colunas].copy()
                df_tabela = df_tabela.dropna(how="all")
                df_tabela = df_tabela.apply(
                    lambda col: col.apply(
                        lambda x: x.strip() if isinstance(x, str) else x
                    )
                )

                return df, df_tabela

            # chamando a função interna
            df, df_tabela = preparar_dados()

            def criar_um_documento(df, row):
                # função que cria apenas um doc a partir de uma linha
                from docx import Document

                id_valor = df.loc[row.name, "ID"]
                id_str = "" if pd.isna(id_valor) else str(int(id_valor))

                doc = Document()
                table = doc.add_table(rows=0, cols=2)
                table.style = "Table Grid"

                # adiciona dados da planilha
                for c in row.index:
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(c)
                    row_cells[1].text = str(row[c])

                # adiciona dados extras
                for chave, valor in dados.items():
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(chave)
                    row_cells[1].text = str(valor)

                return doc, id_str

            def salvar_documento(doc, id_str, sistema):
                import os

                output_folder = "docs_gerados"
                os.makedirs(output_folder, exist_ok=True)

                nome_arquivo = f"Documento de teste - {sistema} - TC{id_str}.docx"
                caminho = os.path.join(output_folder, nome_arquivo)

                doc.save(caminho)
                print(f"Documento gerado: {nome_arquivo}")

            # agora o loop chama as funções internas para cada linha
            for _, row in df_tabela.iterrows():
                doc, id_str = criar_um_documento(df, row)
                salvar_documento(doc, id_str, row["Sistema"])

        except:
            print(
                "\nErro ao criar o(s) documento(s), verifique se o arquivo selecionado é .xlsx ou .xls e tente novamente."
            )

    def adicionar_steps(self):
        print("🚧")
