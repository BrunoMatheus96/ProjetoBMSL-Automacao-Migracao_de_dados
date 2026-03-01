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
        # 1) lê o Excel inteiro usando pandas
        df = pd.read_excel(self.file_path, sheet_name="Planejamento", engine="openpyxl")

        # limpa nomes de coluna (remove espaços extras antes e depois)
        df.columns = df.columns.str.strip()

        # agora selecione só as colunas existentes:
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

        colunas_existentes = [c for c in colunas_desejadas if c in df.columns]
        df = df[colunas_existentes]

        # 2) cria novo documento Word
        doc = Document()

        # cria tabela no Word para os campos gerais
        table = doc.add_table(rows=0, cols=2)
        table.style = "Table Grid"

        # junta os dados fixos com os dados da planilha
        dados_tabela = list(dados)

        # adiciona os valores da planilha como pares (campo, valor)
        for _, row in df.iterrows():
            dados_tabela += [
                ("Caso de Teste", row.get("Import Test Case Name", "")),
                ("Objetivo", row.get("Test Case Target", "")),
                ("Card de Desenvolvimento", row.get("Card de Desenvolvimento", "")),
                ("Sistema", row.get("Sistema", "")),
                ("QA/LOGIN", row.get("Responsável", "")),
                ("URL do Card", row.get("Link do Card de Desenvolvimento", "")),
            ]

        # preenche a tabela com todos os pares (campo, valor)
        for campo, valor in dados_tabela:
            row_cells = table.add_row().cells
            row_cells[0].text = str(campo)
            row_cells[1].text = str(valor)

        # 3) salva o Word
        output_folder = "docs_gerados"
        os.makedirs(output_folder, exist_ok=True)

        # gera o nome do arquivo de saída
        arquivo_saida = os.path.join(
            output_folder,
            f"Documento_de_testes.docx",
        )
        doc.save(arquivo_saida)

        print("Documento gerado com sucesso!")
