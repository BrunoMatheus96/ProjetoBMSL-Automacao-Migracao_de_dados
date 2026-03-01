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
        # 1) lê o Excel
        wb = load_workbook(self.file_path)
        sheet = wb["Planejamento"]

        dados_planilha = []
        for row in range(2, sheet.max_row + 1):
            id = sheet.cell(row, 6).value  
            test_case_name = sheet.cell(row, 1).value  
            step_name = sheet.cell(row, 12).value  
            action = sheet.cell(row, 13).value  
            expected = sheet.cell(row, 14).value  
            system = sheet.cell(row, 17).value
            dados_planilha.append((id, test_case_name, step_name, action, expected, system))

        # 2) abre um novo Word
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

        # loop para cada linha do Excel
        for i, (id, test_case_name, step, action, expected, system) in enumerate(dados_planilha, start=1):

            # adiciona título ou linha do caso de teste
            # adiciona a coluna "Import Test Case name" ao lado de "Caso de teste:"
            doc.add_paragraph(f"Caso de teste: {test_case_name}")

            # monta a frase no formato desejado
            # Step X é apenas a ordem aqui (pode usar step se preferir)
            texto_step = f"{step}: {action}"
            doc.add_paragraph(texto_step)

            # resultado esperado
            doc.add_paragraph(f"Resultado esperado: {expected}")

            # linha em branco entre casos
            doc.add_paragraph("")

        # Salva na pasta docs_gerados
        output_folder = "docs_gerados"
        os.makedirs(output_folder, exist_ok=True)  # cria a pasta se não existir
        caminho_arquivo = os.path.join(output_folder, f"Documento de testes-{system}-TC{id}.docx")
        doc.save(caminho_arquivo)

        print("Documentos gerados com sucesso!")
