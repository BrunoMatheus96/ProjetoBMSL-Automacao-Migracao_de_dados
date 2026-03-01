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
        wb = load_workbook(self.file_path, data_only=True)
        sheet = wb["Planejamento"]

        dados_planilha = []
        for row in range(2, sheet.max_row + 1):
            id = sheet.cell(row=row, column=6).value
            test_case_name = sheet.cell(row=row, column=9).value
            objective = sheet.cell(row=row, column=10).value
            step_name = sheet.cell(row=row, column=12).value
            action = sheet.cell(row=row, column=13).value
            expected_result = sheet.cell(row=row, column=14).value
            history = sheet.cell(row=row, column=16).value
            system = sheet.cell(row=row, column=17).value
            qa = sheet.cell(row=row, column=20).value
            url_card = sheet.cell(row=row, column=25).value
            dados_planilha.append(
                (
                    id,
                    test_case_name,
                    objective,
                    step_name,
                    action,
                    expected_result,
                    history,
                    system,
                    qa,
                    url_card,
                )
            )

        # 2) abre um novo Word
        doc = Document()

        # Cria uma tabela com 2 colunas (exemplo para título/valor)
        table = doc.add_table(rows=0, cols=2)

        # Aqui você define um estilo que já tem bordas visíveis
        table.style = "Table Grid"  # estilo padrão com bordas

        # preparar uma nova lista para tabela
        dados_tabela = []

        # adiciona os dados fixos que já vêm na lista `dados`
        dados_tabela.extend(dados)

        # agora pega apenas os campos desejados de cada linha da planilha
        for (
            id,
            test_case_name,
            objective,
            step_name,
            action,
            expected_result,
            history,
            system,
            qa,
            url_card,
        ) in dados_planilha:
            # cria pares (campo, valor) para os itens que você quer mostrar
            dados_tabela.append(("Caso de Teste", test_case_name))
            dados_tabela.append(("Objetivo", objective))
            dados_tabela.append(("Histórico", history))
            dados_tabela.append(("Sistema", system))
            dados_tabela.append(("QA/LOGIN", qa))
            dados_tabela.append(("URL do Card", url_card))

        # Preenche as linhas
        for campo, valor in dados_tabela:
            row_cells = table.add_row().cells  # adiciona uma nova linha
            row_cells[0].text = campo
            row_cells[1].text = valor.__str__()  # converte valor para string

        # loop para cada linha do Excel
        for i, (
            id,
            test_case_name,
            objective,
            step,
            action,
            expected,
            history,
            system,
            qa,
            url_card,
        ) in enumerate(dados_planilha, start=1):

            # monta a frase no formato desejado
            texto_step = f"Step {i} {step}: {action}"
            doc.add_paragraph(texto_step)

            # resultado esperado
            doc.add_paragraph(f"Resultado esperado: {expected}")

            # linha em branco entre casos
            doc.add_paragraph("")

        # Salva na pasta docs_gerados
        output_folder = "docs_gerados"
        os.makedirs(output_folder, exist_ok=True)  # cria a pasta se não existir
        caminho_arquivo = os.path.join(
            output_folder, f"Documento de testes-{system}-TC{id}.docx"
        )
        doc.save(caminho_arquivo)

        print("Documentos gerados com sucesso!")
