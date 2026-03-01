import tkinter as tk
from tkinter import filedialog


class FluxoEW:
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

    def gerar_documentos(self):
        print("Gerando documentos word a partir do arquivo:", self.file_path)
