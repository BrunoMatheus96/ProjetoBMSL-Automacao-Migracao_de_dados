import tkinter as tk
from tkinter import filedialog


def selec_planilha():
    # Cria uma janela “root” escondida
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal

    # Abre a janela para selecionar um arquivo
    file_path = filedialog.askopenfilename(
        title="Selecione um arquivo", filetypes=[("Todos os arquivos", "*.*")]
    )

    if file_path == "":
        print("Nenhum arquivo selecionado.")
    elif file_path.lower().endswith((".xlsx", ".xls")):
        print("Arquivo Excel selecionado:", file_path)
    else:
        print("O arquivo precisa ser do tipo Excel (.xlsx ou .xls).")


def transformar_planilha(file_path):
    print(f"Transformando a planilha {file_path}... (função ainda não implementada)")
