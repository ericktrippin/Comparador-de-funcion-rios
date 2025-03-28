import tkinter as tk
from tkinter import filedialog
import pandas as pd

def importar_mês_passado():
    caminho_arquivo = filedialog.askopenfilename(title="Selecione a planilha do mês passado", filetypes=(("Arquivos excel", "*.xlsx"),))
    if caminho_arquivo:
        global df_mês_passado
        df_mês_passado = pd.read_excel(caminho_arquivo, usecols="C:E", skiprows=4)
        print(f'planilha do mês passado selecionada:\n{df_mês_passado.head()}')
        listbox.insert(tk.END, "Planilha do mês passado carregada com sucesso!")


def importar_mês_atual():
    caminho_arquivo = filedialog.askopenfilename(title="Selecione a planilha do mês atual", filetypes=(("Arquivos excel", "*.xlsx"),))
    if caminho_arquivo:
        global df_mês_atual
        df_mês_atual =pd.read_excel(caminho_arquivo, usecols="C:E", skiprows=4)
        print(f'planilha do mês passado selecionada:\n{df_mês_atual.head()}')
        comparar_funcionarios()

def comparar_funcionarios():
    rg_mês_atual = df_mês_atual.iloc[:, 2]  # RG está na coluna 2 (E)
    rg_mês_passado = df_mês_passado.iloc[:, 2]  # RG está na coluna 2 (E)
    funcionarios_novos = df_mês_atual[~rg_mês_atual.isin(rg_mês_passado)]

    print(f'Funcionários contratados recentemente:\n{funcionarios_novos}')
    atualizar_lista(funcionarios_novos)


def atualizar_lista(df):
    listbox.delete(0, tk.END)
    listbox.insert(tk.END,("Nome - Função - RG"))

    for _, row in df.iterrows():
        rg = int(row.iloc[2])
        listbox.insert(tk.END, f"{row.iloc[0]} - {row.iloc[1]} - {rg}")


janela = tk.Tk()
janela.title("Comparador de Funcionários")
janela.geometry("600x400")

listbox = tk.Listbox(janela, width=80, height=15)
listbox.pack(pady=10)

botao_passado = tk.Button(janela, text="importar planilha do mês passado", command=importar_mês_passado)
botao_passado.pack(pady=10, padx=10)
botao_atual = tk.Button(janela, text="importar planilha do mês atual", command=importar_mês_atual)
botao_atual.pack(pady=10, padx=10)
janela.mainloop()
