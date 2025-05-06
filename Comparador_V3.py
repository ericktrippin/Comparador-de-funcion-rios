import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd


def importar_mês_passado():
    caminho_arquivo = filedialog.askopenfilename(
        parent=janela,title="Selecione a planilha do mês passado", filetypes=(("Arquivos excel", "*.xlsx *.xls"),))
    if caminho_arquivo:
        global df_mês_passado
        df_mês_passado = pd.read_excel(caminho_arquivo, usecols="C:E", skiprows=4)
        print(f'planilha do mês passado selecionada:\n{df_mês_passado.head()}')
        tree.insert("",'end', values=('Planilha do mês passado carregada',"",""))


def importar_mês_atual():
    caminho_arquivo = filedialog.askopenfilename(
        parent=janela,title="Selecione a planilha do mês atual", filetypes=(("Arquivos excel", "*.xlsx *.xls"),))
    if caminho_arquivo:
        global df_mês_atual
        df_mês_atual =pd.read_excel(caminho_arquivo, usecols="C:E", skiprows=4)
        print(f'planilha do mês passado selecionada:\n{df_mês_atual.head()}')
        comparar_funcionarios()

def comparar_funcionarios():
    rg_mês_atual = df_mês_atual.iloc[:, 2]
    rg_mês_passado = df_mês_passado.iloc[:, 2]
    funcionarios_novos = df_mês_atual[~rg_mês_atual.isin(rg_mês_passado)]

    print(f'Funcionários contratados recentemente:\n{funcionarios_novos}')
    atualizar_tabela    (funcionarios_novos)


def atualizar_tabela(df):
    for i in tree.get_children():
        tree.delete(i)

    for _, row in df.iterrows():
        nome = row.iloc[0]
        funcao = row.iloc[1]
        rg = int(row.iloc[2])
        tree.insert("", "end", values=(nome, funcao, rg))

def main():
    global tree, janela
    janela = tk.Tk()
    janela.title("Comparador de Funcionários")
    janela.geometry("600x400")

    colunas= ('Nome','Função','RG')
    tree = ttk.Treeview(janela, columns=colunas, show='headings')

    for col in colunas:
        tree.heading(col, text=col)
        tree.column(col, anchor=tk.CENTER, width=180)

    tree.pack(expand=True, fill="both", pady=10)

    botao_passado = tk.Button(janela, text="Importar planilha do mês passado", command=importar_mês_passado)
    botao_passado.pack(pady=5)
    botao_atual = tk.Button(janela, text="Importar planilha do mês atual", command=importar_mês_atual)
    botao_atual.pack(pady=5)

    janela.mainloop()

if __name__ == "__main__":
    main()
