import pandas as pd
import numpy as np
from tkinter import Tk, filedialog, Label, Button, Entry

def selecionar_arquivo1():
    global arquivo_excel1
    arquivo_excel1 = filedialog.askopenfilename(title="Selecione o arquivo Excel 1", filetypes=[("Arquivos Excel", "*.xlsx")])

def selecionar_arquivo2():
    global arquivo_excel2
    arquivo_excel2 = filedialog.askopenfilename(title="Selecione o arquivo Excel 2", filetypes=[("Arquivos Excel", "*.xlsx")])

def processar():
    try:
        df = pd.read_excel(arquivo_excel1)
        df2 = pd.read_excel(arquivo_excel2)

        coluna_doc1 = entrada_coluna_doc1.get()
        coluna_doc2 = entrada_coluna_doc2.get()
        coluna_lanc2 = entrada_coluna_lanc2.get()

        doc_em_array = df[coluna_doc1].astype(str).str.strip().values  # DOC do primeiro arquivo (errado)
        doc_em_array2 = df2[coluna_doc2].astype(str).str.strip().values  # DOC do segundo arquivo (certo)
        lanc_em_array2 = df2[coluna_lanc2].astype(str).str.strip().values  # Referência para busca

        df2['ordemCerta'] = ""
        tamanho = len(doc_em_array2)
        print(f"Total de itens a verificar: {tamanho}")

        for i in range(tamanho):
            if lanc_em_array2[i] in doc_em_array:
                posicoes = np.where(doc_em_array == lanc_em_array2[i])[0]  # Todas as posições encontradas
                
                if len(posicoes) > 0:  # Garante que encontrou alguma posição válida
                    print(f"Achou: {lanc_em_array2[i]} na posição {posicoes[0]}")
                    df2.loc[i, 'ordemCerta'] = doc_em_array2[posicoes[0]]  # Mantém o valor original
            else:
                df2.loc[i, 'ordemCerta'] = ""  # Deixa vazio caso não encontre

        nome_arquivo = "salvador.xlsx"
        df2.to_excel(nome_arquivo, index=False)
        label_status.config(text=f"Arquivo salvo como {nome_arquivo}")
    except Exception as e:
        label_status.config(text=f"Erro: {e}")

janela = Tk()
janela.title("Interface para Processamento de Excel")

Label(janela, text="Selecione o arquivo com DOC errado:").pack()
Button(janela, text="Selecionar Arquivo 1", command=selecionar_arquivo1).pack()

Label(janela, text="Selecione o arquivo com DOC certo:").pack()
Button(janela, text="Selecionar Arquivo 2", command=selecionar_arquivo2).pack()

Label(janela, text="Nome da coluna DOC errada:").pack()
entrada_coluna_doc1 = Entry(janela)
entrada_coluna_doc1.pack()

Label(janela, text="Nome da coluna DOC correta:").pack()
entrada_coluna_doc2 = Entry(janela)
entrada_coluna_doc2.pack()

Label(janela, text="Nome da coluna de referência:").pack()
entrada_coluna_lanc2 = Entry(janela)
entrada_coluna_lanc2.pack()

Button(janela, text="Processar", command=processar).pack()
label_status = Label(janela, text="")
label_status.pack()

janela.mainloop()
