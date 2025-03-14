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

        # Pegando os nomes das colunas informadas pelo usuário
        coluna_doc1 = entrada_coluna_doc1.get()  # Documento errado
        coluna_doc2 = entrada_coluna_doc2.get()  # Documento certo
        coluna_lanc2 = entrada_coluna_lanc2.get()  # Número de lançamento

        # Convertendo para arrays e removendo espaços extras
        doc_em_array = df[coluna_doc1].astype(str).str.strip().values  # Lista de documentos errados
        doc_em_array2 = df2[coluna_doc2].astype(str).str.strip().values  # Lista de documentos certos
        lanc_em_array2 = df2[coluna_lanc2].astype(str).str.strip().values  # Lista de números de lançamento

        # Criar um dicionário {número de lançamento: documento certo correspondente}
        mapa_lanc_para_certo = {lanc_em_array2[i]: doc_em_array2[i] for i in range(len(lanc_em_array2))}

        # Criando nova coluna 'OrdemCerta' no df (documento errado)
        df['OrdemCerta'] = ""

        # Percorrer os documentos errados e buscar no dicionário de lançamentos
        for i, doc_errado in enumerate(doc_em_array):
            if doc_errado in mapa_lanc_para_certo:
                df.loc[i, 'OrdemCerta'] = mapa_lanc_para_certo[doc_errado]
            else:
                df.loc[i, 'OrdemCerta'] = ""  # Caso não encontre, deixa em branco

        # Criando um novo DataFrame com as colunas na ordem desejada
        df_final = df2[[coluna_doc2, coluna_lanc2]].copy()  # Pega as colunas do arquivo 2 (certo e lançamento)
        df_final[coluna_doc1] = df[coluna_doc1]  # Adiciona a coluna do documento errado
        df_final['OrdemCerta'] = df['OrdemCerta']  # Adiciona a nova coluna com os documentos corrigidos

        # Salvando o resultado em um novo arquivo Excel
        nome_arquivo = "salvador.xlsx"
        df_final.to_excel(nome_arquivo, index=False)
        label_status.config(text=f"Arquivo salvo como {nome_arquivo}")

    except Exception as e:
        label_status.config(text=f"Erro: {e}")

# Criando a interface gráfica
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

Label(janela, text="Nome da coluna de referência (lançamento):").pack()
entrada_coluna_lanc2 = Entry(janela)
entrada_coluna_lanc2.pack()

Button(janela, text="Processar", command=processar).pack()
label_status = Label(janela, text="")
label_status.pack()

janela.mainloop()
