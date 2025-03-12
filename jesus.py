import pandas as pd
import numpy as np
from tkinter import Tk, filedialog, Label, Button, Entry
#lendo os arquivos e transformando em df(ki orgulho)
def selecionarArquivo():
    global arquivo_excel
    arquivo_excel = filedialog.askopenfilename(title="Selecione o arquivo Excel 1", filetypes=[("Arquivos Excel", "*.xlsx")])
  
def selecionar_arquivo2():
    global arquivo_excel2
    arquivo_excel2 = filedialog.askopenfilename(title="Selecione o arquivo Excel 2", filetypes=[("Arquivos Excel", "*.xlsx")])

def processar():
 df = pd.read_excel(arquivo_excel)
 df2 = pd.read_excel(arquivo_excel2)

    # Pegar os nomes das colunas fornecidos pelo usuário
 coluna_doc1 = entrada_coluna_doc1.get()
 coluna_doc2 = entrada_coluna_doc2.get()
 coluna_lanc2 = entrada_coluna_lanc2.get()

    # Criar arrays com os valores das colunas
 doc_em_array = df[coluna_doc1].values  # DOC do primeiro arquivo (errado)
 doc_em_array2 = df2[coluna_doc2].values  # DOC do segundo arquivo (certo)
 lanc_em_array2 = df2[coluna_lanc2].values

#to criando uma coluna nova no df2 com o nome de ordem certa(ainda bem)
 df2['ordemCerta'] = 0


 for i in range(len(doc_em_array)):
    if lanc_em_array2[i] in doc_em_array:
        posicao = np.where(doc_em_array == lanc_em_array2[i])[0][0]
        print("Achou")
        df2.loc[i, 'ordemCerta'] = doc_em_array2[posicao]
    else:
        print("Não achou")
        df2.loc[i, 'ordemCerta'] = 0
        print(df2['ordemCerta'])
        
 ordemCertas = df2['ordemCerta'].values
        
# print(doc_em_array)
# print("------------------------------------------------")
# print(doc_em_array2)
# print("------------------------------------------------")
# print(lanc_em_array2)
# print("------------------------------------------------")
# print(ordemCertas)

 nome_arquivo = "salvador.xlsx"
 df2.to_excel(nome_arquivo, index=False)
        
#a porra da interface grafica q me ferrei pra entender e o copilot ajudou
janela = Tk()
janela.title("Interface para Processamento de Excel")


label_arquivo1 = Label(janela, text="Selecione o arquivo Excel 1:")
label_arquivo1.pack()
botao_arquivo1 = Button(janela, text="Selecionar Arquivo 1", command=selecionarArquivo)
botao_arquivo1.pack()

label_arquivo2 = Label(janela, text="Selecione o arquivo Excel 2:")
label_arquivo2.pack()
botao_arquivo2 = Button(janela, text="Selecionar Arquivo 2", command=selecionar_arquivo2)
botao_arquivo2.pack()


label_coluna_doc1 = Label(janela, text="Nome da coluna DOC no Arquivo 1:")
label_coluna_doc1.pack()
entrada_coluna_doc1 = Entry(janela)
entrada_coluna_doc1.pack()


label_coluna_doc2 = Label(janela, text="Nome da coluna DOC no Arquivo 2:")
label_coluna_doc2.pack()
entrada_coluna_doc2 = Entry(janela)
entrada_coluna_doc2.pack()


label_coluna_lanc2 = Label(janela, text="Nome da coluna LANC no Arquivo 2:")
label_coluna_lanc2.pack()
entrada_coluna_lanc2 = Entry(janela)
entrada_coluna_lanc2.pack()


botao_processar = Button(janela, text="Processar", command=processar)
botao_processar.pack()


label_status = Label(janela, text="")
label_status.pack()


janela.mainloop()