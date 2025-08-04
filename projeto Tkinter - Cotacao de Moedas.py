from tkinter import ttk
import tkinter as tk
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename
import requests
import pandas as pd
from datetime import datetime
import numpy as np


requisicao = requests.get("https://economia.awesomeapi.com.br/json/all")
dicionario_moedas = requisicao.json()

lista_moedas = list(dicionario_moedas.keys())


def pegarCotacao():
    moeda = combobox_SelecionarMoeda.get()
    data_cotacao = calendario_moeda.get()
    ano = data_cotacao[-4:]
    mes = data_cotacao[3:5]
    dia = data_cotacao[:2]
    link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}"
    requisicao_moeda = requests.get(link)
    cotacao = requisicao_moeda.json()
    valor_moeda = cotacao[0]['bid']
    
    label_mostrarCotacao = tk.Label(text=f"A cotação da moeda {moeda} no dia {data_cotacao} é de R${valor_moeda}", anchor='e')
    label_mostrarCotacao.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="news")


def pegarArquivo():
    caminho_arquivo = askopenfilename(title="Selecione o Arquivo de Moeda")
    var_caminhoArquivo.set(caminho_arquivo)
    if caminho_arquivo:
        arquivoSelecionado['text'] = f"Arquivo Selecionado: {caminho_arquivo}"


def atualizar_cotacoes():
    try:
        df = pd.read_excel(var_caminhoArquivo.get())
        
        dataInicial = calendario_dataInicial.get()
        ano_inicial = dataInicial[-4:]
        mes_inicial = dataInicial[3:5]
        dia_inicial = dataInicial[:2]
        
        dataFinal = calendario_dataFinal.get()
        ano_final = dataFinal[-4:]
        mes_final = dataFinal[3:5]
        dia_final = dataFinal[:2]

        
        for moeda in df.iloc[:, 0]:
            link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/31?" \
                f"start_date={ano_inicial}{mes_inicial}{dia_inicial}&" \
                f"end_date={ano_final}{mes_final}{dia_final}"

            requisicao_moeda = requests.get(link)
            cotacoes = requisicao_moeda.json()

            for cotacao in cotacoes:
                timestamp = int(cotacao['timestamp'])
                bid = float(cotacao['bid'])
                data = datetime.fromtimestamp(timestamp)
                data = data.strftime('%d/%m/%Y')
                if data not in df:
                    print(data)
                    df[data] = np.nan
                
                df.loc[df.iloc[:, 0] == moeda, data] = bid
        
        df.to_excel("Teste.xlsx")
        atualizarCotacoes['text'] = "Arquivo Atualizado com Sucesso!"
    except:
        atualizarCotacoes['text'] = "Selecione um Arquivo Excel no Cormato Correto"

janela = tk.Tk()

janela.title("Ferramenta de Cotação de Moedas")

texto1 = tk.Label(text="Cotação de 1 moeda específica", borderwidth=2, relief='solid')
texto1.grid(row=0, column=0, padx=10, pady=10, sticky="news", columnspan=3)

selecionar_moeda = tk.Label(text="Selecione a moeda que desejar consultar: ", anchor='e')
selecionar_moeda.grid(row=1, column=0, padx=10, pady=10, sticky="news", columnspan=2)

combobox_SelecionarMoeda = ttk.Combobox(values=lista_moedas)
combobox_SelecionarMoeda.grid(row=1, column=2, padx=10, pady=10, sticky="news")

selecionar_dia = tk.Label(text="Selecione o dia que deseja pegar a cotação: ", anchor='e')
selecionar_dia.grid(row=2, column=0, padx=10, pady=10, sticky="news", columnspan=2)

calendario_moeda = DateEntry(year=2025, locale='pt_br')
calendario_moeda.grid(row=2, column=2, padx=10, pady=10, sticky='news')

textoCotacao = tk.Label(text="")
textoCotacao.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky='news')

botao_pegarCotacao = tk.Button(text="Pegar Cotação", command=pegarCotacao)
botao_pegarCotacao.grid(row=3, column=2, padx=10, pady=10, sticky='news')


# Cotação de várias moedas

texto2 = tk.Label(text="Cotação de Múltiplas Moedas", borderwidth=2, relief='solid')
texto2.grid(row=4, column=0, padx=10, pady=10, sticky="news", columnspan=3)

selecione_arquivo = tk.Label(text="Selecione um arquivo em Excel com as Moedas na Coluna A: ")
selecione_arquivo.grid(row=5, column=0, padx=10, pady=10, sticky="news", columnspan=2)


var_caminhoArquivo = tk.StringVar()



botao_selecionarArquivo = tk.Button(text="Clique aqui para selecionar", command=pegarArquivo)
botao_selecionarArquivo.grid(row=5, column=2, padx=10, pady=10, sticky="news")

arquivoSelecionado = tk.Label(text="Nenhum Arquivo Selecionado", anchor='e')
arquivoSelecionado.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky='news')

dataInicial = tk.Label(text="Data Inicial", anchor='e')
dataInicial.grid(row=7, column=0, padx=10, pady=10, sticky='news')

dataFinal = tk.Label(text="Data Final", anchor='e')
dataFinal.grid(row=8, column=0, padx=10, pady=10, sticky='news')

calendario_dataInicial = DateEntry(year=2025, locale='pt_br')
calendario_dataInicial.grid(row=7, column=1, padx=10, pady=10, sticky='news')

calendario_dataFinal = DateEntry(year=2025, locale='pt_br')
calendario_dataFinal.grid(row=8, column=1, padx=10, pady=10, sticky='news')

botao_atualizarCotacoes = tk.Button(text="Atulizar Cotações", command=atualizar_cotacoes)
botao_atualizarCotacoes.grid(row=9, column=0, padx=10, pady=10, stick='news')

atualizarCotacoes = tk.Label(text="")
atualizarCotacoes.grid(row=9, column=1, columnspan=2, padx=10, pady=10, sticky='news')

botao_fechar = tk.Button(text="Fechar", command=janela.quit)
botao_fechar.grid(row=10, column=2, padx=10, pady=10, sticky='news')

janela.mainloop()

