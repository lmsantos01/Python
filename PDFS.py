import PyPDF2
import os
import time
import tkinter as tk
from tkinter import messagebox

start_time = time.time()

def extrair_terceira_info_linha_0(texto):
    linhas = texto.split('\n')
    if len(linhas) > 0:
        informacoes = linhas[0].split()
        if len(informacoes) >= 3:
            return informacoes[3].strip()
    return None

def extrair_terceira_info_linha_1(texto):
    linhas = texto.split('\n')
    if len(linhas) > 0:
        informacoes = linhas[0].split()
        if len(informacoes) >= 3:
            return informacoes[5].strip()
    return None

def extrair_terceira_info_linha_7(texto):
    linhas = texto.split('\n')
    if len(linhas) >= 7:
        informacoes = linhas[5].split()
        if len(informacoes) >= 3:
            return informacoes[3].strip()
    return None

diretorio_saida = 'C:\\lmsantos\\Pontos'
os.makedirs(diretorio_saida, exist_ok=True)

with open('C:\\lmsantos\\ponr010.pdf', 'rb') as arquivo:
    leitor = PyPDF2.PdfFileReader(arquivo)
    nome_arquivo_anterior = None
    escritor = None
    for i in range(leitor.numPages):
        pagina = leitor.getPage(i)
        texto = pagina.extractText()
        nome_arquivo = (f'PONTO_{extrair_terceira_info_linha_7(texto)}_{extrair_terceira_info_linha_1(texto)[2:5]}_{extrair_terceira_info_linha_1(texto)[5:10]}')
        if nome_arquivo and nome_arquivo != nome_arquivo_anterior:
            if escritor:
                nome_arquivo_formatado = nome_arquivo_anterior.replace('/', '')  # Usar o nome do arquivo anterior
                caminho_arquivo_saida = os.path.join(diretorio_saida, f'{nome_arquivo_formatado}.pdf')
                with open(caminho_arquivo_saida, 'wb') as arquivo_saida:
                    escritor.write(arquivo_saida)
            escritor = PyPDF2.PdfFileWriter()
        if escritor:
            escritor.addPage(pagina)
        nome_arquivo_anterior = nome_arquivo
    if escritor:
        nome_arquivo_formatado = nome_arquivo_anterior.replace('/', '')
        caminho_arquivo_saida = os.path.join(diretorio_saida, f'{nome_arquivo_formatado}.pdf')
        with open(caminho_arquivo_saida, 'wb') as arquivo_saida:
            escritor.write(arquivo_saida)

for i in range(1000000):
    pass
end_time = time.time()
execution_time = end_time - start_time
messagebox.showinfo('Espelho de Ponto', f"As páginas do PDF foram divididas e salvas em '{diretorio_saida}' em {execution_time:.6f} segundos.")