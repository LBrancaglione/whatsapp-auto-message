import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

webbrowser.open('https://web.whatsapp.com/')
sleep(30)

#Ler planilha, identificar telefone e nome
workbook = openpyxl.load_workbook('Clientes.xlsx')
pagina_clientes = workbook['Sheet1']

for linha in pagina_clientes.iter_rows(min_row=2):
    #nome, telefone, mesnagem
    nome = linha[0].value
    telefone = linha[1].value
    mensagem = f'Olá {nome} já viu nossos novos produtos ?'

    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(12)
        sleep(2)
        pyautogui.hotkey('enter')
        sleep(2)
        pyautogui.hotkey('ctrl','w')
        sleep(2)
    except: 
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}')
    