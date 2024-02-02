import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os 

webbrowser.open('https://web.whatsapp.com/')
sleep(20)


workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Planilha1']

for linha in pagina_clientes.iter_rows(min_row=2):

    nome = linha[0].value
    telefone = linha[1].value   
    mensagem = f'Olá {nome}.'

    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(30)

        botao_enviar = 'seta.png'
        localizacao = pyautogui.locateOnScreen(botao_enviar)

        # Se o botão de envio for encontrado, clicar nele
        if localizacao is not None:
            x, y = pyautogui.center(localizacao)
            pyautogui.click(x, y)
            sleep(5)
        
        pyautogui.hotkey('ctrl','w')
        sleep(5)
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')




    
