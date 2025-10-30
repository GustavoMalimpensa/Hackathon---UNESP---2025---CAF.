"""AUTOMAÇÃO DE ENVIO DE MENSAGENS PARA WHATSAPP"""

import webbrowser
import openpyxl
from urllib.parse import quote
from time import sleep
import pyautogui 

workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['basedados']


chrome_path = 'C:\Program Files\Google\Chrome\Application\chrome.exe'  
webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))


webbrowser.get('chrome').open('https://web.whatsapp.com/')
sleep(10)

for linha in pagina_clientes.iter_rows(min_row=2) :
    
    nome = linha[0].value
    telefone = linha[1].value

    
    mensagem = (
        "Olá {nome}, boa tarde!\n\n"
        "Somos da CAF Máquinas! 🧑🏻‍💻\n\n"
        "Estamos entrando em contato referente ao status de manutenção da sua máquina! "
        "Nossa equipe está trabalhando com máxima prioridade para solucionar o seu problema o mais rápido possível!🤩\n\n"
        "Caso tenha alguma dúvida ou deseje mais informações, entre em contato conosco ou acesse nosso site.\n\n"
        "🌐 Confira como ficaria: https://www.cafmaquinas.com.br//"
    ).format(nome=nome)

    
    sleep(8)
   
    try:
        
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
       
        webbrowser.get('chrome').open(link_mensagem_whatsapp)
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(4)
        pyautogui.click(seta[0], seta[1])
        sleep(4)
       
        pyautogui.hotkey('ctrl', 'w')
        sleep(4)
    except: 
        print(f'Não foi possivel mandar mensagem para {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf=8') as arquivo:
            arquivo.write(f'{nome},{telefone}')

    
