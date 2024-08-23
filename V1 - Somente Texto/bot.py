import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import sys

# Definir o caminho do Chrome
chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe %s"

# Carregar a planilha
planilha = openpyxl.load_workbook('Contatos.xlsx')
pagina = planilha['Sheet1']

# Itera sobre as linhas da planilha
for linha in pagina.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value

    if telefone is None or telefone == "":
        print(f"Envio das mensagens concluído!")
        sys.exit()

    mensagem = (
        f"🎉 *Parabéns, {nome}!* 🎉\n\n"
        "Estamos em Janeiro, o SEU mês! 🥳 Sabia que aqui na Amoana a gente AMA celebrar aniversários? 🎂✨\n\n"
        "Então, pra deixar seu mês ainda mais especial, preparamos um presentinho pra você: *10% de desconto* em todas as suas compras "
        "*válido por 30 dias*! 🛍️💃\n\n"
        "💖 É só usar o cupom: *FELIZANIVER10* na hora da compra e aproveitar! Porque aniversário bom é aquele que vem com muito estilo, né?\n\n"
        "Corre pra garantir suas peças favoritas e comemorar seu novo ciclo cheia de charme! 🫶\n\n"
        "Qualquer dúvida, é só chamar! Vamos adorar te ajudar.\n\n"
        "Com carinho,\n"
        "_Anne da Amoana_ 🌸"
    )

    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone=55{telefone}&text={quote(mensagem)}'
        chrome = webbrowser.get(chrome_path)  # Define o Chrome como navegador
        chrome.open(link_mensagem_whatsapp)
        sleep(25)

        pyautogui.press('tab')
        pyautogui.press('enter')
        print(f'Mensagem enviada para {nome}')
        sleep(5)

        pyautogui.hotkey('ctrl', 'w')
        sleep(2)

    except Exception as e:
        print(f'Erro ao enviar mensagem para {nome}: {e}')