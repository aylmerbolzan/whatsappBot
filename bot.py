import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import sys

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
        f"👋 Olá {nome},\n\n"
        "Seja bem-vindo ao grupo do *Ministério Hero da Lagoinha Domingos Martins*! 🔥\n\n"
        "Aqui, buscamos viver o propósito que Deus tem para cada homem, fortalecendo nossa fé, caráter e compromisso com a palavra de Deus. "
        "Nosso desejo é que sejamos verdadeiros heróis para nossas famílias e referências de Cristo na sociedade.\n\n"
        "📅 Nosso próximo culto será dia *02 de Novembro*, às 18hs – será uma alegria tê-lo conosco!\n"
        "📍 Segue a nossa localização: https://maps.app.goo.gl/bEukr5s5mx2gZiTh7\n\n"
        "Que Deus abençoe sua vida! 🙌"
    )

    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone=55{telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(25)

        pyautogui.press('tab')
        pyautogui.press('enter')
        print(f'Mensagem enviada para {nome}')
        sleep(5)

        pyautogui.hotkey('ctrl', 'w')
        sleep(2)

    except Exception as e:
        print(f'Erro ao enviar mensagem para {nome}: {e}')