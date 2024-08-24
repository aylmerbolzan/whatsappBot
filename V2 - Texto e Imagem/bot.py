import openpyxl
import pywhatkit
import pyautogui
import sys
from time import sleep

# Carregar a planilha
planilha = openpyxl.load_workbook('contatos.xlsx')
pagina = planilha['Sheet1']

# Caminho da imagem a ser enviada
caminho_imagem = "C:/Users/aylmer.bolzan/Downloads/imagem.jpeg"

# Itera sobre as linhas da planilha
for linha in pagina.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value

    if telefone is None or telefone == "":
        print(f"Envio das mensagens concluído!")
        sys.exit()

    mensagem = (
        f"Oi {nome}! ✨\n\n"
        "Separamos sugestões incríveis para você brilhar no Carnaval! 🎶💃\n\n"
        "Confira o catálogo especial que preparamos pra você:\n"
        "👉 https://amoana.shop/catalogo\n\n"
        "Qualquer dúvida ou se precisar de ajuda para escolher, estou aqui pra te ajudar! 🫶\n\n"
        "Beijos,\n"
        "Anne da *Amoana* 🌸"
    )

    try:
        # Enviar a imagem com a mensagem na legenda
        pywhatkit.sendwhats_image(
            receiver=f"+55{telefone}",
            img_path=caminho_imagem,
            caption=mensagem,
            wait_time=10
        )
        print(f'Imagem e mensagem enviadas para {nome}')

        # Espera 1 segundo antes de fechar a aba
        sleep(1)

        # Fecha a aba ativa no navegador (CTRL + W)
        pyautogui.hotkey('ctrl', 'w')
        print(f'Aba do WhatsApp fechada para {nome}')

        sleep(3)  # Pequena pausa antes do próximo contato

    except Exception as e:
        print(f'Erro ao enviar mensagem para {nome}: {e}')
