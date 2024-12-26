import subprocess
import random
import openpyxl
import re
from shlex import quote
from time import sleep
import pyautogui

# Caminho do executável do Google Chrome
chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe"

# Abrir WhatsApp Web no Google Chrome
subprocess.Popen([chrome_path, 'https://web.whatsapp.com/'])
sleep(30)

# Carregar a planilha de contatos
workbook = openpyxl.load_workbook('contatos_whatsapp.xlsx')
workbook_contatos = workbook['Sheet']

# Lista de mensagens
mensagens = [
    "Que o espírito de Natal 🎄 encha seu coração ❤️ de paz ✨ e alegria 😊.",
    "Neste Natal 🎅, desejo que todos os seus sonhos 🌟 se realizem!",
    "Feliz Natal! 🎁 Que sua casa seja cheia de amor 💕 e harmonia 🕊️.",
    "Que neste Natal 🎄, você encontre motivos para sorrir 😄 e agradecer 🙏.",
    "Desejo um Natal cheio de luz ✨, amor ❤️ e momentos inesquecíveis 🌟.",
    "Que o Natal 🌟 seja um momento de renovação 🌿 e esperança 💫 em sua vida.",
    "Feliz Natal! 🎅 Que você receba bênçãos de saúde 💪, amor 💖 e felicidade 😊.",
    "Que a magia do Natal ✨ esteja presente em cada momento do seu dia 🎄.",
    "Desejo que o Natal 🎁 traga paz 🕊️ e união 🤝 para você e sua família 👨‍👩‍👧‍👦.",
    "Que o espírito natalino 🎅 encha sua vida de luz ✨ e alegria 😊.",
    "Feliz Natal! 🎄 Que seu coração ❤️ transborde de amor 💕 e gratidão 🙏.",
    "Que o Natal 🎁 seja repleto de sorrisos 😊, abraços 🤗 e boas lembranças 📸.",
    "Desejo um Natal especial 🎅, cheio de paz 🕊️ e harmonia ✨.",
    "Que neste Natal 🎄 você celebre 🎉 a vida 🌟 e as pessoas que ama ❤️.",
    "Que o espírito de Natal 🎅 traga paz 🕊️, amor ❤️ e esperança 💫 ao seu coração.",
    "Feliz Natal! 🎄 Que você se sinta abençoado(a) 🙏 neste dia tão especial ✨.",
    "Desejo que o Natal 🌟 seja um recomeço 💫 cheio de boas energias 🌿.",
    "Que sua noite de Natal 🌙 seja iluminada ✨ e cheia de amor ❤️.",
    "Feliz Natal! 🎁 Que a felicidade 😊 e a gratidão 🙏 façam parte deste momento 🌟.",
    "Que o Natal 🎄 traga a você e sua família 👨‍👩‍👧‍👦 muita paz 🕊️ e prosperidade 💰."
]

for line in workbook_contatos.iter_rows(min_row=2):
    # Pegar nome e telefone
    name = line[0].value.split(' ')[0]
    phone = re.sub(r"[^\d]", "", line[1].value)

    # Escolher mensagem aleatória
    mensagem_randomica = random.choice(mensagens)

    try:
        # Criar o link com a mensagem personalizada usando WhatsApp Web
        link_mensagem_whatsapp = f"https://web.whatsapp.com/send?phone={phone}&text={quote(f'Olá {name}, cheguei tarde mas ainda é natal... vim dar uma passadinha rapidinho aqui pra lhe desejar {mensagem_randomica}')}"
        
        # Abrir o link no Chrome usando subprocess
        subprocess.Popen([chrome_path, link_mensagem_whatsapp])
        sleep(random.randint(5, 20))  # Pausa aleatória entre 5 e 20 segundos

        # Clicar no botão "Enviar" e fechar a aba
        send = pyautogui.locateCenterOnScreen('send.png')
        sleep(random.randint(5, 20))  # Pausa aleatória
        pyautogui.click(send[0], send[1])
        sleep(random.randint(5, 20))  # Pausa aleatória
        pyautogui.hotkey('ctrl', 'w')
        sleep(random.randint(5, 20))  # Pausa aleatória
    except Exception as e:
        print(f'Erro ao enviar mensagem para {name}: {e}')
        with open('log.txt', 'a', encoding='utf-8') as f:
            f.write(f'{name}, {phone}\n')
        pyautogui.hotkey('ctrl', 'w')
        sleep(random.randint(5, 20))  # Pausa aleatória