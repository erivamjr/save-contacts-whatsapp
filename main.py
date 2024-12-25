from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from openpyxl import Workbook
import time

# Configuração do WebDriver
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
service = Service("C:\\chromedriver\\chromedriver-win64\\chromedriver.exe")

# Conectar ao Chrome já aberto
driver = webdriver.Chrome(service=service, options=chrome_options)
print("Conectado ao Chrome já aberto.")

# Alternar para a aba do WhatsApp Web
window_handles = driver.window_handles
for handle in window_handles:
    driver.switch_to.window(handle)
    if "WhatsApp" in driver.title:
        print("WhatsApp Web encontrado.")
        break
else:
    print("WhatsApp Web não encontrado. Certifique-se de que está logado.")
    driver.quit()
    exit()



# Criar uma planilha Excel
wb = Workbook()
ws = wb.active
ws.append(["Nome", "Número"])

# Função para rolar o painel lateral incrementalmente
def incremental_scroll_panel():
    panel = driver.find_element(By.ID, "pane-side")  # Painel lateral
    for _ in range(5):  # Rolagens pequenas por vez
        driver.execute_script("arguments[0].scrollTop += 100", panel)  # Incrementa 100 pixels por vez
        time.sleep(1)  # Aguarde o carregamento

LIMIT = 10  # Substitua pelo número máximo de contatos que deseja capturar

# Extrair contatos com rolagem incremental
try:
    print("Extraindo contatos...")
    captured_names = set()
    reached_end = False  # Flag para verificar se atingiu o final da lista
    previous_count = 0  # Para verificar se novos contatos foram carregados

    while len(captured_names) < LIMIT and not reached_end:
        # Depurar HTML antes de capturar contatos
        

        contacts = driver.find_elements(By.CSS_SELECTOR, '#pane-side div._ak8q > div > span[title]')
        if not contacts:
            print("Nenhum contato encontrado no painel atual.")
            break

        for contact in contacts:
            name = contact.get_attribute('title')
            if name not in captured_names:
                captured_names.add(name)
                print(f"Nome capturado: {name}")
                ws.append([name, "Número desconhecido"])
                if len(captured_names) >= LIMIT:
                    break

        # Rolar incrementalmente e verificar se novos contatos foram carregados
        incremental_scroll_panel()
        print("HTML do painel após rolagem:")
        # print numero de contatos capturados
        print(f"Contatos capturados: {len(captured_names)}")

    # Salvar no Excel
    wb.save("contatos_whatsapp.xlsx")
    print(f"Contatos exportados para contatos_whatsapp.xlsx (Total: {len(captured_names)})")

except Exception as e:
    print(f"Erro ao capturar contatos: {e}")

# Finalizar o WebDriver
driver.quit()
