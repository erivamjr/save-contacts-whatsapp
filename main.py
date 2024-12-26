from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from openpyxl import Workbook
import time

# Configuração do WebDriver
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
service = Service("C:\\chromedriver\\chromedriver-win64\\chromedriver.exe")

# Conectar ao Chrome já aberto
driver = webdriver.Chrome(service=service, options=chrome_options)
actions = ActionChains(driver)  # Para ações mais precisas
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

# Perguntar ao usuário a quantidade de contatos a buscar
try:
    contact_limit = int(input("Digite o número de contatos que deseja buscar: "))
except ValueError:
    print("Entrada inválida! Usando o limite padrão de 10 contatos.")
    contact_limit = 10

# Função para processar um contato
def process_contact(contact_index):
    try:
        # Localizar o elemento do contato
        contact_xpath = f'//*[@id="pane-side"]/div[2]/div/div/div[{contact_index}]'
        contact_element = driver.find_element(By.XPATH, contact_xpath)

        # Certificar-se de que o elemento está visível
        driver.execute_script("arguments[0].scrollIntoView();", contact_element)
        time.sleep(1)

        # Usar ActionChains para clicar
        actions.move_to_element(contact_element).click().perform()
        time.sleep(2)  # Aguardar o carregamento

        # Verificar se o contato foi selecionado
        inner_xpath = f'{contact_xpath}/div/div'
        inner_element = driver.find_element(By.XPATH, inner_xpath)
        if inner_element.get_attribute("aria-selected") == "true":
            print(f"Contato {contact_index} selecionado. Processando...")

            # Clicar no header
            header = driver.find_element(By.XPATH, '//*[@id="main"]/header/div[2]')
            driver.execute_script("arguments[0].click();", header)
            time.sleep(2)

            # Verificar se o número está disponível
            try:
                number_element = driver.find_element(By.XPATH, '//*[@id="app"]/div/div[3]/div/div[5]/span/div/span/div/div/section/div[1]/div[2]/div/span/span')
                number = number_element.text
                print(f"Número encontrado: {number}")

                # Capturar o nome
                name_element = driver.find_element(By.XPATH, '//*[@id="app"]/div/div[3]/div/div[5]/span/div/span/div/div/section/div[1]/div[2]/h2/div/span')
                name = name_element.text
                print(f"Nome capturado: {name}")

                # Salvar no Excel
                ws.append([name, number])
            except Exception:
                print(f"Contato {contact_index} é um grupo ou não possui número.")
            
            # Fechar o cabeçalho
            close_header()
        else:
            print(f"Contato {contact_index} não está selecionado. Pulando...")
    except Exception as e:
        print(f"Erro ao processar contato {contact_index}: {e}")


# Função para fechar o cabeçalho do contato
def close_header():
    try:
        # Verifica se o elemento existe antes de tentar interagir
        close_button_xpath = '//*[@id="app"]/div/div[3]/div/div[5]/span/div/span/div/header/div/div[1]/div'
        close_buttons = driver.find_elements(By.XPATH, close_button_xpath)
        if close_buttons:
            close_button = close_buttons[0]  # Seleciona o botão se existir
            driver.execute_script("arguments[0].click();", close_button)
            time.sleep(2)
        else:
            print("Nenhum botão de fechamento encontrado. Continuando...")
    except Exception as e:
        print(f"Erro ao fechar cabeçalho: {e}")


# Iterar sobre a lista de contatos
try:
    print("Processando contatos...")
    contact_list = driver.find_element(By.XPATH, '//*[@id="pane-side"]/div[2]/div/div')
    contacts = contact_list.find_elements(By.XPATH, './div')  # Captura todos os itens da lista
    print(f"Total de contatos encontrados: {len(contacts)}")

    # Processar até o limite de contatos especificado pelo usuário
    for i in range(1, min(contact_limit, len(contacts)) + 1):
        process_contact(i)

    # Salvar os dados no arquivo Excel
    wb.save("contatos_whatsapp.xlsx")
    print("Contatos exportados para contatos_whatsapp.xlsx com sucesso.")

except Exception as e:
    print(f"Erro ao capturar contatos: {e}")

# Finalizar o WebDriver
driver.quit()