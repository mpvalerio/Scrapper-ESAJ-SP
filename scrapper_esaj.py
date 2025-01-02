from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd

# Carregar números de processo do Excel
excel_path = 'PATH EXCEL FILE'
df = pd.read_excel(excel_path)
process_numbers = df['N.processo'].tolist()

# Configurar o driver do Selenium
options = webdriver.ChromeOptions()
options.add_argument('--start-maximized')
options.add_argument('--disable-blink-features=AutomationControlled')

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 20)

# URL inicial
url = "https://esaj.tjsp.jus.br/cpopg/open.do?gateway=true"

def download_process(process_number):
    try:
        # Acessar o site inicial
        driver.get(url)

        # Inserir número do processo
        input_field = wait.until(EC.presence_of_element_located((By.ID, "numeroDigitoAnoUnificado")))
        input_field.clear()
        input_field.send_keys(process_number)

        # Clicar em "Consultar"
        consult_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Consultar']")))
        consult_button.click()

        # Clicar em "Visualizar autos"
        view_files_button = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Visualizar autos")))
        view_files_button.click()

        # Alternar para a nova aba aberta
        driver.switch_to.window(driver.window_handles[-1])

        # Selecionar "Todas"
        select_all = wait.until(EC.element_to_be_clickable((By.ID, "selecionarButton")))
        select_all.click()

        # Clicar em "Versão para impressão"
        print_version_button = wait.until(EC.element_to_be_clickable((By.ID, "salvarButton")))
        print_version_button.click()

        # Selecionar "Um arquivo para cada documento"
        single_file_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='popupDividirDocumentos']/div[1]/fieldset/p[2]/label")))
        single_file_option.click()

        # Clicar em "Continuar"
        continue_button = wait.until(EC.element_to_be_clickable((By.ID, "botaoContinuar")))
        continue_button.click()

        # Esperar a geração do arquivo
        time.sleep(20)  # Ajuste se necessário para seu ambiente

        # Clicar em "Salvar o documento"
        save_button = wait.until(EC.element_to_be_clickable((By.ID, "btnDownloadDocumento")))
        save_button.click()

        # Fechar a aba atual e voltar para a aba principal
        driver.close()
        driver.switch_to.window(driver.window_handles[0])

    except Exception as e:
        print(f"Erro ao processar o número {process_number}: {e}")

# Iterar sobre os números de processo
for process in process_numbers:
    download_process(process)

# Fechar o navegador
driver.quit()
