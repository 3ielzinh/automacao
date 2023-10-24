from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

import pandas as pd

def setup_webdriver():
    driver = webdriver.Chrome()
    driver.get('https://in.gov.br/leiturajornal?secao=do2')
    return driver

def extract_links_and_text(parent_element, class_name):
    elements = parent_element.find_elements(By.CLASS_NAME, class_name)
    return [{"Text": el.find_element(By.TAG_NAME, 'a').text, "Link": el.find_element(By.TAG_NAME, 'a').get_attribute('href')} for el in elements]

def refresh_page(driver):
    # Aguarde até que o botão de recarregamento esteja clicável ou até atingir o tempo limite de 6 segundos
    wait = WebDriverWait(driver, 6)
    try:
        wait.until(EC.element_to_be_clickable((By.id, 'reloadButton'))).click()  # Substitua 'reloadButton' pelo ID correto do botão de recarregamento
    except Exception as e:
        print(f"Erro ao recarregar a página: {str(e)}")

def extract_details_and_portarias(driver, links_and_text, detalhes_list, portarias_list):
    for item in links_and_text:
        try:
            driver.get(item['Link'])

            try:
                # Espera até que o elemento de detalhes seja visível
                detalhes_element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.CLASS_NAME, 'detalhes-dou')))
                detalhes = detalhes_element.text

                # Espera até que o elemento de portaria seja visível
                portaria_element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.CLASS_NAME, 'texto-dou')))
                portaria = portaria_element.text

                detalhes_list.append(detalhes)
                portarias_list.append(portaria)
            except TimeoutException:
                print(f"Tempo de espera excedido. Recarregando a página...")
                driver.refresh()
        except Exception as e:
            print(f"Erro ao extrair detalhes e portarias: {str(e)}")

with setup_webdriver() as driver:
    wait = WebDriverWait(driver, 10)
    alteravisualizacao = wait.until(EC.element_to_be_clickable((By.ID, 'viewMenuOptionTree')))
    alteravisualizacao.click()

    casa_civil, previdencia_social = driver.find_element(By.XPATH, "//span[text()='Casa Civil']"), driver.find_element(By.XPATH, "//span[text()='Ministério da Previdência Social']")
    ul_casa_civil, ul_previdencia_social = casa_civil.find_element(By.XPATH, "./following-sibling::ul"), previdencia_social.find_element(By.XPATH, "./following-sibling::ul")

    links_and_text_casa_civil, links_and_text_previdencia = extract_links_and_text(ul_casa_civil, "file"), extract_links_and_text(ul_previdencia_social, "file")

    detalhes_casa_civil, portarias_casa_civil, detalhes_previdencia, portarias_previdencia = [], [], [], []
    extract_details_and_portarias(driver, links_and_text_casa_civil, detalhes_casa_civil, portarias_casa_civil)
    extract_details_and_portarias(driver, links_and_text_previdencia, detalhes_previdencia, portarias_previdencia)

df_casa_civil = pd.DataFrame({"Text": [item['Text'] for item in links_and_text_casa_civil], "Link": [item['Link'] for item in links_and_text_casa_civil], "Detalhes": detalhes_casa_civil, "Portaria": portarias_casa_civil})
df_previdencia = pd.DataFrame({"Text": [item['Text'] for item in links_and_text_previdencia], "Link": [item['Link'] for item in links_and_text_previdencia], "Detalhes": detalhes_previdencia, "Portaria": portarias_previdencia})

df_banco_de_dados_final = pd.concat([df_casa_civil, df_previdencia], ignore_index=True)

excel_file_path = 'C:\\Users\\Gabriel\\Desktop\\base.xlsx'
try:
    with pd.ExcelFile(excel_file_path) as xls:
        df_banco_de_dados_existing = pd.read_excel(xls, 'BANCO DE DADOS')
except FileNotFoundError:
    df_banco_de_dados_existing = pd.DataFrame()

df_banco_de_dados_final.to_excel(excel_file_path, sheet_name='BANCO DE DADOS', index=False)