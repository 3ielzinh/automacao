from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

# Configurar o navegador
driver = webdriver.Chrome()

# Abrir o site do Diário Oficial da União
driver.get('https://www.in.gov.br/leiturajornal?secao=do2')

# Aguarde até que o elemento com ID "viewMenuOptionTree" seja clicável
wait = WebDriverWait(driver, 10)
alteravisualizacao = wait.until(
    EC.element_to_be_clickable((By.ID, 'viewMenuOptionTree'))
)

# Clique no elemento para alterar a visualização
alteravisualizacao.click()

# Função para extrair links e texto de elementos
def extract_links_and_text(parent_element, class_name):
    elements = parent_element.find_elements(By.CLASS_NAME, class_name)
    result = []
    for element in elements:
        a_element = element.find_element(By.TAG_NAME, 'a')
        href = a_element.get_attribute('href')
        text = a_element.text
        result.append({"Text": text, "Link": href})
    return result

# Encontre os elementos para Casa Civil e Previdência Social
span_casa_civil = driver.find_element(By.XPATH, "//span[text()='Casa Civil']")
ul_casa_civil = span_casa_civil.find_element(By.XPATH, "./following-sibling::ul")
span_previdencia_social = driver.find_element(By.XPATH, "//span[text()='Ministério da Previdência Social']")
ul_previdencia_social = span_previdencia_social.find_element(By.XPATH, "./following-sibling::ul")

# Extraia os links e texto dos elementos
links_and_text_casa_civil = extract_links_and_text(ul_casa_civil, "file")
links_and_text_previdencia = extract_links_and_text(ul_previdencia_social, "file")

# Listas para armazenar detalhes e portarias
detalhes_casa_civil = []
portarias_casa_civil = []
detalhes_previdencia = []
portarias_previdencia = []

# Itere sobre os links e extraia detalhes e portarias
for item in links_and_text_casa_civil:
    link = item['Link']
    driver.get(link)
    detalhes_cc = driver.find_element(By.XPATH, "//div[@class='detalhes-dou']").text
    portaria_cc = driver.find_element(By.XPATH, "//div[@class='texto-dou']").text
    detalhes_casa_civil.append(detalhes_cc)
    portarias_casa_civil.append(portaria_cc)

for item in links_and_text_previdencia:
    link = item['Link']
    driver.get(link)
    detalhes_cc = driver.find_element(By.XPATH, "//div[@class='detalhes-dou']").text
    portaria_cc = driver.find_element(By.XPATH, "//div[@class='texto-dou']").text
    detalhes_previdencia.append(detalhes_cc)
    portarias_previdencia.append(portaria_cc)

# Feche o navegador
driver.quit()

# Crie DataFrames com as informações
df_casa_civil = pd.DataFrame({"Link": [item['Link'] for item in links_and_text_casa_civil],
                              "Detalhes": detalhes_casa_civil,
                              "Portaria": portarias_casa_civil})
df_previdencia = pd.DataFrame({"Link": [item['Link'] for item in links_and_text_previdencia],
                              "Detalhes": detalhes_previdencia,
                              "Portaria": portarias_previdencia})

# Salve os DataFrames em um arquivo Excel
with pd.ExcelWriter('C:\\Users\\Gabriel\\Desktop\\base.xlsx', engine='xlsxwriter') as writer:
    df_casa_civil.to_excel(writer, sheet_name='Casa Civil', index=False)
    df_previdencia.to_excel(writer, sheet_name='Previdência Social', index=False)
