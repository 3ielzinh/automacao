import pandas as pd
import re
import pdfkit
import time

# Função para remover caracteres proibidos do texto
def remover_caracteres_proibidos(texto):
    # Defina uma expressão regular que corresponda a caracteres proibidos (por exemplo, caracteres especiais)
    padrao = r'[^a-zA-Z0-9\s]'  # Esta expressão regular remove tudo que não é letra, número ou espaço em branco
    # Use re.sub para substituir os caracteres proibidos por uma string vazia
    texto_limpo = re.sub(padrao, ' ', texto)
    return texto_limpo

# Carregue o arquivo Excel (substitua 'seu_arquivo.xlsx' pelo caminho do seu arquivo)
df = pd.read_excel('C:\\Users\\Gabriel\\Desktop\\base.xlsx')

# Defina o número de segundos a aguardar entre cada conversão
tempo_de_espera = 1  # Tempo de espera em segundos

# Percorra as linhas da planilha
for index, row in df.iterrows():
    categoria = row['Categoria']

    # Verifique se a coluna "Categoria" contém uma informação
    if not pd.isna(categoria):
        link = row['Link']
        text = row['Text']
        # Remova caracteres proibidos do texto
        text_limpo = remover_caracteres_proibidos(text)
        pdf_filename = f'{text_limpo}.pdf'
        config = pdfkit.configuration(wkhtmltopdf=r"C:\Users\Gabriel\Desktop\HTMLPDF\wkhtmltopdf\bin\wkhtmltopdf.exe")
        options = {
            'print-media-type': None,
            'no-images': None
        }
        pdfkit.from_url(link, pdf_filename, configuration=config, options=options)
        # Aguarde o tempo definido antes de continuar com a próxima conversão
        time.sleep(tempo_de_espera)

input("")
