from selenium import webdriver
import pandas as pd
from IPython.display import display
from selenium.webdriver.chrome.options import Options

# inicializa o programa sem aparecer na tela, em segundo plano
chrome_options = Options()
chrome_options.headless = True

# abre o navegador com as opções definidas acima
navegador = webdriver.Chrome(options=chrome_options)

# passo 1: pegar a cotação das moedas
# dolar
navegador.get('https://www.google.com/search?q=cotacao+dolar&oq=cotacao+dolar&aqs=chrome..69i57j35i39j0i512l3j0i433i512j0i512l4.1498j0j9&sourceid=chrome&ie=UTF-8')
cot_dolar = navegador.find_element('xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')

# euro
navegador.get('https://www.google.com/search?q=cotacao+euro&oq=cotacao+euro&aqs=chrome..69i57j0i512l4j0i10i512j0i512l4.1450j0j7&sourceid=chrome&ie=UTF-8')
cot_euro = navegador.find_element('xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')

# ouro
navegador.get('https://www.melhorcambio.com/ouro-hoje#:~:text=O%20valor%20do%20grama%20do,em%20R%24%20314%2C92.')
cot_ouro = navegador.find_element('xpath', '//*[@id="comercial"]').get_attribute('value')
cot_ouro = cot_ouro.replace(',', '.')

# fecha o navegador após pegar a cotação
navegador.quit()

# passo 2: importar e atualizar a base de dados
df = pd.read_excel('Produtos.xlsx')

# atualizar a cotação das moedas no dataframe
df.loc[df['Moeda'] == 'Dólar', 'Cotação'] = float(cot_dolar)    # loc[linha, coluna]
df.loc[df['Moeda'] == 'Euro', 'Cotação'] = float(cot_euro)
df.loc[df['Moeda'] == 'Ouro', 'Cotação'] = float(cot_ouro)

# atualizar o preço de compra e preço de venda
df['Preço de Compra'] = df['Preço Original'] * df['Cotação']
df['Preço de Venda'] = df['Preço de Compra'] * df['Margem']

# formatar os preços
df['Preço de Compra'], df['Preço de Venda'] = df['Preço de Compra'].map('R${:.2f}'.format), df['Preço de Venda'].map('R${:.2f}'.format)
df['Cotação'] = df['Cotação'].map('R${:.2f}'.format)

# passo 3: exportar a base de dados
df.to_excel('Produtos Novo.xlsx', index=False)   # nome do arquivo a ser criado | index=false para não exportar o index junto

df_novo = pd.read_excel('Produtos Novo.xlsx')
display(df_novo)