from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd

navegador = webdriver.Chrome()
#Passo 1: pegar a cotação do Dólar
navegador.get('https://www.google.com.br/')
navegador.find_element(By.XPATH,'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação dolar')
navegador.find_element(By.XPATH,'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotaçao_dolar = navegador.find_element(By.XPATH,'//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotaçao_dolar)

#Passo 2: pegar a cotação do Euro
navegador.get('https://www.google.com.br/')
navegador.find_element(By.XPATH,'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação euro')
navegador.find_element(By.XPATH,'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotaçao_euro = navegador.find_element(By.XPATH,'//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print(cotaçao_euro)

#Passo 3: pegar a cotação do Ouro
navegador.get('https://www.melhorcambio.com/ouro-hoje')

cotaçao_ouro = navegador.find_element(By.XPATH,'//*[@id="comercial"]').get_attribute('value')
cotaçao_ouro = cotaçao_ouro.replace(',','.')
print(cotaçao_ouro)

#Passo 4: importar a base de dados e atualizar as cotações nela
tabela = pd.read_excel('Produtos.xlsx')

#Passo 5: Calcular os novos preços e salvar/exportar a base de dados
tabela.loc[tabela['Moeda'] == 'Dólar','Cotação'] = float(cotaçao_dolar)
tabela.loc[tabela['Moeda'] == 'Euro','Cotação'] = float(cotaçao_euro)
tabela.loc[tabela['Moeda'] == 'Ouro','Cotação'] = float(cotaçao_ouro)

tabela['Preço de Compra'] = tabela['Preço Original'] * tabela['Cotação']

tabela['Preço de Venda'] = tabela['Preço de Compra'] * tabela['Margem']

tabela.to_excel('Produtos Novo.xlsx', index=False)
print(tabela)
