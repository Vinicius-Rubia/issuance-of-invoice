from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import os
import time
import pandas as pd

options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {
  "download.default_directory": r"C:\Users\55119\Documents\NFs",
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})

servico = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=servico, options=options)
driver.maximize_window()

caminho = os.getcwd()
arquivo = caminho + r"\login.html"
driver.get(arquivo)

tabela = pd.read_excel('NotasEmitir.xlsx')

driver.find_element(By.XPATH, '/html/body/div/form/input[1]').send_keys('Vinicius@gmail.com')
driver.find_element(By.XPATH, '/html/body/div/form/input[2]').send_keys('123456')
driver.find_element(By.XPATH, '/html/body/div/form/button').click()

for linha in tabela.index:
    
    driver.find_element(By.NAME, 'nome').send_keys(tabela.loc[linha, 'Cliente'])
    driver.find_element(By.NAME, 'endereco').send_keys(tabela.loc[linha, 'Endereço'])
    driver.find_element(By.NAME, 'bairro').send_keys(tabela.loc[linha, 'Bairro'])
    driver.find_element(By.NAME, 'municipio').send_keys(tabela.loc[linha, 'Municipio'])
    driver.find_element(By.NAME, 'cep').send_keys(str(tabela.loc[linha, 'CEP']))
    driver.find_element(By.NAME, 'uf').send_keys(tabela.loc[linha, 'UF'])
    driver.find_element(By.NAME, 'cnpj').send_keys(str(tabela.loc[linha, 'CPF/CNPJ']))
    driver.find_element(By.NAME, 'inscricao').send_keys(str(tabela.loc[linha, 'Inscricao Estadual']))
    driver.find_element(By.NAME, 'descricao').send_keys(tabela.loc[linha, 'Descrição'])
    driver.find_element(By.NAME, 'quantidade').send_keys(str(tabela.loc[linha, 'Quantidade']))
    driver.find_element(By.NAME, 'valor_unitario').send_keys(str(tabela.loc[linha, 'Valor Unitario']))
    driver.find_element(By.NAME, 'total').send_keys(str(tabela.loc[linha, 'Valor Total']))
    driver.find_element(By.CLASS_NAME, 'registerbtn').click()
    driver.refresh()

time.sleep(3)