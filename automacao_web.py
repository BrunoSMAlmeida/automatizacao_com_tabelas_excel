"""Automação Web e Busca de Informações com Python
Trabalhamos em uma importadora e o preço dos nossos produtos é vinculado a cotação de: Dólar, Euro e Ouro
Precisamos pegar na internet, de forma automática, a cotação desses 3 itens e saber quanto devemos
cobrar pelos nossos produtos,considerando uma margem de contribuição que temos na nossa base de dados.
Base de Dados: https://drive.google.com/drive/folders/1KmAdo593nD8J9QBaZxPOG1yxHZua4Rtv?usp=sharing
Para isso, vamos criar uma automação web: Usaremos o Selenium."""
import pandas
from selenium import webdriver # permite criar o navegador
from selenium.webdriver.common.keys import Keys # permite escrever no navegador
from selenium.webdriver.common.by import By # permite selecionar itens no navegador
import pandas as pd
import numpy
import openpyxl


"""Passo 1: Pegar a cotação do Dólar:"""
# Abrir o navegador
navegador = webdriver.Chrome()
navegador.maximize_window()

# Acessar o google e digitar no google: cotação dólar
navegador.get("https://www.google.com/")
navegador.find_element("xpath",
                       "/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input").send_keys("cotação dólar")
navegador.find_element("xpath",
                       '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[4]/center/input[1]').click()

# Pegar o valor do dólar que o google informa
cotacao_dolar = navegador.find_element("xpath",
                                       '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")


"""Passo 2: Pegar a cotação do Euro:"""
navegador.get("https://www.google.com/")
navegador.find_element("xpath",
                       "/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input").send_keys("cotação euro")
navegador.find_element("xpath",
                       '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[4]/center/input[1]').click()

# Pegar o valor do Euro que o google informa
cotacao_euro = navegador.find_element("xpath",
                                       '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")


"""Passo 3: Pegar a cotação do Ouro:"""
# Acessar o site que informa a cotação do ouro
navegador.get("https://www.melhorcambio.com/ouro-hoje")
# Pegar o valor do Ouro que o site informa
cotacao_ouro = navegador.find_element("xpath", '//*[@id="comercial"]').get_attribute("value")
cotacao_ouro = cotacao_ouro.replace("," , ".")

# Para fechar o navegador assim que finalizar o passo
navegador.quit()


"""Passo 4: Atualizar a base de preços (atualizando o preço de compra e o de venda):"""
tabela = pandas.read_excel(r"C:\Users\Objectedge\Downloads\Produtos.xlsx")


# Atualizar a coluna de cotação

# Editar a coluna de cotação, onde a coluna moeda for igual a moeda que quero editar.
tabela.loc[tabela["Moeda"] == "Dólar", "Cotação"] = float(cotacao_dolar)
tabela.loc[tabela["Moeda"] == "Euro", "Cotação"] = float(cotacao_euro)
tabela.loc[tabela["Moeda"] == "Ouro", "Cotação"] = float(cotacao_ouro)


# Atualizar a coluna preço de compra = preço original x cotação
tabela["Preço de Compra"] = tabela["Preço Original"] * tabela["Cotação"]
# Atualizar a coluna preço de venda = preço de compra x a margem
tabela["Preço de Venda"] = tabela["Preço de Compra"] * tabela["Margem"]


"""Passo 5: Exportar a base de preços atualizado:"""
# Sempre mude o nome da nova base de dados para não perder a antiga.
tabela.to_excel(r"C:\Users\Objectedge\Downloads\Produtos Novos.xlsx", index=False)

