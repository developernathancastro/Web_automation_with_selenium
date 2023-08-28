#criar um navegador
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import win32com.client as win32

navegador = webdriver.Chrome()

####importar base de dados/visualizar

import pandas as pd

tabela_produtos = pd.read_excel(r'C:\Users\natha\Projeto 2 - Automação Web - Aplicação de Mercado de Trabalho\buscas.xlsx')
print(tabela_produtos)

def verificar_tem_termos_banidos(lista_termos_banidos, nome_texto):
    tem_termos_banidos = False
    for palavra in lista_termos_banidos:
        if palavra in nome_texto:
            tem_termos_banidos = True
    return tem_termos_banidos

def verificar_tem_todos_termos_produto(lista_termos_nome_produto, nome_texto):
    tem_todos_os_termos = True
    for nomes_produto in lista_termos_nome_produto:
        if nomes_produto not in nome_texto:
            tem_todos_os_termos = False
    return tem_todos_os_termos

navegador.get('https://www.google.com/')

def busca_google_shopping(navegador,produto, termos_banidos, preco_minimo, preco_maximo):
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_nome_produto = produto.split(" ")

    lista_ofertas = []
    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)

    navegador.get('https://www.google.com/')
    pesquisar = navegador.find_element(By.ID, 'APjFqb').send_keys(produto, Keys.ENTER)

    ##entrar na aba shopping
    elementos  = navegador.find_elements(By.CLASS_NAME, 'hdtb-mitem')

    for item in elementos:
        if 'Shopping' in item.text:
            item.click()
            break

##pegar as informações do produto
    lista_resultados = navegador.find_elements(By.CLASS_NAME, 'KZmu8e')

    for resultado in lista_resultados:
        elemento_referencia1 =  resultado.find_element(By.CLASS_NAME, 'hn9kf')
        elemento_pai1 = elemento_referencia1.find_element(By.XPATH, '..')
        nome = elemento_pai1.text
        nome_texto = nome[:15]
        nome_texto = nome_texto.lower()

    ##analisar se ele não tem nehum termo banido
        termos_banidos = verificar_tem_termos_banidos(lista_termos_banidos, nome_texto)

    ##se ele tem todos os termos do nome do produto
        tem_todos_os_termos = verificar_tem_todos_termos_produto(lista_termos_nome_produto, nome_texto)

    ##selecionar somente os elementos que tem termos banidos = False eao mesmo tempo tem todos os termos_produto = True
        #try:
        if termos_banidos ==  False and tem_todos_os_termos == True:
            preco = resultado.find_element(By.CLASS_NAME, 'T14wmb').text
            preco_texto = preco[:11]
            preco_texto = preco_texto.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
            preco_texto = float(preco_texto)

        #except:
            #continue

        ##se o preco está entre o preço mínimo e o preço máximo:
            if preco_minimo <= preco_texto <= preco_maximo:
                elemento_referencia2 = resultado.find_element(By.CLASS_NAME, 'ROMz4c')
                elemento_pai2 = elemento_referencia2.find_element(By.XPATH, '..')
                link  = elemento_pai2.get_attribute('href')
                lista_ofertas.append((nome, preco, link))

    return lista_ofertas

##funcao do buscapé

def busca_buscape(navegador,produto, termos_banidos, preco_minimo, preco_maximo):

    ##tratar os valores
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_nome_produto = produto.split(" ")
    lista_ofertas = []
    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)
    ##buscar o produto no buscapé

    navegador.get('https://www.buscape.com.br/?og=17000&og=17000&msclkid=043d7324498010339c001761d9f491d'
                  '4&utm_source=bing&utm_medium=cpc&utm_campaign=BING%20BRAND%20-%20BP&utm_term=%5Bbuscape%5D&u'
                  'tm_content=BING%20BP%20-%20Buscap%C3%A9')

    pesquisar_buscape = navegador.find_element(By.XPATH, '//*[@id="new-header"]/div['
                                                         '1]/div/div/div[3]/di'
                                                         'v/div/div[2]/div/div[1]/input').send_keys(produto)

    efetivar_pesquisa = navegador.find_element(By.CLASS_NAME, 'AutoCompleteStyle_submitButton__GkxPO').click()

    while len(navegador.find_elements(By.CLASS_NAME, 'Select_Select__1S7HV')) < 1:
        time.sleep(1)
    ##pegar os resultados

    lista_resultados = navegador.find_elements(By.CLASS_NAME, 'Paper_Paper__HIHv0')

    for resultado in lista_resultados:
        nome = resultado.find_element(By.TAG_NAME, 'h2').text
        nome_texto = nome
        nome_texto = nome_texto.lower()

    ##analisar se o resultado tem termo banido e tem todos os termos do produto
        tem_termos_banidos = verificar_tem_termos_banidos(lista_termos_banidos, nome_texto)
        tem_todos_os_termos = verificar_tem_todos_termos_produto(lista_termos_nome_produto, nome_texto)
        #try:
        if tem_termos_banidos == False and tem_todos_os_termos == True:
            valor = resultado.find_element(By.CLASS_NAME, 'Text_MobileHeadingS__Zxam2').text
            preco_texto = valor
            preco_texto = preco_texto.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
            preco_texto = float(preco_texto)
        #except:
            #continue
    ##analisar se o preço está entre o preço minimo e preço maximi

            if preco_minimo<= preco_texto <= preco_maximo:
                link = resultado.find_element(By.CLASS_NAME, 'SearchCard_ProductCard_Inner__7JhKb').get_attribute(
                    'href')
                lista_ofertas.append((nome_texto, preco_texto, link))

    # retornar a lista de ofertas do buscapé
    return lista_ofertas


tabela_ofertas = pd.DataFrame()

for linha in tabela_produtos.index:
    produto = tabela_produtos.loc[linha, 'Nome']
    termos_banidos = tabela_produtos.loc[linha, 'Termos banidos']
    preco_minimo = tabela_produtos.loc[linha, 'Preço mínimo']
    preco_maximo = tabela_produtos.loc[linha, 'Preço máximo']

    lista_ofertas_google_shopping = busca_google_shopping(navegador, produto,termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_google_shopping:
        tabela_google_shopping = pd.DataFrame(lista_ofertas_google_shopping, columns= ['Produto', 'Preço', 'Link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_google_shopping])
    else: tabela_google_shopping = None

    lista_ofertas_buscape  = busca_buscape(navegador, produto, termos_banidos, preco_minimo,
                                                          preco_maximo)
    if lista_ofertas_buscape:
        tabela_buscape = pd.DataFrame(lista_ofertas_buscape, columns= ['Produto', 'Preço', 'Link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_buscape])

    else:tabela_buscape = None

print(tabela_ofertas)


#exportar para excel

tabela_ofertas.to_excel('Ofertas.xlsx', index= False)

#enviando e-mail
if len(tabela_ofertas) > 0:
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'nathan.castro2022@outlook.com'
    mail.Subject = 'Produto(s) Encontrado(s) na faixa de preço desejada'
    mail.body = f'''
    <p>Prezados,<p>
    <p>Encontramos alguns produtos em oferta dentro da faixa de preço desejada<p>
    {tabela_ofertas.to_html(index= False)}
    <p.Att., </p>
    '''
    mail.Send()












































##procurar esse produto no google shpping
##verificar se algum dos produtos do google shopping está dentro da minha faixa de preços









