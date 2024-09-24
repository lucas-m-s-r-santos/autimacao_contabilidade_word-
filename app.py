from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep 
from docx import Document
import os
# 1 - Entrar no site 

driver =  webdriver.Chrome()  
driver.get('https://contabil-devaprender.netlify.app')

# 2 - preencher email, senha e entrar.
campo_email = driver.find_element(By.XPATH,"//input[@type='email']")
sleep(2)
campo_email.click()
campo_email.send_keys('teste@gmail.com')
sleep(2)
# Campo senha 
campo_senha = driver.find_element(By.XPATH,"//input[@type='password']")
sleep(2)
campo_senha.click()
campo_senha.send_keys('123456789@2024')
sleep(2)
# campo enter
campo_enter = driver.find_element(By.XPATH,"//button[@type='submit']")
campo_enter.click()
sleep(3)
# 3 - acessar o "cadastrar o balanço patrimonial"
botoes_sistema = driver.find_elements(By.XPATH,"//a[@class='btn btn-primary mt-auto']")
sleep(1)
botoes_sistema[0].click()

def inserir_valores_de_documento_word(caminho_arquivo_word):
    
    # 4 -  extrair os dados do world
    doc = Document(caminho_arquivo_word)

    ativo_circulante = ''
    caixa_equivalente = ''
    contas_receber =  ''
    estoque = ''
    ativo_nao_circulante = ''
    imobilizado = ''
    intangivel = ''
    total_ativo = ''

    for t in doc.tables:
        for l in t.rows:
            if 'Ativo Circulante' in l.cells[0].text.strip():
                ativo_circulante = l.cells[1].text.strip()
        
            elif 'Caixa e Equivalentes' in l.cells[0].text.strip():
                caixa_equivalente = l.cells[1].text.strip()
            
            elif 'Contas a Receber' in l.cells[0].text.strip():
                contas_receber = l.cells[1].text.strip()
                
            elif 'Estoques' in l.cells[0].text.strip():
                estoque = l.cells[1].text.strip()
                
            elif 'Ativo Não Circulante' in l.cells[0].text.strip():
                ativo_nao_circulante = l.cells[1].text.strip()
            
            elif 'Imobilizado' in l.cells[0].text.strip():
                imobilizado = l.cells[1].text.strip()
            
            elif 'Intangível' in l.cells[0].text.strip():
                intangivel  = l.cells[1].text.strip()
                
            elif'Total do Ativo' in l.cells[0].text.strip():
                total_ativo = l.cells[1].text.strip()

    # 5 -  preencher os campos do arquivo world            
    primeira_barra = driver.find_element(By.ID, 'ativo_circulante')
    sleep(2)
    primeira_barra.click()
    primeira_barra.send_keys(ativo_circulante)

    segunda_barra = driver.find_element(By.ID, 'caixa_equivalentes')
    sleep(3)
    segunda_barra.click()
    segunda_barra.send_keys(caixa_equivalente)

    terceira_barra = driver.find_element(By.ID, 'contas_receber')
    sleep(3)
    terceira_barra.click()
    terceira_barra.send_keys(contas_receber)

    quarta_barra = driver.find_element(By.ID, 'estoques')
    sleep(3)
    quarta_barra.click()
    quarta_barra.send_keys(estoque)

    quinta_barra = driver.find_element(By.ID, 'ativo_nao_circulante')
    sleep(3)
    quinta_barra.click()
    quinta_barra.send_keys(ativo_nao_circulante)

    sexta_barra = driver.find_element(By.ID, 'imobilizado')
    sleep(3)
    sexta_barra.click()
    sexta_barra.send_keys(imobilizado)

    setiema_barra = driver.find_element(By.ID,'intangivel')
    sleep(3)
    setiema_barra.click()
    setiema_barra.send_keys(intangivel)

    oitava_barra = driver.find_element(By.ID,'total_ativo')
    sleep(3)
    oitava_barra.click()
    oitava_barra.send_keys(total_ativo)

    cadastrar = driver.find_element(By.XPATH,"//button[@class='btn btn-primary']")
    sleep(3)
    cadastrar.click()

# 6 -  repito os passos 4 e 5 até chegar ao final dos arquivos!
pasta_relatorios = r'C:\Users\Micro\Desktop\freelancer_contabilidade\relatorios'
for nome_arquivo in os.listdir(pasta_relatorios):
    if nome_arquivo.endswith('.docx'):
        caminho_arquivo_word = os.path.join(pasta_relatorios,nome_arquivo)
        inserir_valores_de_documento_word(caminho_arquivo_word)