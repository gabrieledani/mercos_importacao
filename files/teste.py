from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.action_chains import ActionChains        

import time
import pandas as pd

opcoes = webdriver.ChromeOptions()
servico = Service(ChromeDriverManager().install())
opcoes.add_experimental_option("excludeSwitches", ["enable-logging"])
opcoes.add_experimental_option("detach", True)
navegador = webdriver.Chrome(service=servico,options=opcoes)

wait = WebDriverWait(navegador, 150)

#element_to_be_clickable
#presence_of_element_located
#visibility_of_element_located

#abre o sistema
navegador.get("https://app.mercos.com/")
#navegador.maximize_window()
time.sleep(5)

action = ActionChains(navegador)

'''
# send keys
action.send_keys("Arrays")
# perform the operation
action.perform()        
'''

#autenticacao
usuario = wait.until(ec.presence_of_element_located((By.ID,'id_usuario')))
usuario.send_keys('gabrieledani@gmail.com')
senha = wait.until(ec.presence_of_element_located((By.ID,'id_senha')))
senha.send_keys('Vedafil2022')

lg = wait.until(ec.element_to_be_clickable((By.ID,"botaoEfetuarLogin")))
navegador.execute_script("arguments[0].scrollIntoView();", lg)
lg.click()

vlr_frete = 'CIF (Frete Pago)'
vlr_cond = '28/42/56'
vlr_repr = 'Hengst Indústria de Filtros Ltda'

df_pedidos = pd.read_excel('banco_hengst_2015.xlsx')

#file = open('resultado.txt','w')

for pedido in df_pedidos.itertuples(name='pedidos',index=False):
    print('Pedido->',pedido)
    #file.write('Pedido->'+str(pedido))
    cnpj = str(pedido.CNPJ)
    cnpj = cnpj.replace('.','')
    cnpj = cnpj.replace('/','')
    cnpj = cnpj.replace('-','')
    
    if pedido.acao == 1:
        print('primeiro produto do pedido de vários')
        #file.write('primeiro produto do pedido de vários\n')

        print('aba pedidos')
        #file.write('aba pedidos\n')
        aba = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="aba_pedidos"]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", aba)
        aba.click()
        time.sleep(1)

        print('botão pedidos')
        #file.write('botão pedidos\n')
        criar = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="btn_criar_pedido"]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", criar)
        criar.click()
        time.sleep(1)
        
        print('cliente',cnpj)
        #file.write('cliente'+str(cnpj)+'\n')
        cliente = wait.until(ec.presence_of_element_located((By.XPATH,'//input[@id="id_codigo_cliente"]')))
        cliente.send_keys(cnpj)
        #action.send_keys(cnpj).perform()
        time.sleep(1)
        cli_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_campo_id_codigo_cliente"]/ul/li[1]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", cli_pro)
        cli_pro.click()
        time.sleep(1)

        print('representada',vlr_repr)
        #file.write('representada'+vlr_repr+'\n')
        representada = wait.until(ec.presence_of_element_located((By.XPATH,'//input[@id="id_codigo_representada"]')))
        representada.send_keys(vlr_repr)
        #action.send_keys(vlr_repr).perform()
        time.sleep(1)
        rep_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_campo_id_codigo_representada"]/ul/li[1]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", rep_pro)
        rep_pro.click()
        time.sleep(1)

        print('produto',pedido.produto)
        #file.write('produto'+pedido.produto+'\n')
        produto = wait.until(ec.presence_of_element_located((By.XPATH,'//input[@id="produto_autocomplete"]')))
        produto.send_keys(pedido.produto)
        #action.send_keys(pedido.produto).perform()
        time.sleep(1)
        adc_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        time.sleep(1)
        #avegador.execute_script("arguments[0].scrollIntoView();", adc_pro)
        adc_pro.click()
        time.sleep(3)

        print('quantidade',pedido.quantidade)
        #file.write('quantidade'+str(pedido.quantidade)+'\n')
        quantidade = wait.until(ec.presence_of_element_located((By.XPATH,'//input[@id="id_quantidade"]')))
        quantidade.send_keys(pedido.quantidade)
        time.sleep(1)

        print('preço antes',pedido.valor)
        #file.write('preço antes'+str(pedido.valor)+'\n')
        valor = wait.until(ec.element_to_be_clickable((By.XPATH,'//input[@id="id_preco_final"]')))
        valor.click()
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        print('preço formatado',preco)
        #file.write('preço formatado'+preco+'\n')
        valor.send_keys(preco)
        time.sleep(1)
        #action.send_keys(preco).perform()
        valor.send_keys(Keys.TAB)
        time.sleep(1)

        print('info_ad')
        #file.write('info_ad\n')
        info_ad = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_informacoes_adicionais"]')))
        info_ad.send_keys('Importado via Excel')
        time.sleep(1)

        print('salva produto')
        #file.write('salva produto\n')
        slv_pro =  wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="adicao_produto"]/form/div[3]/a[1]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", slv_pro)
        slv_pro.click()
        time.sleep(5)

    elif pedido.acao == 0:
        print('continua informando produtos',pedido.produto)
        #file.write('continua informando produtos '+pedido.produto+'\n')
        
        print('produto',pedido.produto)
        #file.write('produto'+pedido.produto+'\n')
        produto = wait.until(ec.presence_of_element_located((By.XPATH,'//input[@id="produto_autocomplete"]')))
        produto.send_keys(pedido.produto)
        #action.send_keys(pedido.produto).perform()
        time.sleep(1)
        adc_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        time.sleep(1)
        #avegador.execute_script("arguments[0].scrollIntoView();", adc_pro)
        adc_pro.click()
        time.sleep(3)

        print('quantidade',pedido.quantidade)
        #file.write('quantidade'+str(pedido.quantidade)+'\n')
        quantidade = wait.until(ec.presence_of_element_located((By.XPATH,'//input[@id="id_quantidade"]')))
        quantidade.send_keys(pedido.quantidade)
        time.sleep(1)

        print('preço antes',pedido.valor)
        #file.write('preço antes'+str(pedido.valor)+'\n')
        valor = wait.until(ec.element_to_be_clickable((By.XPATH,'//input[@id="id_preco_final"]')))
        valor.click()
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        print('preço formatado',preco)
        #file.write('preço formatado'+preco+'\n')
        valor.send_keys(preco)
        time.sleep(1)
        #action.send_keys(preco).perform()
        valor.send_keys(Keys.TAB)
        time.sleep(1)

        print('info_ad')
        #file.write('info_ad\n')
        info_ad = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_informacoes_adicionais"]')))
        info_ad.send_keys('Importado via Excel')
        time.sleep(1)

        print('salva produto')
        #file.write('salva produto\n')
        slv_pro =  wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="adicao_produto"]/form/div[3]/a[1]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", slv_pro)
        slv_pro.click()
        time.sleep(5)

    elif pedido.acao == 3:
        print('ultimo produto',pedido.produto)
        #file.write('ultimo produto '+pedido.produto+'\n')
        
        print('produto',pedido.produto)
        #file.write('produto'+pedido.produto+'\n')
        produto = wait.until(ec.presence_of_element_located((By.XPATH,'//input[@id="produto_autocomplete"]')))
        produto.send_keys(pedido.produto)
        #action.send_keys(pedido.produto).perform()
        time.sleep(1)
        adc_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        time.sleep(1)
        #avegador.execute_script("arguments[0].scrollIntoView();", adc_pro)
        adc_pro.click()
        time.sleep(3)

        print('quantidade',pedido.quantidade)
        #file.write('quantidade'+str(pedido.quantidade)+'\n')
        quantidade = wait.until(ec.presence_of_element_located((By.XPATH,'//input[@id="id_quantidade"]')))
        quantidade.send_keys(pedido.quantidade)
        time.sleep(1)

        print('preço antes',pedido.valor)
        #file.write('preço antes'+str(pedido.valor)+'\n')
        valor = wait.until(ec.element_to_be_clickable((By.XPATH,'//input[@id="id_preco_final"]')))
        valor.click()
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        print('preço formatado',preco)
        #file.write('preço formatado'+preco+'\n')
        valor.send_keys(preco)
        time.sleep(1)
        #action.send_keys(preco).perform()
        valor.send_keys(Keys.TAB)
        time.sleep(1)

        print('info_ad')
        #file.write('info_ad\n')
        info_ad = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_informacoes_adicionais"]')))
        info_ad.send_keys('Importado via Excel')
        time.sleep(1)

        print('salva produto')
        #file.write('salva produto\n')
        slv_pro =  wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="adicao_produto"]/form/div[3]/a[1]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", slv_pro)
        slv_pro.click()
        time.sleep(5)
        
        print('terminei produtos')
        #file.write('terminei produtos\n')
        terminei_prod = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="botao_terminei_de_adicionar"]')))
        navegador.execute_script("arguments[0].scrollIntoView();", terminei_prod)
        navegador.execute_script("arguments[0].click();", terminei_prod)
        time.sleep(3)

        print('vendedor',pedido.vendedor)
        #file.write('vendedor'+pedido.vendedor+'\n')
        vendedor = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_criador"]')))
        vendedor.send_keys(pedido.vendedor)
        time.sleep(1)

        print('pagamento',vlr_cond)
        #file.write('pagamento'+vlr_cond+'\n')
        cond_pgto = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_cond_pagamento"]')))
        cond_pgto.send_keys(vlr_cond)
        #//*[@id="js-ultima-condicao-pagamento-utilizada"]
        time.sleep(1)

        print('frete',vlr_frete)
        #file.write('frete'+vlr_frete+'\n')
        frete = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_transportadora"]')))
        frete.send_keys(vlr_frete)
        #//*[@id="js-link-ultima-transportadora"]
        time.sleep(1)
        
        print('salva pedido')
        #file.write('salva pedido\n')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="simplest_modal"]/div[2]/form/div[2]/a[1]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", slv_pdv)
        #navegador.execute_script("arguments[0].click();", slv_pdv)
        slv_pdv.click()
        time.sleep(1)

        print('GERAR PEDIDO...')
        #file.write('GERAR PEDIDO...\n')
        gera_pedido = wait.until(ec.element_to_be_clickable((By.LINK_TEXT,'Transformar em pedido')))
        #gera_pedido = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="gritter-item-1"]/div[2]/div[2]/div/a[1]')))
        navegador.execute_script("arguments[0].scrollIntoView();", gera_pedido)
        #navegador.execute_script("arguments[0].click();", gera_pedido)
        gera_pedido.click()
        time.sleep(3)
        
        print('alterar pedido')
        #file.write('alterar pedido\n')
        alterar = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="alterar_informacoes"]')))
        navegador.execute_script("arguments[0].scrollIntoView();", alterar)
        #navegador.execute_script("arguments[0].click();", alterar)
        alterar.click()
        time.sleep(4)

        print('data emissao')
        #file.write('data emissao\n')
        data_emis = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_data_emissao"]')))
        data_emis.clear()
        data_emis.send_keys(pedido.data.strftime("%d/%m/%Y"))
        time.sleep(1)

        print('salva pedido')
        #file.write('salva pedido\n')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="simplest_modal"]/div[2]/form/div[2]/a[1]')))
        navegador.execute_script("arguments[0].scrollIntoView();", slv_pdv)
        #navegador.execute_script("arguments[0].click();", slv_pdv)
        slv_pdv.click()
        time.sleep(3)
    
    elif pedido.acao == 4:
        print('informa cliente e representada e primeiro produto')
        #file.write('informa cliente e representada e primeiro produto\n')

        print('aba pedidos')
        #file.write('aba pedidos\n')
        aba = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="aba_pedidos"]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", aba)
        aba.click()
        time.sleep(1)

        print('botão pedidos')
        #file.write('botão pedidos\n')
        criar = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="btn_criar_pedido"]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", criar)
        criar.click()
        time.sleep(1)
        
        print('cliente',cnpj)
        #file.write('cliente'+str(cnpj)+'\n')
        action.send_keys(cnpj).perform()
        time.sleep(1)
        cli_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_campo_id_codigo_cliente"]/ul/li[1]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", cli_pro)
        cli_pro.click()
        time.sleep(1)

        print('representada',vlr_repr)
        #file.write('representada'+vlr_repr+'\n')
        action.send_keys(vlr_repr).perform()
        time.sleep(1)
        rep_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_campo_id_codigo_representada"]/ul/li[1]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", rep_pro)
        rep_pro.click()
        time.sleep(1)

        print('produto',pedido.produto)
        #file.write('produto'+pedido.produto+'\n')
        action.send_keys(pedido.produto).perform()
        time.sleep(1)
        adc_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        #avegador.execute_script("arguments[0].scrollIntoView();", adc_pro)
        adc_pro.click()
        time.sleep(3)

        print('quantidade',pedido.quantidade)
        #file.write('quantidade'+str(pedido.quantidade)+'\n')
        quantidade = wait.until(ec.presence_of_element_located((By.XPATH,'//input[@id="id_quantidade"]')))
        quantidade.send_keys(pedido.quantidade)
        time.sleep(1)

        print('preço antes',pedido.valor)
        #file.write('preço antes'+str(pedido.valor)+'\n')
        valor = wait.until(ec.element_to_be_clickable((By.XPATH,'//input[@id="id_preco_final"]')))
        valor.click()
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        print('preço formatado',preco)
        #file.write('preço formatado'+preco+'\n')
        valor.send_keys(preco)
        time.sleep(1)
        #action.send_keys(preco).perform()
        valor.send_keys(Keys.TAB)
        time.sleep(1)

        print('info_ad')
        #file.write('info_ad\n')
        info_ad = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_informacoes_adicionais"]')))
        info_ad.send_keys('Importado via Excel')
        time.sleep(1)

        print('salva produto')
        #file.write('salva produto\n')
        slv_pro =  wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="adicao_produto"]/form/div[3]/a[1]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", slv_pro)
        slv_pro.click()
        time.sleep(3)
        
        print('terminei produtos')
        #file.write('terminei produtos\n')
        terminei_prod = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="botao_terminei_de_adicionar"]')))
        navegador.execute_script("arguments[0].scrollIntoView();", terminei_prod)
        navegador.execute_script("arguments[0].click();", terminei_prod)
        time.sleep(3)

        print('vendedor',pedido.vendedor)
        #file.write('vendedor'+pedido.vendedor+'\n')
        vendedor = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_criador"]')))
        vendedor.send_keys(pedido.vendedor)
        time.sleep(1)

        print('pagamento',vlr_cond)
        #file.write('pagamento'+vlr_cond+'\n')
        cond_pgto = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_cond_pagamento"]')))
        cond_pgto.send_keys(vlr_cond)
        #//*[@id="js-ultima-condicao-pagamento-utilizada"]
        time.sleep(1)

        print('frete',vlr_frete)
        #file.write('frete'+vlr_frete+'\n')
        frete = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_transportadora"]')))
        frete.send_keys(vlr_frete)
        #//*[@id="js-link-ultima-transportadora"]
        time.sleep(1)
        
        print('salva pedido')
        #file.write('salva pedido\n')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="simplest_modal"]/div[2]/form/div[2]/a[1]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", slv_pdv)
        #navegador.execute_script("arguments[0].click();", slv_pdv)
        slv_pdv.click()
        time.sleep(1)

        print('GERAR PEDIDO...')
        #file.write('GERAR PEDIDO...\n')
        gera_pedido = wait.until(ec.element_to_be_clickable((By.LINK_TEXT,'Transformar em pedido')))
        #gera_pedido = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="gritter-item-1"]/div[2]/div[2]/div/a[1]')))
        navegador.execute_script("arguments[0].scrollIntoView();", gera_pedido)
        #navegador.execute_script("arguments[0].click();", gera_pedido)
        gera_pedido.click()
        time.sleep(3)
        
        print('alterar pedido')
        #file.write('alterar pedido\n')
        alterar = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="alterar_informacoes"]')))
        navegador.execute_script("arguments[0].scrollIntoView();", alterar)
        #navegador.execute_script("arguments[0].click();", alterar)
        alterar.click()
        time.sleep(4)

        print('data emissao')
        #file.write('data emissao\n')
        data_emis = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_data_emissao"]')))
        data_emis.clear()
        data_emis.send_keys(pedido.data.strftime("%d/%m/%Y"))
        time.sleep(1)

        print('salva pedido')
        #file.write('salva pedido\n')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="simplest_modal"]/div[2]/form/div[2]/a[1]')))
        navegador.execute_script("arguments[0].scrollIntoView();", slv_pdv)
        #navegador.execute_script("arguments[0].click();", slv_pdv)
        slv_pdv.click()
        time.sleep(3)

time.sleep(20)
#file.close()