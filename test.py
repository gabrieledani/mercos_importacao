from selenium import webdriver
from selenium.webdriver.firefox.service import Service as FirefoxService
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.action_chains import ActionChains        

import time
import pandas as pd

servico = FirefoxService(executable_path=GeckoDriverManager().install())
navegador = webdriver.Firefox(service=servico)

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

df_pedidos = pd.read_excel('banco_hengst_2013_b.xlsx')

for pedido in df_pedidos.itertuples(name='pedidos',index=False):
    print('Pedido->',pedido)
    cnpj = str(pedido.CNPJ)
    cnpj = cnpj.replace('.','')
    cnpj = cnpj.replace('/','')
    cnpj = cnpj.replace('-','')
    
    if pedido.acao == 1:
        print('primeiro produto do pedido de vários')

        print('aba pedidos')
        aba = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="aba_pedidos"]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", aba)
        aba.click()
        time.sleep(1)

        print('botão pedidos')
        criar = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="btn_criar_pedido"]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", criar)
        criar.click()
        time.sleep(1)
        
        print('cliente')
        action.send_keys(cnpj).perform()
        time.sleep(1)
        cli_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_campo_id_codigo_cliente"]/ul/li[1]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", cli_pro)
        cli_pro.click()
        time.sleep(1)

        print('representada')
        action.send_keys(vlr_repr).perform()
        time.sleep(1)
        rep_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_campo_id_codigo_representada"]/ul/li[1]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", rep_pro)
        rep_pro.click()
        time.sleep(1)

        print('produto')
        action.send_keys(pedido.produto).perform()
        time.sleep(2)
        adc_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        #avegador.execute_script("arguments[0].scrollIntoView();", adc_pro)
        adc_pro.click()
        time.sleep(1)

        print('quantidade')
        quantidade = wait.until(ec.presence_of_element_located((By.XPATH,'//input[@id="id_quantidade"]')))
        quantidade.send_keys(pedido.quantidade)
        time.sleep(1)

        print('preço')
        valor = wait.until(ec.presence_of_element_located((By.XPATH,'//input[@id="id_preco_final"]')))
        valor.clear()
        time.sleep(1)
        preco = str(round(pedido.valor,10)).replace('.',',')
        valor.click()
        action.send_keys(preco).perform()
        #valor.send_keys(preco)
        time.sleep(1)

        print('info_ad')
        info_ad = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_informacoes_adicionais"]')))
        info_ad.send_keys('Importado via Excel')
        time.sleep(1)

        print('salva produto')
        slv_pro =  wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="adicao_produto"]/form/div[3]/a[1]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", slv_pro)
        slv_pro.click()
        time.sleep(2)

    elif pedido.acao == 0:
        print('continua informando produtos',pedido.produto)
            
        print('produto')
        action.send_keys(pedido.produto).perform()
        time.sleep(2)
        adc_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        #avegador.execute_script("arguments[0].scrollIntoView();", adc_pro)
        adc_pro.click()
        time.sleep(1)

        print('quantidade')
        quantidade = wait.until(ec.presence_of_element_located((By.XPATH,'//input[@id="id_quantidade"]')))
        quantidade.send_keys(pedido.quantidade)
        time.sleep(1)

        print('preço')
        valor = wait.until(ec.presence_of_element_located((By.XPATH,'//input[@id="id_preco_final"]')))
        valor.clear()
        time.sleep(1)
        preco = str(round(pedido.valor,10)).replace('.',',')
        valor.click()
        action.send_keys(preco).perform()
        #valor.send_keys(preco)
        time.sleep(1)

        print('info_ad')
        info_ad = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_informacoes_adicionais"]')))
        info_ad.send_keys('Importado via Excel')
        time.sleep(1)

        print('salva produto')
        slv_pro =  wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="adicao_produto"]/form/div[3]/a[1]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", slv_pro)
        slv_pro.click()
        time.sleep(2)


    elif pedido.acao == 3:
        print('ultimo produto',pedido.produto)
        
        print('produto')
        action.send_keys(pedido.produto).perform()
        time.sleep(2)
        adc_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        #avegador.execute_script("arguments[0].scrollIntoView();", adc_pro)
        adc_pro.click()
        time.sleep(1)

        print('quantidade')
        quantidade = wait.until(ec.presence_of_element_located((By.XPATH,'//input[@id="id_quantidade"]')))
        quantidade.send_keys(pedido.quantidade)
        time.sleep(1)

        print('preço')
        valor = wait.until(ec.presence_of_element_located((By.XPATH,'//input[@id="id_preco_final"]')))
        valor.clear()
        time.sleep(1)
        preco = str(round(pedido.valor,10)).replace('.',',')
        valor.click()
        action.send_keys(preco).perform()
        #valor.send_keys(preco)
        time.sleep(1)

        print('info_ad')
        info_ad = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_informacoes_adicionais"]')))
        info_ad.send_keys('Importado via Excel')
        time.sleep(1)

        print('salva produto')
        slv_pro =  wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="adicao_produto"]/form/div[3]/a[1]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", slv_pro)
        slv_pro.click()
        time.sleep(2)
        
        print('terminei produtos')
        terminei_prod = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="botao_terminei_de_adicionar"]')))
        navegador.execute_script("arguments[0].scrollIntoView();", terminei_prod)
        navegador.execute_script("arguments[0].click();", terminei_prod)
        time.sleep(1)

        print('vendedor')
        vendedor = wait.until(ec.presence_of_element_located((By.XPATH,'//input[@id="id_criador"]')))
        vendedor.send_keys(pedido.vendedor)
        time.sleep(1)

        print('pagamento')
        cond_pgto = wait.until(ec.presence_of_element_located((By.XPATH,'//input[@id="id_cond_pagamento"]')))
        cond_pgto.send_keys(vlr_cond)
        time.sleep(1)

        print('frete')
        frete = wait.until(ec.presence_of_element_located((By.XPATH,'//input[@id="id_transportadora"]')))
        frete.send_keys(vlr_frete)
        time.sleep(1)
        
        print('salva pedido')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="simplest_modal"]/div[2]/form/div[2]/a[1]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", slv_pdv)
        #navegador.execute_script("arguments[0].click();", slv_pdv)
        slv_pdv.click()
        time.sleep(1)

        print('GERAR PEDIDO...')
        gera_pedido = wait.until(ec.element_to_be_clickable((By.LINK_TEXT,'Transformar em pedido')))
        #gera_pedido = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="gritter-item-1"]/div[2]/div[2]/div/a[1]')))
        navegador.execute_script("arguments[0].scrollIntoView();", gera_pedido)
        navegador.execute_script("arguments[0].click();", gera_pedido)
        
        print('alterar pedido')
        alterar = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="alterar_informacoes"]')))
        navegador.execute_script("arguments[0].scrollIntoView();", alterar)
        navegador.execute_script("arguments[0].click();", alterar)
        time.sleep(4)

        print('data emissao')
        data_emis = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_data_emissao"]')))
        data_emis.clear()
        data_emis.send_keys(pedido.data.strftime("%d/%m/%Y"))

        print('salva pedido')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="simplest_modal"]/div[2]/form/div[2]/a[1]')))
        navegador.execute_script("arguments[0].scrollIntoView();", slv_pdv)
        navegador.execute_script("arguments[0].click();", slv_pdv)
    
    elif pedido.acao == 2:
        print('informa cliente e representada e primeiro produto')

        #aba Pedidos
        aba = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="aba_pedidos"]')))
        navegador.execute_script("arguments[0].scrollIntoView();", aba)
        navegador.execute_script("arguments[0].click();", aba)

        criar = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="btn_criar_pedido"]')))
        navegador.execute_script("arguments[0].scrollIntoView();", criar)
        navegador.execute_script("arguments[0].click();", criar)
        
        print('cliente')
        action.send_keys(cnpj).perform()
        cli_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_campo_id_codigo_cliente"]/ul/li[1]')))
        navegador.execute_script("arguments[0].scrollIntoView();", cli_pro)
        navegador.execute_script("arguments[0].click();", cli_pro)

        print('representada')
        action.send_keys(vlr_repr).perform()
        rep_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_campo_id_codigo_representada"]/ul/li[1]')))
        navegador.execute_script("arguments[0].scrollIntoView();", rep_pro)
        navegador.execute_script("arguments[0].click();", rep_pro)

        print('produto')
        action.send_keys(pedido.produto).perform()
        time.sleep(2)
        adc_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        navegador.execute_script("arguments[0].scrollIntoView();", adc_pro)
        navegador.execute_script("arguments[0].click();", adc_pro)

        print('quantidade')
        action.send_keys(pedido.quantidade).perform()

        print('preço')
        valor = wait.until(ec.presence_of_element_located((By.XPATH,'//input[@id="id_preco_final"]')))
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        valor.send_keys(preco)

        print('info_ad')
        info_ad = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_informacoes_adicionais"]')))
        info_ad.send_keys('Importado via Excel')

        print('salva produto')
        slv_pro =  wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="adicao_produto"]/form/div[3]/a[1]')))
        navegador.execute_script("arguments[0].scrollIntoView();", slv_pro)
        navegador.execute_script("arguments[0].click();", slv_pro)
        
        print('terminei produtos')
        terminei_prod = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="botao_terminei_de_adicionar"]')))
        navegador.execute_script("arguments[0].scrollIntoView();", terminei_prod)
        navegador.execute_script("arguments[0].click();", terminei_prod)

        print('vendedor')
        vendedor = wait.until(ec.presence_of_element_located((By.XPATH,'//input[@id="id_criador"]')))
        vendedor.send_keys(pedido.vendedor)

        print('pagamento')
        cond_pgto = wait.until(ec.presence_of_element_located((By.XPATH,'//input[@id="id_cond_pagamento"]')))
        cond_pgto.send_keys(vlr_cond)

        print('frete')
        frete = wait.until(ec.presence_of_element_located((By.XPATH,'//input[@id="id_transportadora"]')))
        frete.send_keys(vlr_frete)
        
        print('salva pedido')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="simplest_modal"]/div[2]/form/div[2]/a[1]')))
        navegador.execute_script("arguments[0].scrollIntoView();", slv_pdv)
        navegador.execute_script("arguments[0].click();", slv_pdv)

        print('GERAR PEDIDO...')
        #gera_pedido = wait.until(ec.element_to_be_clickable((By.LINK_TEXT,'Transformar em pedido')))
        gera_pedido = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="gritter-item-1"]/div[2]/div[2]/div/a[1]')))
        navegador.execute_script("arguments[0].scrollIntoView();", gera_pedido)
        navegador.execute_script("arguments[0].click();", gera_pedido)

        print('alterar pedido')
        alterar = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="alterar_informacoes"]')))
        navegador.execute_script("arguments[0].scrollIntoView();", alterar)
        navegador.execute_script("arguments[0].click();", alterar)
        time.sleep(4)

        print('data emissao')
        data_emis = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_data_emissao"]')))
        data_emis.clear()
        data_emis.send_keys(pedido.data.strftime("%d/%m/%Y"))

        print('salva pedido')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="simplest_modal"]/div[2]/form/div[2]/a[1]')))
        navegador.execute_script("arguments[0].scrollIntoView();", slv_pdv)
        navegador.execute_script("arguments[0].click();", slv_pdv)

time.sleep(20)
