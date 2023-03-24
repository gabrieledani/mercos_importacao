from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec

import time
import pandas as pd

opcoes = webdriver.ChromeOptions()
servico = Service(ChromeDriverManager().install())
opcoes.add_experimental_option("excludeSwitches", ["enable-logging"])
opcoes.add_experimental_option("detach", True)
navegador = webdriver.Chrome(service=servico,options=opcoes)

#abre o sistema
navegador.get("https://app.mercos.com/")
#navegador.maximize_window()
time.sleep(5)

wait = WebDriverWait(navegador, 50)

#autenticacao
usuario = wait.until(ec.presence_of_element_located((By.ID,'id_usuario')))
usuario.send_keys('gabrieledani@gmail.com')
senha = wait.until(ec.presence_of_element_located((By.ID,'id_senha')))
senha.send_keys('Vedafil2022')

lg = wait.until(ec.element_to_be_clickable((By.ID,"botaoEfetuarLogin")))
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

        #aba Pedidos
        aba = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="aba_pedidos"]')))
        aba.click()
        criar = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="btn_criar_pedido"]')))
        criar.click()
        
        print('cliente')
        cliente = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_codigo_cliente"]')))
        cliente.send_keys(cnpj)
        cli_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_campo_id_codigo_cliente"]/ul/li[1]')))
        cli_pro.click()

        print('representada')
        representada = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_codigo_representada"]')))
        representada.send_keys(vlr_repr)
        rep_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_campo_id_codigo_representada"]/ul/li[1]')))
        rep_pro.click()

        print('produto')
        produto = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[2]/div[3]/section/div[3]/div[10]/div[7]/div[1]/div[2]/div[1]/input')))
        produto.send_keys(pedido.produto)
        adc_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        adc_pro.click()

        print('quantidade')
        quantidade = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_quantidade"]')))
        quantidade.send_keys(pedido.quantidade)

        print('preço')
        valor = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_preco_final"]')))
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        valor.send_keys(preco)

        print('info_ad')
        info_ad = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_informacoes_adicionais"]')))
        info_ad.send_keys('Importado via Excel')

        print('salva produto')
        slv_pro =  wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="adicao_produto"]/form/div[3]/a[1]')))
        slv_pro.click()

    elif pedido.acao == 0:
        print('continua informando produtos',pedido.produto)
        
        print('produto')
        produto = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[2]/div[3]/section/div[3]/div[10]/div[7]/div[1]/div[2]/div[1]/input')))
        produto.send_keys(pedido.produto)
        adc_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        adc_pro.click()

        print('quantidade')
        quantidade = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_quantidade"]')))
        quantidade.send_keys(pedido.quantidade)

        print('preço')
        valor = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_preco_final"]')))
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        valor.send_keys(preco)

        print('info_ad')
        info_ad = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_informacoes_adicionais"]')))
        info_ad.send_keys('Importado via Excel')

        print('salva produto')
        slv_pro =  wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="adicao_produto"]/form/div[3]/a[1]')))
        slv_pro.click()

    elif pedido.acao == 3:
        print('ultimo produto',pedido.produto)
        
        print('produto')
        produto = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[2]/div[3]/section/div[3]/div[10]/div[7]/div[1]/div[2]/div[1]/input')))
        produto.send_keys(pedido.produto)
        adc_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        adc_pro.click()

        print('quantidade')
        quantidade = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_quantidade"]')))
        quantidade.send_keys(pedido.quantidade)

        print('preço')
        valor = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_preco_final"]')))
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        valor.send_keys(preco)

        print('info_ad')
        info_ad = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_informacoes_adicionais"]')))
        info_ad.send_keys('Importado via Excel')

        print('salva produto')
        slv_pro =  wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="adicao_produto"]/form/div[3]/a[1]')))
        slv_pro.click()
        
        print('terminei produtos')
        terminei_prod = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="botao_terminei_de_adicionar"]')))
        navegador.execute_script("arguments[0].scrollIntoView();", terminei_prod)
        navegador.execute_script("arguments[0].click();", terminei_prod)

        print('vendedor')
        vendedor = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_criador"]')))
        vendedor.send_keys(pedido.vendedor)

        print('pagamento')
        cond_pgto = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_cond_pagamento"]')))
        cond_pgto.send_keys(vlr_cond)

        print('frete')
        frete = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_transportadora"]')))
        frete.send_keys(vlr_frete)
        
        print('salva pedido')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="simplest_modal"]/div[2]/form/div[2]/a[1]')))
        slv_pdv.click()

        print('GERAR PEDIDO...')
        #gera_pedido = wait.until(ec.element_to_be_clickable((By.LINK_TEXT,'Transformar em pedido')))
        gera_pedido = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="gritter-item-1"]/div[2]/div[2]/div/a[1]')))
        gera_pedido.click()

        print('alterar pedido')
        alterar = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="alterar_informacoes"]')))
        navegador.execute_script("arguments[0].scrollIntoView();", alterar)
        navegador.execute_script("arguments[0].click();", alterar)

        print('data emissao')
        data_emis = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_data_emissao"]')))
        data_emis.clear()
        data_emis.send_keys(pedido.data.strftime("%d/%m/%Y"))

        print('salva pedido')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="simplest_modal"]/div[2]/form/div[2]/a[1]')))
        slv_pdv.click()
    
    elif pedido.acao == 2:
        print('informa cliente e representada e primeiro produto')

        #aba Pedidos
        aba = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="aba_pedidos"]')))
        aba.click()
        criar = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="btn_criar_pedido"]')))
        criar.click()
        
        print('cliente')
        cliente = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_codigo_cliente"]')))
        cliente.send_keys(cnpj)
        cli_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_campo_id_codigo_cliente"]/ul/li[1]')))
        cli_pro.click()

        print('representada')
        representada = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_codigo_representada"]')))
        representada.send_keys(vlr_repr)
        rep_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_campo_id_codigo_representada"]/ul/li[1]')))
        rep_pro.click()

        print('produto')
        produto = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[2]/div[3]/section/div[3]/div[10]/div[7]/div[1]/div[2]/div[1]/input')))
        produto.send_keys(pedido.produto)
        adc_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        adc_pro.click()

        print('quantidade')
        quantidade = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_quantidade"]')))
        quantidade.send_keys(pedido.quantidade)

        print('preço')
        valor = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_preco_final"]')))
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        valor.send_keys(preco)

        print('info_ad')
        info_ad = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_informacoes_adicionais"]')))
        info_ad.send_keys('Importado via Excel')

        print('salva produto')
        slv_pro =  wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="adicao_produto"]/form/div[3]/a[1]')))
        slv_pro.click()
        
        print('terminei produtos')
        terminei_prod = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="botao_terminei_de_adicionar"]')))
        navegador.execute_script("arguments[0].scrollIntoView();", terminei_prod)
        navegador.execute_script("arguments[0].click();", terminei_prod)

        print('vendedor')
        vendedor = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_criador"]')))
        vendedor.send_keys(pedido.vendedor)

        print('pagamento')
        cond_pgto = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_cond_pagamento"]')))
        cond_pgto.send_keys(vlr_cond)

        print('frete')
        frete = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_transportadora"]')))
        frete.send_keys(vlr_frete)
        
        print('salva pedido')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="simplest_modal"]/div[2]/form/div[2]/a[1]')))
        slv_pdv.click()

        print('GERAR PEDIDO...')
        #gera_pedido = wait.until(ec.element_to_be_clickable((By.LINK_TEXT,'Transformar em pedido')))
        gera_pedido = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="gritter-item-1"]/div[2]/div[2]/div/a[1]')))
        gera_pedido.click()

        print('alterar pedido')
        alterar = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="alterar_informacoes"]')))
        navegador.execute_script("arguments[0].scrollIntoView();", alterar)
        navegador.execute_script("arguments[0].click();", alterar)

        print('data emissao')
        data_emis = wait.until(ec.presence_of_element_located((By.XPATH,'//*[@id="id_data_emissao"]')))
        data_emis.clear()
        data_emis.send_keys(pedido.data.strftime("%d/%m/%Y"))

        print('salva pedido')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="simplest_modal"]/div[2]/form/div[2]/a[1]')))
        slv_pdv.click()

time.sleep(20)
