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

#opcoes = webdriver.FirefoxOptions()
servico = FirefoxService(executable_path=GeckoDriverManager().install())
#opcoes.add_experimental_option("excludeSwitches", ["enable-logging"])
#opcoes.add_experimental_option("detach", True)
navegador = webdriver.Firefox(service=servico)

wait = WebDriverWait(navegador, 50)
action = ActionChains(navegador)

#abre o sistema
navegador.get("https://app.mercos.com/")
#navegador.maximize_window()
time.sleep(1)

#autenticacao
#usuario = wait.until(ec.visibility_of_element_located((By.ID,'id_usuario')))
usuario = navegador.find_element(By.ID,'id_usuario')
usuario.send_keys('gabrieledani@gmail.com')

#senha = wait.until(ec.visibility_of_element_located((By.ID,'id_senha')))
senha = navegador.find_element(By.ID,'id_senha')
senha.send_keys('Vedafil2022')

#lg = wait.until(ec.element_to_be_clickable((By.ID,"botaoEfetuarLogin")))
lg = navegador.find_element(By.ID,"botaoEfetuarLogin")
#navegador.execute_script("arguments[0].scrollIntoView();", lg)
#navegador.execute_script("arguments[0].click();", lg)
lg.click()

vlr_frete = 'CIF (Frete Pago)'
vlr_cond = '28/42/56'
vlr_repr = 'Hengst Indústria de Filtros Ltda'

df_pedidos = pd.read_excel('banco_hengst_2016_.xlsx')

#tempooooooooo
time.sleep(10)

for pedido in df_pedidos.itertuples(name='pedidos',index=False):
    print('Pedido->',pedido)
    cnpj = str(pedido.CNPJ)
    cnpj = cnpj.replace('.','')
    cnpj = cnpj.replace('/','')
    cnpj = cnpj.replace('-','')
    
    if pedido.acao == 1:
        print('primeiro produto do pedido de vários')

        print('aba pedidos')
        aba = wait.until(ec.element_to_be_clickable((By.ID,'aba_pedidos')))
        time.sleep(1)
        aba.click()
        #time.sleep(1)

        print('botão pedidos')
        criar = wait.until(ec.element_to_be_clickable((By.ID,'btn_criar_pedido')))
        time.sleep(1)
        criar.click()
        #time.sleep(2)
        
        print('cliente')
        cliente = wait.until(ec.visibility_of_element_located((By.ID,'id_codigo_cliente')))
        #action.send_keys(cnpj).perform()
        cliente.send_keys(cnpj)
        time.sleep(1)
        action.send_keys(Keys.ENTER).perform()
        time.sleep(1)
        cli_encontrado = wait.until(ec.visibility_of_element_located((By.ID,'selecionado_autocomplete_id_codigo_cliente')))
        print('cliente foi',cli_encontrado.is_displayed())

        print('representada')
        repre = wait.until(ec.visibility_of_element_located((By.ID,'id_codigo_representada')))
        #action.send_keys(vlr_repr).perform()
        repre.send_keys(vlr_repr)
        time.sleep(1)
        action.send_keys(Keys.ENTER).perform()
        time.sleep(1)
        repre_encontrada = wait.until(ec.visibility_of_element_located((By.ID,'selecionado_autocomplete_id_codigo_representada')))
        print('representada foi',repre_encontrada.is_displayed())

        print('produto')
        produto = navegador.find_element(By.ID,'produto_autocomplete')
        #wait.until_not( ec.element_attribute_to_include((By.ID,'produto_autocomplete'),'disabled'))
        #wait.until(ec.element_to_be_clickable((produto)))
        while produto.get_property('disabled') == True or produto != navegador.switch_to.active_element:
            time.sleep(1)
        print(produto.is_enabled())
        print('habilitou produto')
        produto.send_keys(pedido.codigo)
        time.sleep(2)
        #action.send_keys(pedido.codigo).perform()
        adc_pro = wait.until(ec.visibility_of_element_located((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        #wait.until(ec.element_to_be_clickable(adc_pro))
        #time.sleep(1)
        adc_pro.click()
        #action.send_keys(Keys.ENTER).perform()
        time.sleep(1)

        print('quantidade')
        quantidade = wait.until(ec.visibility_of_element_located((By.ID,'id_quantidade')))
        #quantidade = navegador.find_element(By.ID,'id_quantidade')
        quantidade.send_keys(pedido.quantidade)
        time.sleep(1)

        print('preço')
        valor = wait.until(ec.element_to_be_clickable((By.ID,'id_preco_final')))
        valor.click()
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        valor.send_keys(preco)
        time.sleep(1)

        print('info_ad')
        info_ad = wait.until(ec.element_to_be_clickable((By.ID,'id_informacoes_adicionais')))
        info_ad.click()
        info_ad.send_keys('Importado via Excel')
        time.sleep(1)

        print('salva produto')
        slv_pro =  wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="adicao_produto"]/form/div[3]/a[1]')))
        slv_pro.click()
        time.sleep(1)

    elif pedido.acao == 0:
        print('continua informando produtos',pedido.produto)
        
        print('produto')
        produto = wait.until(ec.visibility_of_element_located((By.ID,'produto_autocomplete')))
        while produto.get_property('disabled') == True or produto != navegador.switch_to.active_element:
            time.sleep(1)
        produto.send_keys(pedido.codigo)
        time.sleep(2)
        #action.send_keys(pedido.codigo).perform()
        adc_pro = wait.until(ec.visibility_of_element_located((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        #wait.until(ec.element_to_be_clickable(adc_pro))
        #time.sleep(1)
        adc_pro.click()
        #action.send_keys(Keys.ENTER).perform()
        time.sleep(1)

        print('quantidade')
        quantidade = wait.until(ec.visibility_of_element_located((By.ID,'id_quantidade')))
        quantidade.send_keys(pedido.quantidade)
        time.sleep(1)

        print('preço')
        valor = wait.until(ec.element_to_be_clickable((By.ID,'id_preco_final')))
        valor.click()
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        valor.send_keys(preco)
        time.sleep(1)

        print('info_ad')
        info_ad = wait.until(ec.element_to_be_clickable((By.ID,'id_informacoes_adicionais')))
        info_ad.click()
        info_ad.send_keys('Importado via Excel')
        time.sleep(1)

        print('salva produto')
        slv_pro =  wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="adicao_produto"]/form/div[3]/a[1]')))
        slv_pro.click()
        time.sleep(1)

    elif pedido.acao == 3:
        print('ultimo produto',pedido.produto)
        
        print('produto')
        produto = wait.until(ec.visibility_of_element_located((By.ID,'produto_autocomplete')))
        while produto.get_property('disabled') == True or produto != navegador.switch_to.active_element:
            time.sleep(1)
        produto.send_keys(pedido.codigo)
        time.sleep(2)
        #action.send_keys(pedido.codigo).perform()
        adc_pro = wait.until(ec.visibility_of_element_located((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        #wait.until(ec.element_to_be_clickable(adc_pro))
        #time.sleep(1)
        adc_pro.click()
        #action.send_keys(Keys.ENTER).perform()
        time.sleep(1)

        print('quantidade')
        quantidade = wait.until(ec.visibility_of_element_located((By.ID,'id_quantidade')))
        quantidade.send_keys(pedido.quantidade)
        time.sleep(1)

        print('preço')
        valor = wait.until(ec.element_to_be_clickable((By.ID,'id_preco_final')))
        valor.click()
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        valor.send_keys(preco)
        time.sleep(1)

        print('info_ad')
        info_ad = wait.until(ec.element_to_be_clickable((By.ID,'id_informacoes_adicionais')))
        info_ad.click()
        info_ad.send_keys('Importado via Excel')
        time.sleep(1)

        print('salva produto')
        slv_pro =  wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="adicao_produto"]/form/div[3]/a[1]')))
        slv_pro.click()
        time.sleep(3)
        
        print('terminei produtos')
        terminei_prod = wait.until(ec.presence_of_element_located((By.ID,'botao_terminei_de_adicionar')))
        navegador.execute_script("arguments[0].scrollIntoView();", terminei_prod)
        wait.until(ec.element_to_be_clickable(terminei_prod))
        terminei_prod.click()
        time.sleep(3)

        print('vendedor')
        vendedor =wait.until(ec.visibility_of_element_located((By.ID,'id_criador')))
        vendedor.send_keys(pedido.vendedor)
        time.sleep(1)

        print('pagamento')
        cond_pgto = navegador.find_element(By.ID,'id_cond_pagamento')
        cond_pgto.send_keys(vlr_cond)
        time.sleep(1)

        print('frete')
        frete = navegador.find_element(By.ID,'id_transportadora')
        frete.send_keys(vlr_frete)
        time.sleep(1)
        
        print('salva pedido')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="simplest_modal"]/div[2]/form/div[2]/a[1]')))
        slv_pdv.click()
        time.sleep(2)

        print('GERAR PEDIDO...')
        gera_pedido = wait.until(ec.presence_of_element_located((By.LINK_TEXT,'Transformar em pedido')))
        navegador.execute_script("arguments[0].scrollIntoView();", gera_pedido)
        wait.until(ec.element_to_be_clickable((By.LINK_TEXT,'Transformar em pedido')))
        gera_pedido.click()
        time.sleep(3)

        print('alterar pedido')
        alterar = wait.until(ec.presence_of_element_located((By.ID,'alterar_informacoes')))
        navegador.execute_script("arguments[0].scrollIntoView();", alterar)
        wait.until(ec.element_to_be_clickable((By.ID,'alterar_informacoes')))
        alterar.click()
        time.sleep(3)

        print('data emissao')
        data_emis = wait.until(ec.visibility_of_element_located((By.ID,'id_data_emissao')))
        data_emis.clear()
        data_emis.send_keys(pedido.data.strftime("%d/%m/%Y"))
        time.sleep(1)

        print('salva pedido')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="simplest_modal"]/div[2]/form/div[2]/a[1]')))
        navegador.execute_script("arguments[0].scrollIntoView();", slv_pdv)
        slv_pdv.click()
        time.sleep(3)
    
    elif pedido.acao == 4:
        print('informa cliente e representada e primeiro produto')

        print('aba pedidos')
        aba = wait.until(ec.element_to_be_clickable((By.ID,'aba_pedidos')))
        time.sleep(1)
        aba.click()
        #time.sleep(1)

        print('botão pedidos')
        criar = wait.until(ec.element_to_be_clickable((By.ID,'btn_criar_pedido')))
        time.sleep(1)
        criar.click()
        #time.sleep(2)
        
        print('cliente')
        cliente = wait.until(ec.visibility_of_element_located((By.ID,'id_codigo_cliente')))
        #action.send_keys(cnpj).perform()
        cliente.send_keys(cnpj)
        time.sleep(1)
        action.send_keys(Keys.ENTER).perform()
        time.sleep(1)
        cli_encontrado = wait.until(ec.visibility_of_element_located((By.ID,'selecionado_autocomplete_id_codigo_cliente')))
        print('cliente foi',cli_encontrado.is_displayed())

        print('representada')
        repre = wait.until(ec.visibility_of_element_located((By.ID,'id_codigo_representada')))
        #action.send_keys(vlr_repr).perform()
        repre.send_keys(vlr_repr)
        time.sleep(1)
        action.send_keys(Keys.ENTER).perform()
        time.sleep(1)
        repre_encontrada = wait.until(ec.visibility_of_element_located((By.ID,'selecionado_autocomplete_id_codigo_representada')))
        print('representada foi',repre_encontrada.is_displayed())

        print('produto')
        produto = navegador.find_element(By.ID,'produto_autocomplete')
        #wait.until_not( ec.element_attribute_to_include((By.ID,'produto_autocomplete'),'disabled'))
        #wait.until(ec.element_to_be_clickable((produto)))
        while produto.get_property('disabled') == True or produto != navegador.switch_to.active_element:
            time.sleep(1)
        print(produto.is_enabled())
        print('habilitou produto')
        produto.send_keys(pedido.codigo)
        time.sleep(2)
        #action.send_keys(pedido.codigo).perform()
        adc_pro = wait.until(ec.visibility_of_element_located((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        #wait.until(ec.element_to_be_clickable(adc_pro))
        #time.sleep(1)
        adc_pro.click()
        #action.send_keys(Keys.ENTER).perform()
        time.sleep(1)

        print('quantidade')
        quantidade = wait.until(ec.visibility_of_element_located((By.ID,'id_quantidade')))
        #quantidade = navegador.find_element(By.ID,'id_quantidade')
        quantidade.send_keys(pedido.quantidade)
        time.sleep(1)

        print('preço')
        valor = wait.until(ec.element_to_be_clickable((By.ID,'id_preco_final')))
        valor.click()
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        valor.send_keys(preco)
        time.sleep(1)

        print('info_ad')
        info_ad = wait.until(ec.element_to_be_clickable((By.ID,'id_informacoes_adicionais')))
        info_ad.click()
        info_ad.send_keys('Importado via Excel')
        time.sleep(1)

        print('salva produto')
        slv_pro =  wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="adicao_produto"]/form/div[3]/a[1]')))
        slv_pro.click()
        time.sleep(1)
        
        print('terminei produtos')
        terminei_prod = wait.until(ec.presence_of_element_located((By.ID,'botao_terminei_de_adicionar')))
        navegador.execute_script("arguments[0].scrollIntoView();", terminei_prod)
        wait.until(ec.element_to_be_clickable(terminei_prod))
        terminei_prod.click()
        time.sleep(3)

        print('vendedor')
        vendedor =wait.until(ec.visibility_of_element_located((By.ID,'id_criador')))
        vendedor.send_keys(pedido.vendedor)
        time.sleep(1)

        print('pagamento')
        cond_pgto = navegador.find_element(By.ID,'id_cond_pagamento')
        cond_pgto.send_keys(vlr_cond)
        time.sleep(1)

        print('frete')
        frete = navegador.find_element(By.ID,'id_transportadora')
        frete.send_keys(vlr_frete)
        time.sleep(1)
        
        print('salva pedido')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="simplest_modal"]/div[2]/form/div[2]/a[1]')))
        slv_pdv.click()
        time.sleep(2)

        print('GERAR PEDIDO...')
        gera_pedido = wait.until(ec.presence_of_element_located((By.LINK_TEXT,'Transformar em pedido')))
        navegador.execute_script("arguments[0].scrollIntoView();", gera_pedido)
        wait.until(ec.element_to_be_clickable((By.LINK_TEXT,'Transformar em pedido')))
        gera_pedido.click()
        time.sleep(3)

        print('alterar pedido')
        alterar = wait.until(ec.presence_of_element_located((By.ID,'alterar_informacoes')))
        navegador.execute_script("arguments[0].scrollIntoView();", alterar)
        wait.until(ec.element_to_be_clickable((By.ID,'alterar_informacoes')))
        alterar.click()
        time.sleep(3)

        print('data emissao')
        data_emis = wait.until(ec.visibility_of_element_located((By.ID,'id_data_emissao')))
        data_emis.clear()
        data_emis.send_keys(pedido.data.strftime("%d/%m/%Y"))
        time.sleep(1)

        print('salva pedido')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="simplest_modal"]/div[2]/form/div[2]/a[1]')))
        navegador.execute_script("arguments[0].scrollIntoView();", slv_pdv)
        slv_pdv.click()
        time.sleep(3)

time.sleep(20)
