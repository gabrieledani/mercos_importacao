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

wait = WebDriverWait(navegador, 150)

action = ActionChains(navegador)


#abre o sistema
navegador.get("https://app.mercos.com/")
#navegador.maximize_window()
time.sleep(2)

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

df_pedidos = pd.read_excel('banco_hengst_2015.xlsx')

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
        aba = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="aba_pedidos"]')))
        navegador.execute_script("arguments[0].scrollIntoView();", aba)
        aba.click()
        time.sleep(1)

        print('botão pedidos')
        criar = navegador.find_element(By.XPATH,'//*[@id="btn_criar_pedido"]')
        time.sleep(1)
        #navegador.execute_script("arguments[0].scrollIntoView();", criar)
        time.sleep(1)
        wait.until(ec.element_to_be_clickable(criar))
        time.sleep(1)
        criar.click()
        time.sleep(1)
        
        print('cliente')
        action.send_keys(cnpj).perform()
        time.sleep(1)
        action.send_keys(Keys.ENTER).perform()
        #cli_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_campo_id_codigo_cliente"]/ul/li[1]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", cli_pro)
        #cli_pro.click()
        time.sleep(1)

        print('representada')
        action.send_keys(vlr_repr).perform()
        time.sleep(1)
        action.send_keys(Keys.ENTER).perform()
        #rep_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_campo_id_codigo_representada"]/ul/li[1]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", rep_pro)
        #rep_pro.click()
        time.sleep(1)

        print('produto')
        action.send_keys(pedido.codigo).perform()
        time.sleep(2)
        action.send_keys(Keys.ENTER).perform()
        time.sleep(1)
        #adc_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", adc_pro)
        #navegador.execute_script("arguments[0].click();", adc_pro)

        print('quantidade')
        quantidade = wait.until(ec.visibility_of_element_located((By.XPATH,'//input[@id="id_quantidade"]')))
        quantidade.send_keys(pedido.quantidade)
        time.sleep(1)

        print('preço')
        valor = wait.until(ec.visibility_of_element_located((By.XPATH,'//input[@id="id_preco_final"]')))
        valor.click()
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        valor.send_keys(preco)
        time.sleep(1)

        print('info_ad')
        info_ad = wait.until(ec.visibility_of_element_located((By.XPATH,'//*[@id="id_informacoes_adicionais"]')))
        info_ad.send_keys('Importado via Excel')
        time.sleep(1)

        print('salva produto')
        slv_pro =  wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="adicao_produto"]/form/div[3]/a[1]')))
        #navegador.execute_script("arguments[0].scrollIntoView();", slv_pro)
        slv_pro.click()
        time.sleep(3)

    elif pedido.acao == 0:
        print('continua informando produtos',pedido.produto)
        
        print('produto')
        #######//*[@id="produto_autocomplete"]
        produto = wait.until(ec.visibility_of_element_located((By.XPATH,'//input[@id="produto_autocomplete"]')))
        action.send_keys(pedido.codigo).perform()
        time.sleep(2)
        action.send_keys(Keys.ENTER).perform()
        time.sleep(1)

        print('quantidade')
        quantidade = wait.until(ec.visibility_of_element_located((By.XPATH,'//input[@id="id_quantidade"]')))
        quantidade.send_keys(pedido.quantidade)
        time.sleep(1)

        print('preço')
        valor = wait.until(ec.visibility_of_element_located((By.XPATH,'//input[@id="id_preco_final"]')))
        valor.click()
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        valor.send_keys(preco)
        time.sleep(1)

        print('info_ad')
        info_ad = wait.until(ec.visibility_of_element_located((By.XPATH,'//*[@id="id_informacoes_adicionais"]')))
        info_ad.send_keys('Importado via Excel')
        time.sleep(1)

        print('salva produto')
        slv_pro =  wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="adicao_produto"]/form/div[3]/a[1]')))
        slv_pro.click()
        time.sleep(3)

    elif pedido.acao == 3:
        print('ultimo produto',pedido.produto)
        
        print('produto')
        #######//*[@id="produto_autocomplete"]
        produto = wait.until(ec.visibility_of_element_located((By.XPATH,'//input[@id="produto_autocomplete"]')))
        action.send_keys(pedido.codigo).perform()
        time.sleep(2)
        action.send_keys(Keys.ENTER).perform()
        time.sleep(1)

        print('quantidade')
        quantidade = wait.until(ec.visibility_of_element_located((By.XPATH,'//input[@id="id_quantidade"]')))
        quantidade.send_keys(pedido.quantidade)
        time.sleep(1)

        print('preço')
        valor = wait.until(ec.visibility_of_element_located((By.XPATH,'//input[@id="id_preco_final"]')))
        valor.click()
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        valor.send_keys(preco)
        time.sleep(1)

        print('info_ad')
        info_ad = wait.until(ec.visibility_of_element_located((By.XPATH,'//*[@id="id_informacoes_adicionais"]')))
        info_ad.send_keys('Importado via Excel')
        time.sleep(1)

        print('salva produto')
        slv_pro =  wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="adicao_produto"]/form/div[3]/a[1]')))
        slv_pro.click()
        time.sleep(3)
        
        print('terminei produtos')
        terminei_prod = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="botao_terminei_de_adicionar"]')))
        navegador.execute_script("arguments[0].scrollIntoView();", terminei_prod)
        terminei_prod.click()
        time.sleep(3)

        print('vendedor')
        vendedor = wait.until(ec.visibility_of_element_located((By.XPATH,'//*[@id="id_criador"]')))
        vendedor.send_keys(pedido.vendedor)
        time.sleep(1)

        print('pagamento')
        cond_pgto = wait.until(ec.visibility_of_element_located((By.XPATH,'//*[@id="id_cond_pagamento"]')))
        cond_pgto.send_keys(vlr_cond)
        time.sleep(1)

        print('frete')
        frete = wait.until(ec.visibility_of_element_located((By.XPATH,'//*[@id="id_transportadora"]')))
        frete.send_keys(vlr_frete)
        time.sleep(1)
        
        print('salva pedido')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="simplest_modal"]/div[2]/form/div[2]/a[1]')))
        slv_pdv.click()
        time.sleep(1)

        print('GERAR PEDIDO...')
        gera_pedido = wait.until(ec.element_to_be_clickable((By.LINK_TEXT,'Transformar em pedido')))
        navegador.execute_script("arguments[0].scrollIntoView();", gera_pedido)
        gera_pedido.click()
        time.sleep(3)

        print('alterar pedido')
        alterar = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="alterar_informacoes"]')))
        navegador.execute_script("arguments[0].scrollIntoView();", alterar)
        alterar.click()
        time.sleep(3)

        print('data emissao')
        data_emis = wait.until(ec.visibility_of_element_located((By.XPATH,'//*[@id="id_data_emissao"]')))
        data_emis.clear()
        data_emis.send_keys(pedido.data.strftime("%d/%m/%Y"))
        time.sleep(1)

        print('salva pedido')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="simplest_modal"]/div[2]/form/div[2]/a[1]')))
        navegador.execute_script("arguments[0].scrollIntoView();", slv_pdv)
        slv_pdv.click()
        time.sleep(1)
    
    elif pedido.acao == 4:
        print('informa cliente e representada e primeiro produto')

        print('aba pedidos')
        #aba = navegador.find_element(By.XPATH,'//*[@id="aba_pedidos"]')
        #navegador.execute_script("arguments[0].scrollIntoView();", aba)
        aba = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="aba_pedidos"]')))
        time.sleep(1)
        aba.click()
        time.sleep(2)

        print('botão pedidos')
        criar = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="btn_criar_pedido"]')))
        time.sleep(1)
        criar.click()
        time.sleep(2)
        
        print('cliente')
        action.send_keys(cnpj).perform()
        time.sleep(1)
        action.send_keys(Keys.ENTER).perform()
        time.sleep(1)

        print('representada')
        action.send_keys(vlr_repr).perform()
        time.sleep(1)
        action.send_keys(Keys.ENTER).perform()
        time.sleep(2)

        print('produto')
        action.send_keys(pedido.codigo).perform()
        time.sleep(2)
        action.send_keys(Keys.ENTER).perform()
        time.sleep(2)

        print('quantidade')
        quantidade = wait.until(ec.visibility_of_element_located((By.XPATH,'//input[@id="id_quantidade"]')))
        #action.send_keys(pedido.quantidade).perform()
        #time.sleep(1)
        #action.send_keys(Keys.TAB).perform()
        quantidade.send_keys(pedido.quantidade)
        time.sleep(2)

        print('preço')
        valor = wait.until(ec.visibility_of_element_located((By.XPATH,'//input[@id="id_preco_final"]')))
        valor.click()
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        valor.send_keys(preco)
        time.sleep(1)

        print('info_ad')
        info_ad = wait.until(ec.visibility_of_element_located((By.XPATH,'//*[@id="id_informacoes_adicionais"]')))
        info_ad.click()
        info_ad.send_keys('Importado via Excel')
        time.sleep(1)

        print('salva produto')
        slv_pro =  wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="adicao_produto"]/form/div[3]/a[1]')))
        slv_pro.click()
        
        print('terminei produtos')
        terminei_prod = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="botao_terminei_de_adicionar"]')))
        navegador.execute_script("arguments[0].scrollIntoView();", terminei_prod)
        terminei_prod.click()
        time.sleep(3)

        print('vendedor')
        vendedor = wait.until(ec.visibility_of_element_located((By.XPATH,'//*[@id="id_criador"]')))
        vendedor.send_keys(pedido.vendedor)
        time.sleep(1)

        print('pagamento')
        cond_pgto = wait.until(ec.visibility_of_element_located((By.XPATH,'//*[@id="id_cond_pagamento"]')))
        cond_pgto.send_keys(vlr_cond)
        time.sleep(1)

        print('frete')
        frete = wait.until(ec.visibility_of_element_located((By.XPATH,'//*[@id="id_transportadora"]')))
        frete.send_keys(vlr_frete)
        time.sleep(1)
        
        print('salva pedido')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="simplest_modal"]/div[2]/form/div[2]/a[1]')))
        slv_pdv.click()
        time.sleep(1)

        print('GERAR PEDIDO...')
        gera_pedido = wait.until(ec.element_to_be_clickable((By.LINK_TEXT,'Transformar em pedido')))
        navegador.execute_script("arguments[0].scrollIntoView();", gera_pedido)
        gera_pedido.click()
        time.sleep(3)

        print('alterar pedido')
        alterar = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="alterar_informacoes"]')))
        navegador.execute_script("arguments[0].scrollIntoView();", alterar)
        alterar.click()
        time.sleep(3)

        print('data emissao')
        data_emis = wait.until(ec.visibility_of_element_located((By.XPATH,'//*[@id="id_data_emissao"]')))
        data_emis.clear()
        data_emis.send_keys(pedido.data.strftime("%d/%m/%Y"))
        time.sleep(1)

        print('salva pedido')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="simplest_modal"]/div[2]/form/div[2]/a[1]')))
        navegador.execute_script("arguments[0].scrollIntoView();", slv_pdv)
        slv_pdv.click()
        time.sleep(10)

time.sleep(20)
