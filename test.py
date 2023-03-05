from selenium import webdriver
from selenium.webdriver.firefox.service import Service as FirefoxService
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec

import time
import pandas as pd

servico = FirefoxService(executable_path=GeckoDriverManager().install())
navegador = webdriver.Firefox(service=servico)

wait = WebDriverWait(navegador, 20)
#element_to_be_clickable
#presence_of_element_located
#visibility_of_element_located
#fastrack = wait.until(ec.visibility_of_element_located((By.XPATH, "//a[@data-tracking-id='0_Fastrack']")))
#fastrack.click()


#abre o sistema
navegador.get("https://app.mercos.com/")
#navegador.maximize_window()
time.sleep(5)

#autenticacao
usuario = navegador.find_element(By.ID,'id_usuario')
usuario.send_keys('gabrieledani@gmail.com')
senha = navegador.find_element(By.ID,'id_senha')
senha.send_keys('Vedafil2022')

navegador.find_element(By.ID,"botaoEfetuarLogin").click()
time.sleep(5)

#Emitir pedidos
#navegador.find_element(By.ID,'aba_pedidos').click()
#time.sleep(5)

vlr_frete = 'CIF (Frete Pago)'
vlr_cond = '28/42/56'
vlr_repr = 'Hengst Indústria de Filtros Ltda'

df_pedidos = pd.read_excel('banco_hengst_2014.xlsx')

for pedido in df_pedidos.itertuples(name='pedidos',index=False):
    print('Pedido->',pedido)
    cnpj = str(pedido.CNPJ)
    cnpj = cnpj.replace('.','')
    cnpj = cnpj.replace('/','')
    cnpj = cnpj.replace('-','')
    
    
    if pedido.acao == 1:
        #Primeiro produto de um pedido de vários produtos
        print('primeiro produto do pedido de vários')

        navegador.find_element(By.ID,'aba_pedidos').click()
        time.sleep(5)
        #navegador.find_element(By.ID,'btn_criar_pedido').click()
        criar = wait.until(ec.element_to_be_clickable((By.ID,'btn_criar_pedido')))
        criar.click()
        time.sleep(5)
        
        print('cliente')
        #cliente = navegador.find_element(By.ID,'id_codigo_cliente')
        cliente = wait.until(ec.presence_of_element_located((By.ID,'id_codigo_cliente')))
        cliente.send_keys(cnpj)
        time.sleep(6)
        cliente.send_keys(Keys.TAB)
        time.sleep(6)

        print('representada')
        representada = navegador.find_element(By.ID,'id_codigo_representada')
        representada.send_keys(vlr_repr)
        time.sleep(5)
        representada.send_keys(Keys.TAB)
        time.sleep(5)

        print('produto')
        #produto = navegador.find_element(By.ID,'produto_autocomplete')
        produto = wait.until(ec.presence_of_element_located((By.ID,'produto_autocomplete')))
        produto.send_keys(pedido.produto)
        time.sleep(7)
        #produto.send_keys(Keys.TAB)
        #navegador.find_element(By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]').click()
        adc_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        adc_pro.click()
        time.sleep(5)

        print('quantidade')
        quantidade = wait.until(ec.presence_of_element_located((By.ID,'id_quantidade')))
        quantidade.send_keys(pedido.quantidade)
        time.sleep(3)

        print('preço')
        valor = navegador.find_element(By.ID,'id_preco_final')
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        print(preco)
        valor.send_keys(preco)
        time.sleep(3)
        valor.send_keys(Keys.TAB)
        time.sleep(3)

        print('info_ad')
        info_ad = navegador.find_element(By.ID,'id_informacoes_adicionais')
        info_ad.send_keys('Importado via Excel')
        time.sleep(3)

        print('salva produto')
        navegador.find_element(By.CSS_SELECTOR,'a.botao.medio.primario').click()
        time.sleep(6)

    elif pedido.acao == 0:
        print('continua informando produtos',pedido.produto)
        
        print('produto')
        #produto = navegador.find_element(By.ID,'produto_autocomplete')
        produto = wait.until(ec.presence_of_element_located((By.ID,'produto_autocomplete')))
        produto.send_keys(pedido.produto)
        time.sleep(7)
        #produto.send_keys(Keys.TAB)
        #navegador.find_element(By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]').click()
        adc_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        adc_pro.click()
        time.sleep(5)

        print('quantidade')
        quantidade = wait.until(ec.presence_of_element_located((By.ID,'id_quantidade')))
        quantidade.send_keys(pedido.quantidade)
        time.sleep(3)

        print('preço')
        valor = navegador.find_element(By.ID,'id_preco_final')
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        print(preco)
        valor.send_keys(preco)
        time.sleep(3)
        valor.send_keys(Keys.TAB)
        time.sleep(3)

        print('info_ad')
        info_ad = navegador.find_element(By.ID,'id_informacoes_adicionais')
        info_ad.send_keys('Importado via Excel')
        time.sleep(3)

        print('salva produto')
        navegador.find_element(By.CSS_SELECTOR,'a.botao.medio.primario').click()
        time.sleep(6)

    elif pedido.acao == 3:
        
        print('ultimo produto produtos',pedido.produto)
        
        print('produto')
        #produto = navegador.find_element(By.ID,'produto_autocomplete')
        produto = wait.until(ec.presence_of_element_located((By.ID,'produto_autocomplete')))
        produto.send_keys(pedido.produto)
        time.sleep(7)
        #produto.send_keys(Keys.TAB)
        #navegador.find_element(By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]').click()
        adc_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        adc_pro.click()
        time.sleep(5)

        print('quantidade')
        quantidade = wait.until(ec.presence_of_element_located((By.ID,'id_quantidade')))
        quantidade.send_keys(pedido.quantidade)
        time.sleep(3)

        print('preço')
        valor = navegador.find_element(By.ID,'id_preco_final')
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        print(preco)
        valor.send_keys(preco)
        time.sleep(3)
        valor.send_keys(Keys.TAB)
        time.sleep(3)

        print('info_ad')
        info_ad = navegador.find_element(By.ID,'id_informacoes_adicionais')
        info_ad.send_keys('Importado via Excel')
        time.sleep(3)

        print('salva produto')
        navegador.find_element(By.CSS_SELECTOR,'a.botao.medio.primario').click()
        time.sleep(6)
        
        print('terminei produtos')
        #navegador.find_element(By.ID,'botao_terminei_de_adicionar').click()
        terminei_prod = navegador.find_element(By.ID,'botao_terminei_de_adicionar')
        navegador.execute_script("arguments[0].scrollIntoView();", terminei_prod)
        navegador.execute_script("arguments[0].click();", terminei_prod)
        time.sleep(6)

        print('vendedor')
        #vendedor = navegador.find_element(By.ID,'id_criador')
        vendedor = wait.until(ec.presence_of_element_located((By.ID,'id_criador')))
        vendedor.send_keys(pedido.vendedor)
        time.sleep(3)
        vendedor.send_keys(Keys.TAB)
        time.sleep(3)

        print('pagamento')
        cond_pgto = navegador.find_element(By.ID,'id_cond_pagamento')
        cond_pgto.send_keys(vlr_cond)
        time.sleep(3)

        print('frete')
        frete = navegador.find_element(By.ID,'id_transportadora')
        frete.send_keys(vlr_frete)
        time.sleep(3)
        frete.send_keys(Keys.TAB)
        time.sleep(3)
        
        print('salva pedido')
        navegador.find_element(By.CSS_SELECTOR,'a.js-valida-frete-selecionado.botao.medio.primario').click()
        time.sleep(6)

        #Gerar Pedido
        print('GERAR PEDIDO...')
        #navegador.find_element(By.LINK_TEXT,'Transformar em pedido').click()
        gera_pedido = wait.until(ec.element_to_be_clickable((By.LINK_TEXT,'Transformar em pedido')))
        gera_pedido.click()
        time.sleep(7)

        print('alterar pedido')
        #navegador.find_element(By.ID,'alterar_informacoes').click()
        alterar = navegador.find_element(By.ID,'alterar_informacoes')
        navegador.execute_script("arguments[0].scrollIntoView();", alterar)
        navegador.execute_script("arguments[0].click();", alterar)
        time.sleep(7)

        print('data emissao')
        #data_emis = navegador.find_element(By.ID,'id_data_emissao')
        data_emis = wait.until(ec.presence_of_element_located((By.ID,'id_data_emissao')))
        #navegador.execute_script("arguments[0].click();", data_emis)
        data_emis.click()
        data_emis.send_keys(pedido.data.strftime("%d/%m/%Y"))
        time.sleep(3)
        data_emis.send_keys(Keys.TAB)
        time.sleep(3)

        print('salva pedido')
        navegador.find_element(By.CSS_SELECTOR,'a.js-valida-frete-selecionado.botao.medio.primario').click()
        time.sleep(6)
    
    elif pedido.acao == 2:
        #Pedido com 1 Produto - começa e termina
        print('informa cliente e representada e primeiro produto')

        navegador.find_element(By.ID,'aba_pedidos').click()
        time.sleep(5)
        #navegador.find_element(By.ID,'btn_criar_pedido').click()
        criar = wait.until(ec.element_to_be_clickable((By.ID,'btn_criar_pedido')))
        criar.click()
        time.sleep(5)
        
        print('cliente')
        #cliente = navegador.find_element(By.ID,'id_codigo_cliente')
        cliente = wait.until(ec.presence_of_element_located((By.ID,'id_codigo_cliente')))
        cliente.send_keys(cnpj)
        time.sleep(6)
        cliente.send_keys(Keys.TAB)
        time.sleep(5)

        print('representada')
        representada = navegador.find_element(By.ID,'id_codigo_representada')
        representada.send_keys(vlr_repr)
        time.sleep(6)
        representada.send_keys(Keys.TAB)
        time.sleep(5)

        print('produto')

        #element_to_be_clickable
        #presence_of_element_located
        #visibility_of_element_located
        #fastrack = wait.until(ec.visibility_of_element_located((By.XPATH, "//a[@data-tracking-id='0_Fastrack']")))
        #fastrack.click()

        #produto = navegador.find_element(By.ID,'produto_autocomplete')
        produto = wait.until(ec.presence_of_element_located((By.ID,'produto_autocomplete')))
        produto.send_keys(pedido.produto)
        time.sleep(7)
        #produto.send_keys(Keys.TAB)
        #navegador.find_element(By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]').click()
        adc_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        adc_pro.click()
        time.sleep(5)

        print('quantidade')
        #quantidade = navegador.find_element(By.ID,'id_quantidade')
        quantidade = wait.until(ec.presence_of_element_located((By.ID,'id_quantidade')))
        quantidade.send_keys(pedido.quantidade)
        time.sleep(3)
        #quantidade.send_keys(Keys.TAB)
        #time.sleep(5)

        print('preço')
        valor = navegador.find_element(By.ID,'id_preco_final')
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        print(preco)
        valor.send_keys(preco)
        time.sleep(3)
        valor.send_keys(Keys.TAB)
        time.sleep(3)

        print('info_ad')
        info_ad = navegador.find_element(By.ID,'id_informacoes_adicionais')
        info_ad.send_keys('Importado via Excel')
        time.sleep(3)

        print('salva produto')
        navegador.find_element(By.CSS_SELECTOR,'a.botao.medio.primario').click()
        time.sleep(6)
        
        print('terminei produtos')
        #navegador.find_element(By.ID,'botao_terminei_de_adicionar').click()
        terminei_prod = navegador.find_element(By.ID,'botao_terminei_de_adicionar')
        navegador.execute_script("arguments[0].scrollIntoView();", terminei_prod)
        navegador.execute_script("arguments[0].click();", terminei_prod)
        time.sleep(6)

        print('vendedor')
        #vendedor = navegador.find_element(By.ID,'id_criador')
        vendedor = wait.until(ec.presence_of_element_located((By.ID,'id_criador')))
        vendedor.send_keys(pedido.vendedor)
        time.sleep(3)
        vendedor.send_keys(Keys.TAB)
        time.sleep(3)

        print('pagamento')
        cond_pgto = navegador.find_element(By.ID,'id_cond_pagamento')
        cond_pgto.send_keys(vlr_cond)
        time.sleep(3)

        print('frete')
        frete = navegador.find_element(By.ID,'id_transportadora')
        frete.send_keys(vlr_frete)
        time.sleep(3)
        frete.send_keys(Keys.TAB)
        time.sleep(3)
        
        print('salva pedido')
        navegador.find_element(By.CSS_SELECTOR,'a.js-valida-frete-selecionado.botao.medio.primario').click()
        time.sleep(6)

        #Gerar Pedido
        print('GERAR PEDIDO...')
        #navegador.find_element(By.LINK_TEXT,'Transformar em pedido').click()
        gera_pedido = wait.until(ec.element_to_be_clickable((By.LINK_TEXT,'Transformar em pedido')))
        gera_pedido.click()
        time.sleep(7)

        print('alterar pedido')
        #navegador.find_element(By.ID,'alterar_informacoes').click()
        alterar = navegador.find_element(By.ID,'alterar_informacoes')
        navegador.execute_script("arguments[0].scrollIntoView();", alterar)
        navegador.execute_script("arguments[0].click();", alterar)
        time.sleep(7)

        print('data emissao')
        data_emis = wait.until(ec.presence_of_element_located((By.ID,'id_data_emissao')))
        #navegador.execute_script("arguments[0].click();", data_emis)
        data_emis.click()
        data_emis.send_keys(pedido.data.strftime("%d/%m/%Y"))
        time.sleep(3)
        data_emis.send_keys(Keys.TAB)
        time.sleep(3)

        print('salva pedido')
        navegador.find_element(By.CSS_SELECTOR,'a.js-valida-frete-selecionado.botao.medio.primario').click()
        time.sleep(6)

time.sleep(20)

