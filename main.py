from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

import time
import pandas as pd

opcoes = webdriver.ChromeOptions()
servico = Service(ChromeDriverManager().install())
opcoes.add_experimental_option("excludeSwitches", ["enable-logging"])
navegador = webdriver.Chrome(service=servico,options=opcoes)
#abre o sistema
navegador.get("https://app.mercos.com/")
#navegador.maximize_window()
time.sleep(5)
#autenticacao
usuario = navegador.find_element(By.ID,'id_usuario')
usuario.send_keys('Daniel.vedafil@gmail.com')
senha = navegador.find_element(By.ID,'id_senha')
senha.send_keys('Mercos.2022')
navegador.find_element(By.ID,"botaoEfetuarLogin").click()
time.sleep(5)
#Emitir pedidos
navegador.find_element(By.ID,'aba_pedidos').click()
time.sleep(5)


#vlr_repr = 'Hengst Indústria de Filtros Ltda'#03.429.968/0001-26
vlr_frete = 'CIF (Frete Pago)'
vlr_cond = '28/42/56'

df_pedidos = pd.read_excel('banco_hengst_2012.xlsx')

for pedido in df_pedidos.itertuples(name='pedidos',index=False):
    print('Pedido->',pedido)
    cnpj = str(pedido.CNPJ)
    cnpj = cnpj.replace('.','')
    cnpj = cnpj.replace('/','')
    cnpj = cnpj.replace('-','')
    
    if pedido.acao == 2:
        #Pedido com 1 Produto - começa e termina
        print('informa cliente e representada e primeiro produto')

        navegador.find_element(By.ID,'aba_pedidos').click()
        time.sleep(1)
        navegador.find_element(By.ID,'btn_criar_pedido').click()
        time.sleep(1)
        
        print('cliente')
        cliente = navegador.find_element(By.ID,'id_codigo_cliente')
        
        cliente.send_keys(cnpj)
        time.sleep(2)
        cliente.send_keys(Keys.TAB)
        time.sleep(2)

        print('representada')
        representada = navegador.find_element(By.ID,'id_codigo_representada')
        
        representada.send_keys(pedido.representada)
        time.sleep(2)
        representada.send_keys(Keys.TAB)
        time.sleep(2)

        print('produto')
        produto = navegador.find_element(By.ID,'produto_autocomplete')
        
        produto.send_keys(pedido.produto)
        time.sleep(6)
        produto.send_keys(Keys.TAB)
        time.sleep(4)

        print('quantidade')
        quantidade = navegador.find_element(By.ID,'id_quantidade')
        
        quantidade.send_keys(pedido.quantidade)
        time.sleep(3)
        quantidade.send_keys(Keys.TAB)
        time.sleep(2)

        print('preço')
        valor = navegador.find_element(By.ID,'id_preco_final')
        
        valor.clear()
        
        valor.send_keys(str(round(pedido.valor,10)).replace('.',','))
        time.sleep(1)
        valor.send_keys(Keys.TAB)
        time.sleep(1)

        print('info_ad')
        info_ad = navegador.find_element(By.ID,'id_informacoes_adicionais')
        
        info_ad.send_keys('Importado via Excel')
        time.sleep(1)

        print('salva produto')
        navegador.find_element(By.CSS_SELECTOR,'a.botao.medio.primario').click()
        time.sleep(3)
        
        print('terminei produtos')
        navegador.find_element(By.ID,'botao_terminei_de_adicionar').click()
        time.sleep(3)

        print('vendedor')
        vendedor = navegador.find_element(By.ID,'id_criador')
        
        vendedor.send_keys(pedido.vendedor)
        time.sleep(1)
        vendedor.send_keys(Keys.TAB)
        time.sleep(1)

        print('pagamento')
        cond_pgto = navegador.find_element(By.ID,'id_cond_pagamento')
        
        cond_pgto.send_keys(vlr_cond)
        time.sleep(1)

        print('frete')
        frete = navegador.find_element(By.ID,'id_transportadora')
        
        frete.send_keys(vlr_frete)
        time.sleep(1)
        frete.send_keys(Keys.TAB)
        time.sleep(1)
        
        print('salva pedido')
        navegador.find_element(By.CSS_SELECTOR,'a.js-valida-frete-selecionado.botao.medio.primario').click()
        time.sleep(3)

        #Gerar Pedido
        print('GERAR PEDIDO...')
        navegador.find_element(By.LINK_TEXT,'Transformar em pedido').click()
        time.sleep(3)

        print('alterar pedido')
        navegador.find_element(By.ID,'alterar_informacoes').click()
        time.sleep(4)

        print('data emissao')
        data_emis = navegador.find_element(By.ID,'id_data_emissao')
        
        data_emis.click()
        
        data_emis.send_keys(pedido.data.strftime("%d/%m/%Y"))
        time.sleep(1)
        data_emis.send_keys(Keys.TAB)
        time.sleep(1)

        print('salva pedido')
        navegador.find_element(By.CSS_SELECTOR,'a.js-valida-frete-selecionado.botao.medio.primario').click()
        time.sleep(3)
    
    elif pedido.acao == 1:
        #Primeiro produto de um pedido de vários produtos
        print('primeiro produto do pedido de vários')

        navegador.find_element(By.ID,'aba_pedidos').click()
        time.sleep(1)
        navegador.find_element(By.ID,'btn_criar_pedido').click()
        time.sleep(1)
        
        print('cliente')
        cliente = navegador.find_element(By.ID,'id_codigo_cliente')
        
        cliente.send_keys(cnpj)
        time.sleep(2)
        cliente.send_keys(Keys.TAB)
        time.sleep(2)

        print('representada')
        representada = navegador.find_element(By.ID,'id_codigo_representada')
        
        representada.send_keys(pedido.representada)
        time.sleep(2)
        representada.send_keys(Keys.TAB)
        time.sleep(2)

        print('produto')
        produto = navegador.find_element(By.ID,'produto_autocomplete')
        
        produto.send_keys(pedido.produto)
        time.sleep(6)
        produto.send_keys(Keys.TAB)
        time.sleep(4)

        print('quantidade')
        quantidade = navegador.find_element(By.ID,'id_quantidade')
        
        quantidade.send_keys(pedido.quantidade)
        time.sleep(3)
        quantidade.send_keys(Keys.TAB)
        time.sleep(2)

        print('preço')
        valor = navegador.find_element(By.ID,'id_preco_final')
        
        valor.clear()
        
        valor.send_keys(str(round(pedido.valor,10)).replace('.',','))
        time.sleep(1)
        valor.send_keys(Keys.TAB)
        time.sleep(1)

        print('info_ad')
        info_ad = navegador.find_element(By.ID,'id_informacoes_adicionais')
        
        info_ad.send_keys('Importado via Excel')
        time.sleep(1)

        print('salva produto')
        navegador.find_element(By.CSS_SELECTOR,'a.botao.medio.primario').click()
        time.sleep(3)

    elif pedido.acao == 0:
        print('continua informando produtos',pedido.produto)
        
        print('produto')
        produto = navegador.find_element(By.ID,'produto_autocomplete')
        
        produto.send_keys(pedido.produto)
        time.sleep(6)
        produto.send_keys(Keys.TAB)
        time.sleep(4)

        print('quantidade')
        quantidade = navegador.find_element(By.ID,'id_quantidade')
        
        quantidade.send_keys(pedido.quantidade)
        time.sleep(3)
        quantidade.send_keys(Keys.TAB)
        time.sleep(2)

        print('preço')
        valor = navegador.find_element(By.ID,'id_preco_final')
        
        valor.clear()
        
        valor.send_keys(str(round(pedido.valor,10)).replace('.',','))
        time.sleep(1)
        valor.send_keys(Keys.TAB)
        time.sleep(1)

        print('info_ad')
        info_ad = navegador.find_element(By.ID,'id_informacoes_adicionais')
        
        info_ad.send_keys('Importado via Excel')
        time.sleep(1)

        print('salva produto')
        navegador.find_element(By.CSS_SELECTOR,'a.botao.medio.primario').click()
        time.sleep(3)

    elif pedido.acao == 3:
        
        print('ultimo produto produtos',pedido.produto)
        
        print('produto')
        produto = navegador.find_element(By.ID,'produto_autocomplete')
        
        produto.send_keys(pedido.produto)
        time.sleep(6)
        produto.send_keys(Keys.TAB)
        time.sleep(4)

        print('quantidade')
        quantidade = navegador.find_element(By.ID,'id_quantidade')
        
        quantidade.send_keys(pedido.quantidade)
        time.sleep(3)
        quantidade.send_keys(Keys.TAB)
        time.sleep(2)

        print('preço')
        valor = navegador.find_element(By.ID,'id_preco_final')
        
        valor.clear()
        
        valor.send_keys(str(round(pedido.valor,10)).replace('.',','))
        time.sleep(1)
        valor.send_keys(Keys.TAB)
        time.sleep(1)

        print('info_ad')
        info_ad = navegador.find_element(By.ID,'id_informacoes_adicionais')
        
        info_ad.send_keys('Importado via Excel')
        time.sleep(1)

        print('salva produto')
        navegador.find_element(By.CSS_SELECTOR,'a.botao.medio.primario').click()
        time.sleep(3)
        
        print('terminei produtos')
        navegador.find_element(By.ID,'botao_terminei_de_adicionar').click()
        time.sleep(3)

        print('vendedor')
        vendedor = navegador.find_element(By.ID,'id_criador')
        
        vendedor.send_keys(pedido.vendedor)
        time.sleep(1)
        vendedor.send_keys(Keys.TAB)
        time.sleep(1)

        print('pagamento')
        cond_pgto = navegador.find_element(By.ID,'id_cond_pagamento')
        
        cond_pgto.send_keys(vlr_cond)
        time.sleep(1)

        print('frete')
        frete = navegador.find_element(By.ID,'id_transportadora')
        
        frete.send_keys(vlr_frete)
        time.sleep(1)
        frete.send_keys(Keys.TAB)
        time.sleep(1)
        
        print('salva pedido')
        navegador.find_element(By.CSS_SELECTOR,'a.js-valida-frete-selecionado.botao.medio.primario').click()
        time.sleep(3)

        #Gerar Pedido
        print('GERAR PEDIDO...')
        navegador.find_element(By.LINK_TEXT,'Transformar em pedido').click()
        time.sleep(3)

        print('alterar pedido')
        navegador.find_element(By.ID,'alterar_informacoes').click()
        time.sleep(4)

        print('data emissao')
        data_emis = navegador.find_element(By.ID,'id_data_emissao')
        
        data_emis.click()
        
        data_emis.send_keys(pedido.data.strftime("%d/%m/%Y"))
        time.sleep(1)
        data_emis.send_keys(Keys.TAB)
        time.sleep(1)

        print('salva pedido')
        navegador.find_element(By.CSS_SELECTOR,'a.js-valida-frete-selecionado.botao.medio.primario').click()
        time.sleep(3)

time.sleep(20)

