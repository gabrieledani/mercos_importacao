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

wait = WebDriverWait(navegador, 50)

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
usuario = wait.until(ec.presence_of_element_located((By.ID,'id_usuario')))
usuario.send_keys('gabrieledani@gmail.com')
senha = wait.until(ec.presence_of_element_located((By.ID,'id_senha')))
senha.send_keys('Vedafil2022')

lg = wait.until(ec.element_to_be_clickable((By.ID,"botaoEfetuarLogin")))
lg.click()
time.sleep(5)

vlr_frete = 'CIF (Frete Pago)'
vlr_cond = '28/42/56'
vlr_repr = 'Hengst Indústria de Filtros Ltda'

df_pedidos = pd.read_excel('banco_hengst_2013.xlsx')

for pedido in df_pedidos.itertuples(name='pedidos',index=False):
    print('Pedido->',pedido)
    cnpj = str(pedido.CNPJ)
    cnpj = cnpj.replace('.','')
    cnpj = cnpj.replace('/','')
    cnpj = cnpj.replace('-','')
    
    if pedido.acao == 1:
        #Primeiro produto de um pedido de vários produtos
        print('primeiro produto do pedido de vários')

        # aba Pedidos
        aba = wait.until(ec.element_to_be_clickable((By.XPATH,'/html/body/div[2]/nav/ul/li[3]/a')))
        aba.click()
        time.sleep(2)
        criar = wait.until(ec.element_to_be_clickable((By.XPATH,'/html/body/div[2]/div[2]/section/div[2]/div[1]/div[2]/a[1]')))
        criar.click()
        time.sleep(2)
        
        print('cliente')
        cliente = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[2]/div[3]/section/div[3]/div[10]/div[5]/div/div[2]/form/div[1]/input')))
        cliente.send_keys(cnpj)
        time.sleep(2)
        cliente.send_keys(Keys.TAB)
        time.sleep(2)

        print('representada')
        representada = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[2]/div[3]/section/div[3]/div[10]/div[6]/div[2]/form/div[1]/input')))
        representada.send_keys(vlr_repr)
        time.sleep(2)
        representada.send_keys(Keys.TAB)
        time.sleep(2)

        print('produto')
        produto = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[2]/div[3]/section/div[3]/div[10]/div[7]/div[1]/div[2]/div[1]/input')))
        produto.send_keys(pedido.produto)
        time.sleep(3)
        adc_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        adc_pro.click()
        time.sleep(2)

        print('quantidade')
        quantidade = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[5]/div/form/div[2]/div[12]/div[1]/div/div[2]/div[2]/div/div[1]/input')))
        quantidade.send_keys(pedido.quantidade)
        time.sleep(3)

        print('preço')
        valor = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[5]/div/form/div[2]/div[15]/div[1]/div/div/input')))
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        #print(preco)
        valor.send_keys(preco)
        time.sleep(2)
        valor.send_keys(Keys.TAB)
        time.sleep(2)

        print('info_ad')
        info_ad = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[5]/div/form/div[2]/div[17]/div/div/div[2]/textarea')))
        info_ad.send_keys('Importado via Excel')
        time.sleep(2)

        print('salva produto')
        slv_pro =  wait.until(ec.element_to_be_clickable((By.XPATH,'/html/body/div[5]/div/form/div[3]/a[1]')))
        slv_pro.click()
        time.sleep(3)

    elif pedido.acao == 0:
        print('continua informando produtos',pedido.produto)
        
        print('produto')
        produto = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[2]/div[3]/section/div[3]/div[10]/div[7]/div[1]/div[2]/div[1]/input')))
        produto.send_keys(pedido.produto)
        time.sleep(3)
        adc_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        adc_pro.click()
        time.sleep(2)

        print('quantidade')
        quantidade = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[5]/div/form/div[2]/div[12]/div[1]/div/div[2]/div[2]/div/div[1]/input')))
        quantidade.send_keys(pedido.quantidade)
        time.sleep(3)

        print('preço')
        valor = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[5]/div/form/div[2]/div[15]/div[1]/div/div/input')))
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        #print(preco)
        valor.send_keys(preco)
        time.sleep(2)
        valor.send_keys(Keys.TAB)
        time.sleep(2)

        print('info_ad')
        info_ad = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[5]/div/form/div[2]/div[17]/div/div/div[2]/textarea')))
        info_ad.send_keys('Importado via Excel')
        time.sleep(2)

        print('salva produto')
        slv_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'/html/body/div[5]/div/form/div[3]/a[1]')))
        slv_pro.click()
        time.sleep(2)

    elif pedido.acao == 3:
        
        print('ultimo produto produtos',pedido.produto)
        
        print('produto')
        produto = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[2]/div[3]/section/div[3]/div[10]/div[7]/div[1]/div[2]/div[1]/input')))
        produto.send_keys(pedido.produto)
        time.sleep(3)
        adc_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        adc_pro.click()
        time.sleep(2)

        print('quantidade')
        quantidade = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[5]/div/form/div[2]/div[12]/div[1]/div/div[2]/div[2]/div/div[1]/input')))
        quantidade.send_keys(pedido.quantidade)
        time.sleep(3)

        print('preço')
        valor = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[5]/div/form/div[2]/div[15]/div[1]/div/div/input')))
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        print(preco)
        valor.send_keys(preco)
        time.sleep(2)
        valor.send_keys(Keys.TAB)
        time.sleep(2)

        print('info_ad')
        info_ad = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[5]/div/form/div[2]/div[17]/div/div/div[2]/textarea')))
        info_ad.send_keys('Importado via Excel')
        time.sleep(2)

        print('salva produto')
        slv_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'/html/body/div[5]/div/form/div[3]/a[1]')))
        slv_pro.click()
        time.sleep(3)
        
        print('terminei produtos')
        terminei_prod = wait.until(ec.element_to_be_clickable((By.XPATH,'/html/body/div[2]/div[3]/section/div[3]/div[10]/div[7]/div[2]/div[2]/a')))
        #terminei_prod.click()
        #terminei_prod = navegador.find_element(By.ID,'botao_terminei_de_adicionar')
        navegador.execute_script("arguments[0].scrollIntoView();", terminei_prod)
        navegador.execute_script("arguments[0].click();", terminei_prod)
        time.sleep(3)

        print('vendedor')
        vendedor = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[5]/div/div[2]/form/div[1]/div[1]/div[1]/div[4]/div/div[2]/select')))
        vendedor.send_keys(pedido.vendedor)
        time.sleep(2)
        vendedor.send_keys(Keys.TAB)
        time.sleep(2)

        print('pagamento')
        cond_pgto = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[5]/div/div[2]/form/div[1]/div[2]/div[2]/div[1]/div/div[2]/select')))
        cond_pgto.send_keys(vlr_cond)
        time.sleep(2)

        print('frete')
        frete = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[5]/div/div[2]/form/div[1]/div[3]/div[2]/div[2]/div/div[2]/select')))
        frete.send_keys(vlr_frete)
        time.sleep(2)
        frete.send_keys(Keys.TAB)
        time.sleep(2)
        
        print('salva pedido')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'/html/body/div[5]/div/div[2]/form/div[2]/a[1]')))
        slv_pdv.click()
        time.sleep(3)

        #Gerar Pedido
        print('GERAR PEDIDO...')
        #navegador.find_element(By.LINK_TEXT,'Transformar em pedido').click()
        #gera_pedido = wait.until(ec.element_to_be_clickable((By.LINK_TEXT,'Transformar em pedido')))
        gera_pedido = wait.until(ec.element_to_be_clickable((By.XPATH,'/html/body/div[2]/div[3]/section/div[3]/div[10]/div[9]/div/button[1]')))
        #gera_pedido.click()
        navegador.execute_script("arguments[0].scrollIntoView();", gera_pedido)
        navegador.execute_script("arguments[0].click();", gera_pedido)
        time.sleep(3)

        print('alterar pedido')
        alterar = wait.until(ec.element_to_be_clickable((By.XPATH,'/html/body/div[2]/div[3]/section/div[3]/div[10]/div[8]/div[2]/a')))
        #alterar.click()
        #alterar = navegador.find_element(By.ID,'alterar_informacoes')
        navegador.execute_script("arguments[0].scrollIntoView();", alterar)
        navegador.execute_script("arguments[0].click();", alterar)
        time.sleep(3)

        print('data emissao')
        data_emis = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[7]/div/div[2]/form/div[1]/div[1]/div[1]/div[2]/div/div[2]/input')))
        data_emis.click()
        data_emis.send_keys(pedido.data.strftime("%d/%m/%Y"))
        time.sleep(3)
        data_emis.send_keys(Keys.TAB)
        time.sleep(3)

        print('salva pedido')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'/html/body/div[5]/div/div[2]/form/div[2]/a[1]')))
        slv_pdv.click()
        time.sleep(3)
    
    elif pedido.acao == 2:
        #Pedido com 1 Produto - começa e termina
        print('informa cliente e representada e primeiro produto')

        # aba Pedidos
        aba = wait.until(ec.element_to_be_clickable((By.XPATH,'/html/body/div[2]/nav/ul/li[3]/a')))
        aba.click()
        time.sleep(2)
        criar = wait.until(ec.element_to_be_clickable((By.XPATH,'/html/body/div[2]/div[2]/section/div[2]/div[1]/div[2]/a[1]')))
        criar.click()
        time.sleep(2)
        
        print('cliente')
        cliente = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[2]/div[3]/section/div[3]/div[10]/div[5]/div/div[2]/form/div[1]/input')))
        cliente.send_keys(cnpj)
        time.sleep(2)
        cliente.send_keys(Keys.TAB)
        time.sleep(2)

        print('representada')
        representada = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[2]/div[3]/section/div[3]/div[10]/div[6]/div[2]/form/div[1]/input')))
        representada.send_keys(vlr_repr)
        time.sleep(2)
        representada.send_keys(Keys.TAB)
        time.sleep(2)

        print('produto')
        produto = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[2]/div[3]/section/div[3]/div[10]/div[7]/div[1]/div[2]/div[1]/input')))
        produto.send_keys(pedido.produto)
        time.sleep(3)
        adc_pro = wait.until(ec.element_to_be_clickable((By.XPATH,'//*[@id="div_adicionar_produto"]/ul/li[1]')))
        adc_pro.click()
        time.sleep(2)

        print('quantidade')
        quantidade = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[5]/div/form/div[2]/div[12]/div[1]/div/div[2]/div[2]/div/div[1]/input')))
        quantidade.send_keys(pedido.quantidade)
        time.sleep(3)

        print('preço')
        valor = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[5]/div/form/div[2]/div[15]/div[1]/div/div/input')))
        valor.clear()
        preco = str(round(pedido.valor,10)).replace('.',',')
        #print(preco)
        valor.send_keys(preco)
        time.sleep(2)
        valor.send_keys(Keys.TAB)
        time.sleep(2)

        print('info_ad')
        info_ad = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[5]/div/form/div[2]/div[17]/div/div/div[2]/textarea')))
        info_ad.send_keys('Importado via Excel')
        time.sleep(2)

        print('salva produto')
        slv_pro =  wait.until(ec.element_to_be_clickable((By.XPATH,'/html/body/div[5]/div/form/div[3]/a[1]')))
        slv_pro.click()
        time.sleep(3)
        
        print('terminei produtos')
        terminei_prod = wait.until(ec.element_to_be_clickable((By.XPATH,'/html/body/div[2]/div[3]/section/div[3]/div[10]/div[7]/div[2]/div[2]/a')))
        #terminei_prod.click()
        #terminei_prod = navegador.find_element(By.ID,'botao_terminei_de_adicionar')
        navegador.execute_script("arguments[0].scrollIntoView();", terminei_prod)
        navegador.execute_script("arguments[0].click();", terminei_prod)
        time.sleep(3)

        print('vendedor')
        vendedor = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[5]/div/div[2]/form/div[1]/div[1]/div[1]/div[4]/div/div[2]/select')))
        vendedor.send_keys(pedido.vendedor)
        time.sleep(2)
        vendedor.send_keys(Keys.TAB)
        time.sleep(2)

        print('pagamento')
        cond_pgto = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[5]/div/div[2]/form/div[1]/div[2]/div[2]/div[1]/div/div[2]/select')))
        cond_pgto.send_keys(vlr_cond)
        time.sleep(2)

        print('frete')
        frete = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[5]/div/div[2]/form/div[1]/div[3]/div[2]/div[2]/div/div[2]/select')))
        frete.send_keys(vlr_frete)
        time.sleep(2)
        frete.send_keys(Keys.TAB)
        time.sleep(2)
        
        print('salva pedido')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'/html/body/div[5]/div/div[2]/form/div[2]/a[1]')))
        slv_pdv.click()
        time.sleep(3)

        #Gerar Pedido
        print('GERAR PEDIDO...')
        #navegador.find_element(By.LINK_TEXT,'Transformar em pedido').click()
        #gera_pedido = wait.until(ec.element_to_be_clickable((By.LINK_TEXT,'Transformar em pedido')))
        gera_pedido = wait.until(ec.element_to_be_clickable((By.XPATH,'/html/body/div[2]/div[3]/section/div[3]/div[10]/div[9]/div/button[1]')))
        #gera_pedido.click()
        navegador.execute_script("arguments[0].scrollIntoView();", gera_pedido)
        navegador.execute_script("arguments[0].click();", gera_pedido)
        time.sleep(3)

        print('alterar pedido')
        alterar = wait.until(ec.element_to_be_clickable((By.XPATH,'/html/body/div[2]/div[3]/section/div[3]/div[10]/div[8]/div[2]/a')))
        #alterar.click()
        #alterar = navegador.find_element(By.ID,'alterar_informacoes')
        navegador.execute_script("arguments[0].scrollIntoView();", alterar)
        navegador.execute_script("arguments[0].click();", alterar)
        time.sleep(3)

        print('data emissao')
        data_emis = wait.until(ec.presence_of_element_located((By.XPATH,'/html/body/div[7]/div/div[2]/form/div[1]/div[1]/div[1]/div[2]/div/div[2]/input')))
        data_emis.click()
        data_emis.send_keys(pedido.data.strftime("%d/%m/%Y"))
        time.sleep(3)
        data_emis.send_keys(Keys.TAB)
        time.sleep(3)

        print('salva pedido')
        slv_pdv = wait.until(ec.element_to_be_clickable((By.XPATH,'/html/body/div[5]/div/div[2]/form/div[2]/a[1]')))
        slv_pdv.click()
        time.sleep(3)

time.sleep(20)

