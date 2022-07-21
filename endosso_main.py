from ast import expr_context
from doctest import Example
from select import select
import sys, os
from datetime import date, datetime
import time
from time import sleep
from tkinter import E
from traceback import print_list
from unittest import defaultTestLoader
from attr import field # usar 'sleep(value)' ao invés de 'time.sleep(value)'
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.select import Select
import pymssql
import endosso_functions as EF
import db
import globalconf as GC
# from pyUFbr.baseuf import ufbr
from unidecode import unidecode
import glob

def initializeRpa(rs):
    conn = pymssql.connect(server=GC.server, user=GC.user, password=GC.password, database=GC.database)
    cursor = conn.cursor()

    selectCobertura = f''' SELECT * FROM coberturas_mds_endosso where rs = {rs} '''
    cursor.execute(selectCobertura)
    listCoberturas = cursor.fetchall()
    conn.close()

    correctStatus =  ['Ativo','Ativa','Emitido','Emitida']

    (
    seguradora
    ,tipoEndosso
    ,subtipoEndosso
    ,nuapolice
    ,nupropostacia
    ,segurado
    ,cpf_cnpj
    ,consultor
    ,datasolicitacao
    ,inicioVigencia
    ,fimVigencia
    ,email
    ,cep
    ,endereco
    ,cidade
    ,estado
    ,numero
    ,complemento
    ,ddd
    ,telefone
    ,dddc
    ,celular
    ,produto
    ,premioLiq
    ,adicional
    ,custo
    ,formarPgto
    ,qtdParcelas
    ,dataVencimento
    ,valorIOF
    ,premioTotal
    ,nupropostaquiver
    ,datacadastroquiver
    ,status
    ,statusRPA
    ,caminho_doc   
    ,Fabricante
    ,Modelo
    ,ano
    ,Placa
    ,Chassi
    ,FIPE
    ,CEPRISCO
    ,dataSubs
    ,classeBonus
    ,codCI
    ,categoria
    ,franquia_reduzida
    ,premio_liquido
    ,nomeCondutor
    ,cpfCondutor
    ,estadoCivilCondutor
    ,idadeCondutor
    ,detalhamento_franquia_vidro
    ,nota_fiscal
    ,nome_arq
    ,renavam
    ,genero
    ,cep_pernoite
    ,data_saida
    ,combustivel
    ) = db.initDB(rs)

    if classeBonus[0] == '0':
        classeBonus = classeBonus[1]

    insist = 5
    alerts = 5
    insisttext = 2
    defaultsleep = 2

    def AjusteCadastral(driver,wait):
        try:
            if 'cancelamento' in tipoEndosso.lower():
                return 1
            
            driver.switch_to.default_content()
            driver.switch_to.frame('ZonaInterna')
            driver.switch_to.frame('ZonaInterna')
            driver.switch_to.frame('Documento')

            driver.execute_script('abreCadastroCliente();')
            time.sleep(defaultsleep)

            driver.switch_to.default_content()
            driver.switch_to.frame('Cliente1')
            sleep(2)
            EF.fill(driver, '//*[@id="Cliente_EMail"]', email,insist)

            #Endereço
            sleep(2)
            EF.click(driver, '/html/body/form/div[7]/div/div[2]/div[6]/div[3]/div/div/div/div[1]/h2/label', insist)
            sleep(2)

            driver.switch_to.frame('FrameClienteEndereco1')

            campos_enderecos = driver.find_elements(By.XPATH,'/html/body/form/div[3]/div/span[1]/span/div/div[3]/div[3]/div/table/tbody/tr')
            
            #Faz a validação para trocar de paginas se ter muitos endereços
            validacao_endereco = False
            ValidacaoIncluir = False
            EscreverEndereco = False
            sleep(2)
            ultima_pag = driver.find_element_by_xpath('/html/body/form/div[3]/div/span[1]/span/div/div[5]/div/table/tbody/tr/td[2]/table/tbody/tr/td[4]/span').get_attribute("innerHTML")
            ultima_pag = int(ultima_pag)
            pag_atual = 1
            limite = 4
            sleep(1)
            #Validação que troca de pagina 
            while (pag_atual <= ultima_pag):
                if pag_atual <= ultima_pag and validacao_endereco == True:
                    break
                while not validacao_endereco == True:
                    for x in range(2,len(campos_enderecos)+1):
                        #Pega os campos de endereço
                        lista = []
                        EnderecoValida =EF.getText(driver,'/html/body/form/div[3]/div/span[1]/span/div/div[3]/div[3]/div/table/tbody/tr[{}]/td[3]'.format(x),insist,'innerText').lower()
                        NumeroValida = EF.getText(driver,'/html/body/form/div[3]/div/span[1]/span/div/div[3]/div[3]/div/table/tbody/tr[{}]/td[4]'.format(x),insist,'innerText').lower()
                        CepValida = EF.getText(driver,'/html/body/form/div[3]/div/span[1]/span/div/div[3]/div[3]/div/table/tbody/tr[{}]/td[5]'.format(x),insist,'innerText').lower()
                        ComplementoValida = EF.getText(driver,'/html/body/form/div[3]/div/span[1]/span/div/div[3]/div[3]/div/table/tbody/tr[{}]/td[6]'.format(x),insist,'innerText').lower()
                        CidadeValida = EF.getText(driver,'/html/body/form/div[3]/div/span[1]/span/div/div[3]/div[3]/div/table/tbody/tr[{}]/td[7]'.format(x),insist,'innerText').lower()
                        EstadoValida = EF.getText(driver,'/html/body/form/div[3]/div/span[1]/span/div/div[3]/div[3]/div/table/tbody/tr[{}]/td[8]'.format(x),insist,'innerText').lower()

                        lista.append(EnderecoValida)
                        lista.append(NumeroValida)
                        lista.append(CepValida)
                        lista.append(ComplementoValida)
                        lista.append(CidadeValida)
                        lista.append(EstadoValida)
                        #Retira os acentos
                        try:
                            EnderecoValida1 = unidecode(lista[0])
                        except:
                            pass
                        try:
                            CidadeValida1 = unidecode(lista[4])
                        except:
                            pass
                        try:
                            ComplementoValida1 = unidecode(lista[3])
                        except:
                            pass
                        try:
                            EstadoValida1 = unidecode(lista[5])
                        except:
                            pass
            

                        #Faz a validação
                        ValidacaoEndereco = EnderecoValida != endereco.lower() 
                        ValidaEnderecoAcento = EnderecoValida1 ==  endereco.lower()
                        ValidacaoNumero = NumeroValida != numero.lower()
                        ValidacaoCep = CepValida != cep.lower()
                        ValidacaoComplemento = ComplementoValida.strip() != complemento.lower()
                        ValidaComplementoAcento = ComplementoValida1 == complemento.lower()
                        ValidacaoCidade = CidadeValida != cidade.lower()
                        ValidacaoEstado = EstadoValida != estado.lower()
                        ValidacaoEstadoAcento = EstadoValida1 ==  estado.lower()
                        if ValidacaoCidade == True:
                            ValidaCidadeAcento = CidadeValida1 == cidade.lower()
                        else:
                            ValidaCidadeAcento = CidadeValida1 != cidade.lower()
                        # if '0' in numero:
                        #     listanum = [numero]
                        #     ultimonum = [z in listanum for z]
                        try:
                            if ValidacaoEndereco and  ValidacaoCep and ValidacaoComplemento:
                                if ValidaEnderecoAcento:
                                    print('Deu Certo')
                                    validacao_endereco = True
                                    break
                                else:
                                    print('Endereço diferente')
                                    x += 1
                                    continue
                            elif ValidacaoNumero:
                                print('Número diferente')
                                x += 1
                                continue
                            elif ValidacaoCep:
                                print('Cep diferente')
                                x += 1
                                continue
                            elif ValidacaoComplemento and ValidacaoCep:
                                if ValidaComplementoAcento:
                                    print('Deu Certo')
                                    validacao_endereco = True
                                    break
                                else:
                                    print('Complemento Diferente')
                                    x += 1
                                    continue
                            elif ValidacaoCidade:
                                if ValidaCidadeAcento:
                                    print('Deu Certo')
                                    validacao_endereco = True
                                    break
                                else:
                                    print('Cidade diferente')
                                    x += 1
                                    continue
                            elif ValidacaoEstado:
                                if ValidacaoEstadoAcento:
                                    print('Deu Certo')
                                    validacao_endereco = True
                                    break
                                else:
                                    print('Estado diferente')
                                    x += 1
                                    continue
                            else:
                                print('Deu Certo')
                                validacao_endereco = True
                                break
                        except:
                            driver.switch_to.default_content()
                            driver.switch_to.frame('Cliente1')
                            driver.execute_script("incluirClienteEndereco1();")
                            ValidacaoIncluir = False
                            pass
                
                    if ValidacaoIncluir == True and validacao_endereco == False:
                        driver.switch_to.default_content()
                        driver.switch_to.frame('Cliente1')
                        driver.execute_script("incluirClienteEndereco1();")
                        EscreverEndereco = True
                    else:
                        if validacao_endereco == False:
                            if pag_atual == ultima_pag:
                                driver.switch_to.default_content()
                                driver.switch_to.frame('Cliente1')
                                driver.execute_script("incluirClienteEndereco1();")
                                EscreverEndereco = True
                            else:
                                pass
                            
                    if EscreverEndereco == True:
                        driver.switch_to.frame('ZonaInterna')
                        sleep(2)
                        EF.fill(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[1]/div/input',cep,insist)
                        sleep(2)
                        driver.execute_script(f'document.getElementById("ClienteEnder_Endereco").value="{endereco}" ')
                        sleep(1)
                        driver.execute_script(f'document.getElementById("ClienteEnder_Numero").value="{numero}" ')

                        EF.fill(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[3]/div/div[1]/div/input',complemento,insist)
                        sleep(1)

                        EF.fill(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[3]/div/div[3]/div/input',cidade,insist)
                        sleep(1)

                        EF.fill(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[3]/div/div[4]/div/input',estado,insist)
                        sleep(1)

                        driver.execute_script("if (!document.getElementById('table_Toolbar').disabled) { evento('Gravar'); }")
                        sleep(1)
                        driver.execute_script("if (!document.getElementById('table_Toolbar').disabled) { evento('Incluir'); }")
                        sleep(1)
                        driver.execute_script("if (!document.getElementById('table_Toolbar').disabled) { evento('Voltar'); }")
                        sleep(1)
                        driver.switch_to.default_content()
                        alert = '/html/body/div[5]/div'
                        sleep(2)
                        try:
                            if driver.find_elements(By.XPATH,alert):
                                EF.click(driver,'/html/body/div[5]/div/div[3]/button[1]',insist)
                        except:pass
                        validacao_endereco = True
                        sleep(1)

                    if(pag_atual != ultima_pag):
                        sleep(1)
                        xpath = '//*[@id="next_BAR_GridCadastro2"]'
                        btn_prox =  EF.click(driver,xpath,insist)
                        sleep(1)
                        pag_atual += 1
                    else:
                        break   

            #Telefone
            sleep(1)
            driver.switch_to.default_content()
            driver.switch_to.frame('Cliente1')
            sleep(3)
            EF.click(driver, '//*[@id="icone-ClienteTelefone1"]', insist)
            sleep(2)
            driver.switch_to.frame('FrameClienteTelefone1')
            sleep(3)
            trTelefones = driver.find_elements(By.XPATH,'/html/body/form/div[3]/div/span[1]/span/div/div[3]/div[3]/div/table/tbody/tr')
            lenTelefone = len(trTelefones)+1

            listFindTelefone = []
            sleep(1)
            ultima_pag = driver.find_element_by_xpath('/html/body/form/div[3]/div/span[1]/span/div/div[5]/div/table/tbody/tr/td[2]/table/tbody/tr/td[4]/span').get_attribute("innerHTML")
            ultima_pag = int(ultima_pag)
            pag_atual = 1
            limite = 4
            while (pag_atual <= ultima_pag):
                for x in range(1,lenTelefone):
                    numeroTelefone = EF.getText(driver,'/html/body/form/div[3]/div/span[1]/span/div/div[3]/div[3]/div/table/tbody/tr[{}]/td[3]'.format(x),insist,'innerText').lower()

                    telefonevalida = numeroTelefone == telefone and numeroTelefone != '' or numeroTelefone == telefone.replace('-','') and numeroTelefone != ''
                    telefonevazio = numeroTelefone == ''

                    celularvalida = celular == numeroTelefone or celular.replace('-','') == numeroTelefone
                    if telefonevalida or celularvalida:
                        listFindTelefone.append(True)
                    else:
                        if telefonevazio:
                            continue
                        else:
                            listFindTelefone.append(False)

                if(pag_atual != ultima_pag):
                    xpath = '/html/body/form/div[3]/div/span[1]/span/div/div[5]/div/table/tbody/tr/td[2]/table/tbody/tr/td[6]'
                    btn_prox =  EF.click(driver,xpath,insist)
                    sleep(1)
                    pag_atual += 1
                else:
                    break   

            if True not in listFindTelefone:
                print('fazer validação para inserir telefone')
            
                driver.switch_to.default_content()
                driver.switch_to.frame('Cliente1')
                driver.execute_script("incluirClienteTelefone1();")

                driver.switch_to.frame('ZonaInterna')
                if ddd == '':
                    EF.fill(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[1]/div/div[1]/div/input',dddc,insist,1)
                else:
                    EF.fill(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[1]/div/div[1]/div/input',ddd,insist,1)

                if telefone == '':
                    EF.fill(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[1]/div/div[2]/div/input',celular,insist)
                else:
                    EF.fill(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[1]/div/div[2]/div/input',telefone,insist)

                EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[1]/div/div[4]/div/span/span[1]/span/span[1]',insist)

                EF.click(driver,'/html/body/span[2]/span/span[2]/ul/li[12]',insist)

                driver.switch_to.default_content()
                driver.switch_to.frame('Cliente1')

                driver.execute_script("evento('Gravar')")

            driver.switch_to.default_content()
            sleep(5)
            alert = '/html/body/div[3]/div'
            textoAlert = '/html/body/div[5]/div/div[2]'
            if driver.find_elements(By.XPATH,alert):
                try:
                    textAlert = driver.find_element(By.XPATH,textoAlert).get_attribute('innerText').lower()
                    if 'já possui uma proposta' in textAlert:
                        # raise ValueError('Erro - endosso já possui proposta')
                        EF.click(driver,'/html/body/div[3]/div/div[2]/div[1]/div/div/button[1]',insist)
        
                    elif 'nível corretora' in textAlert:
                            EF.click(driver,'/html/body/div[3]/div/div[2]/div[1]/div/div/button[1]',insist)
                            EF.click(driver,'/html/body/div[5]/div/div[2]/div[1]/div/div/button[1]',insist)
                            sleep(3)
                    elif 'cadastro de clientes' in textAlert:
                        pass
                    else:
                        EF.click(driver,'/html/body/div[5]/div/div[3]/button[1]',insist)
                except:
                    pass

            driver.switch_to.frame('Cliente1')
            EF.click(driver,'//*[@id="Button1"]',insist)
            driver.switch_to.default_content()
            alertnovo = '/html/body/div[6]/div'
            alert = '/html/body/div[3]/div'
            if driver.find_elements(By.XPATH,alertnovo):
                EF.click(driver, '/html/body/div[6]/div/div[3]/button[1]', insist)
            if driver.find_elements(By.XPATH,alert):
                EF.click(driver, '/html/body/div[5]/div/div[3]/button[1]', insist)
                
        except ValueError as err:
            raise ValueError(err)
        except Exception as err:
            raise Exception('Erro - Inserir dados cadastrais')
        
    def CadastroCoberturas(driver,wait,rs):
        try:
            sleep(2)
            driver.switch_to.default_content()
            try:
                ExperienciaCliente = '/html/body/div[3]'
                if driver.find_element(By.XPATH,ExperienciaCliente):
                    EF.click(driver,'/html/body/div[3]/div[1]/a/span',insist)
            except:pass
            EF.acessarframe(driver, 'ZonaInterna',2)
            driver.switch_to.frame('Documento')
            driver.switch_to.frame('ZonaInterna')

            #Clica em coberturas
            driver.execute_script("expandirComprimeCard('BoxDocumentoItemVeiculoCob');openBoxDocumentoItemVeiculoCob();")
            driver.switch_to.default_content()
            try:
                ExperienciaCliente = '/html/body/div[3]'
                if driver.find_element(By.XPATH,ExperienciaCliente):
                    EF.click(driver,'/html/body/div[3]/div[1]/a/span',insist)
            except:pass
            EF.acessarframe(driver, 'ZonaInterna',2)
            driver.switch_to.frame('Documento')
            driver.switch_to.frame('ZonaInterna')
            sleep(2)
            driver.switch_to.frame('FrameDocumentoItemVeiculoCob')
            
            Findelements = 0
            while Findelements == 0:
                driver.switch_to.default_content()
                try:
                    ExperienciaCliente = '/html/body/div[3]'
                    if driver.find_element(By.XPATH,ExperienciaCliente):
                        EF.click(driver,'/html/body/div[3]/div[1]/a/span',insist)
                except:pass
                EF.acessarframe(driver, 'ZonaInterna',2)
                driver.switch_to.frame('Documento')
                driver.switch_to.frame('ZonaInterna')
                driver.switch_to.frame('FrameDocumentoItemVeiculoCob')
                elemBtnExluir = driver.find_elements(By.ID,'BtExcluir')
                if len(elemBtnExluir) != 0:
                    sleep(1)
                    try:
                        elemBtnExluir[0].click()
                    except:
                        driver.switch_to.default_content()
                        try:
                            ExperienciaCliente = '/html/body/div[3]'
                            if driver.find_element(By.XPATH,ExperienciaCliente):
                                EF.click(driver,'/html/body/div[3]/div[1]/a/span',insist)
                        except:pass
                        EF.acessarframe(driver, 'ZonaInterna',2)
                        driver.switch_to.frame('Documento')
                        driver.switch_to.frame('ZonaInterna')
                        driver.switch_to.frame('FrameDocumentoItemVeiculoCob')
                        if len(elemBtnExluir) != 0:
                            sleep(1)
                            elemBtnExluir[0].click()
                        else:
                            Findelements = 1
                else:
                    Findelements = 1
                driver.switch_to.default_content()
                alertnovo = '/html/body/div[4]/div'
                alert = '/html/body/div[3]/div'
                try:
                    ExperienciaCliente = '/html/body/div[3]'
                    if driver.find_element(By.XPATH,ExperienciaCliente):
                        EF.click(driver,'/html/body/div[3]/div[1]/a/span',insist)
                except:pass
                try:
                    if driver.find_elements(By.XPATH,alert):
                        EF.click(driver, '/html/body/div[3]/div/div[3]/button[1]', insist)
                    if driver.find_elements(By.XPATH,alertnovo):
                        EF.click(driver,'/html/body/div[4]/div/div[3]/button[1]',insist)
                except:pass
                
                EF.acessarframe(driver, 'ZonaInterna',2)
                driver.switch_to.frame('Documento')
                driver.switch_to.frame('ZonaInterna')
                sleep(3)
                driver.switch_to.frame('FrameDocumentoItemVeiculoCob')
            sleep(2)
            driver.switch_to.default_content()
            EF.acessarframe(driver, 'ZonaInterna',2)
            driver.switch_to.frame('Documento')
            driver.switch_to.frame('ZonaInterna')
            
            #Clica em incluir coberturas
            driver.execute_script("incluirDocumentoItemVeiculoCob();")
            sleep(2)
            driver.switch_to.default_content()
            EF.acessarframe(driver, 'ZonaInterna',2)
            driver.switch_to.frame('Documento')
            driver.switch_to.frame('ZonaInterna')
            driver.switch_to.frame('ZonaInterna')
            chave = False
            for x in listCoberturas:
                nomeCobertura = x[2]
                valor_ajuste_Cobertura = x[3].replace('r$','').replace('*','').strip()
                if 'franquia' in nomeCobertura and 'casco' in nomeCobertura:
                    nomeCoberturasplit = nomeCobertura.split(' ')
                    nomeCobertura = nomeCoberturasplit[1]
                findCobertura = 0
                if chave == False:
                    wait.until(EC.visibility_of_element_located((By.ID,'select2-DocumentoIteCob_Cobertura-container'))).click()
                # driver.find_element(By.ID,'select2-DocumentoIteCob_Cobertura-container').click()
                EF.fill(driver,'/html/body/span[2]/span/span[1]/input',nomeCobertura,insist,1)
                sleep(1)
                listItens = driver.find_elements(By.XPATH,'/html/body/span[2]/span/span[2]/ul/li')

                for item in listItens:
                    itemValue = item.get_attribute('innerText')
                    #print(itemValue)
                    if itemValue.lower() == nomeCobertura and valor_ajuste_Cobertura != '0,00' and valor_ajuste_Cobertura != '0,0':
                        item.click()
                        findCobertura = 1
                        chave = False
                        break
                    else:
                        driver.find_element(By.XPATH,'/html/body/span[2]/span/span[1]/input').clear()
                        chave = True

                    if 'Nenhum resultado encontrado' in itemValue:
                        conn = pymssql.connect(server=GC.server, user=GC.user, password=GC.password, database=GC.database)
                        cursor = conn.cursor()
                        data_atual = date.today()
                        path_to_save =   f'Evidencia-{rs}-Endosso-Erro-Cobertura-{data_atual}.png'
                        driver.save_screenshot(GC.evidencia_path + path_to_save)
                        insert = f'''
                                    insert into coberturas_erro_log(
                                    rs
                                    ,erro_name
                                    ,erro_path
                                    )VALUES(
                                    {rs}
                                    ,'Incluir cobertura'
                                    ,'{path_to_save}'
                                    )
                                    '''
                        cursor.execute(insert)
                        conn.commit()
                        conn.close()

                if findCobertura != 0: 
                    sleep(3)
                    if 'casco' not in nomeCobertura.lower():
                        #Insere valor da cobertura
                        if ',' in valor_ajuste_Cobertura:
                            driver.execute_script(f"document.getElementById('DocumentoIteCob_ImpSegurada').setAttribute('value', '{valor_ajuste_Cobertura}')")
                        else:
                            driver.execute_script(f"document.getElementById('DocumentoIteCob_ImpSegurada').setAttribute('value', '0,0')")
                        sleep(3)

                        if 'vidro' in nomeCobertura.lower():
                            EF.fill(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[3]/div[1]/div/div[2]/div[3]/div/div/textarea',detalhamento_franquia_vidro,insist)
                    else:
                        #Insere ajuste
                        if seguradora.lower() == 'porto seguro cia' or seguradora.lower() == 'porto':
                            EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[3]/div[1]/div/div[2]/div[1]/div/div/input',insist)
                            sleep(3)
                            # #tipo de franquia
                            EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[3]/div[1]/div/div[2]/div[2]/div[1]/div/span/span[1]/span/span[1]',insist)
                            selectable = driver.find_elements(By.XPATH,'/html/body/span[2]/span/span[2]/ul/li')
                            for i in range(0,len(selectable)):
                                option1 = EF.getText(driver,f'/html/body/span[2]/span/span[2]/ul/li[{i}]',insist,'innerText')
                            
                                option1 = option1.split(' ')
                                option1 = option1[0]
                                if option1.lower() == 'obrigatoria':
                                    option1 = 'obrigatória'
                                if option1.lower() == franquia_reduzida.lower():
                                    EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[3]/div[1]/div/div[2]/div[2]/div[2]/div/input',insist)
                                    selectTipoFranquia = Select(driver.find_element(By.ID,'DocumentoIteCob_TipoFranquia'))
                                    i = i-1
                                    selectTipoFranquia.select_by_index(i)
                                    break
                            sleep(3)
                            EF.fill(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[3]/div[1]/div/div[2]/div[2]/div[2]/div/input',valor_ajuste_Cobertura,insist,1)
                        else:
                            EF.fill(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[3]/div/input',valor_ajuste_Cobertura,insist,1)
                            sleep(3)
                            #Seleciona 'tem franquia'
                            EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[3]/div[1]/div/div[2]/div[1]/div/div/input',insist)
                            sleep(3)
                            #tipo de franquia
                            selectTipoFranquia = Select(driver.find_element(By.ID,'DocumentoIteCob_TipoFranquia'))
                            try:
                                if '50' in franquia_reduzida:
                                    selectTipoFranquia.select_by_index(27)
                                elif '75' in franquia_reduzida:
                                    selectTipoFranquia.select_by_index(23)
                                elif '25' in franquia_reduzida:
                                    selectTipoFranquia.select_by_index(34)
                            except:pass
                            sleep(3)
                            #Insere valor da franquia
                            EF.fill(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[3]/div[1]/div/div[2]/div[2]/div[2]/div/input',franquia_reduzida,insist,1)
                            sleep(3)

                    driver.execute_script("evento('Gravar')")
                    driver.execute_script("evento('Incluir')")
                
            driver.execute_script("evento('Voltar')")
            driver.switch_to.default_content()
            sleep(3)
            try:
                try:
                    if driver.find_element(By.XPATH,'/html/body/div[4]/div'):
                        EF.click(driver,'/html/body/div[4]/div/div[3]/button[1]',insist)
                except:pass
                if driver.find_element(By.XPATH,'/html/body/div[3]/div/div[3]/button[1]'):
                    EF.click(driver,'/html/body/div[3]/div/div[3]/button[1]',insist)
            except:pass
        except ValueError as err:
            raise ValueError(err)
        except Exception as err:
            raise Exception('Erro - Incluir cobertura')

    def CadastroCondutor(driver,wait,rs):
        try:            
            driver.switch_to.default_content()
            EF.acessarframe(driver, 'ZonaInterna',2)
            driver.switch_to.frame('Documento')
            driver.switch_to.frame('ZonaInterna')

            #clica em condutores
            sleep(2)
            driver.execute_script("expandirComprimeCard('BoxDocsItensAutoCond');openBoxDocsItensAutoCond();")
        
            driver.switch_to.frame('FrameDocsItensAutoCond')
            sleep(2)

            #clica em editar condutor
            if '0km' in Modelo.lower():
                driver.switch_to.default_content()
                driver.switch_to.frame('ZonaInterna')
                driver.switch_to.frame('ZonaInterna')
                driver.switch_to.frame('Documento')
                driver.switch_to.frame('ZonaInterna')
                driver.execute_script("incluirDocsItensAutoCond();")
                sleep(2)
                if segurado.lower() == nomeCondutor.lower():
                    driver.switch_to.default_content()
                    EF.click(driver,'/html/body/div[4]/div/div[2]/div[1]/div/div/button[1]',insist)
                    try:
                        ExperienciaCliente = '/html/body/div[3]'
                        if driver.find_element(By.XPATH,ExperienciaCliente):
                            EF.click(driver,'/html/body/div[3]/div[1]/a/span',insist)
                    except:pass
                    chave = False
                    try:
                        if driver.find_elements(By.XPATH,'/html/body/div[3]/div'):
                            EF.click(driver,'/html/body/div[3]/div/div[2]/div[1]/div/div/button[1]', insist)
                            chave = True
                    except:pass
                    driver.switch_to.frame('ZonaInterna')
                    driver.switch_to.frame('ZonaInterna')
                    driver.switch_to.frame('Documento')
                    driver.switch_to.frame('ZonaInterna')
                    driver.switch_to.frame('ZonaInterna')
                    if chave ==False:
                        EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[1]/div[2]/div/div/input',insist)
                        casado = 'casado' in estadoCivilCondutor.lower()
                        solteiro = 'solteiro' in estadoCivilCondutor.lower()
                        viuvo = 'viúvo' in estadoCivilCondutor.lower()
                        desquitado = 'desquitado' in estadoCivilCondutor.lower()
                        divorciado = 'divorciado' in estadoCivilCondutor.lower()
                        sleep(2)
                        EF.click(driver,'/html/body/div[3]/div/div[2]/div[1]/div/div/button[2]', insist)
                        sleep(1)
                        EF.fill(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[1]/div[1]/div/input',nomeCondutor,insist,1)
                        elem = driver.find_element(By.NAME,"DocumentoIteCond_Idade")
                        elem.send_keys(idadeCondutor)
                        sleep(2)
                        if casado:
                            EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[4]/div/div/label[1]/input',insist)
                        elif solteiro:
                            EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[4]/div/div/label[2]/input',insist)
                        elif viuvo:
                            EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[4]/div/div/label[3]/input',insist)
                        elif desquitado:
                            EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[4]/div/div/label[4]/input',insist)
                        elif divorciado:
                            EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[4]/div/div/label[5]/input',insist)
                        else:
                            EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[4]/div/div/label[6]/input',insist)
                        EF.fill(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[3]/div[1]/div/input',cpfCondutor,insist)
                        if nomeCondutor.lower() != segurado.lower():
                            EF.fill(driver, '/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[4]/div[1]/div/input', 'Não Informado', insist, 1)
                        elif nomeCondutor.lower() == segurado.lower():
                            EF.fill(driver, '/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[4]/div[1]/div/input', 'O Próprio', insist, 1)
                        
                        sleep(2)
                        
                        #diz o genero do condutor
                        if 'masculino' in genero:
                            EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[3]/div/div/label[1]/input',insist)
                        elif 'feminino' in genero:
                            EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[3]/div/div/label[2]/input',insist)
                else:
                    driver.switch_to.default_content()
                    EF.click(driver,'/html/body/div[3]/div/div[2]/div[1]/div/div/button[2]',insist)
                    try:
                        ExperienciaCliente = '/html/body/div[3]'
                        if driver.find_element(By.XPATH,ExperienciaCliente):
                            EF.click(driver,'/html/body/div[3]/div[1]/a/span',insist)
                    except:pass
                    sleep(1)
                    EF.acessarframe(driver, 'ZonaInterna',2)
                    driver.switch_to.frame('Documento')
                    driver.switch_to.frame('ZonaInterna')
                    driver.switch_to.frame('ZonaInterna')
                    sleep(2)
                    EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[1]/div[2]/div/div/input',insist)
                    
                    casado = 'casado' in estadoCivilCondutor.lower()
                    solteiro = 'solteiro' in estadoCivilCondutor.lower()
                    viuvo = 'viúvo' in estadoCivilCondutor.lower()
                    desquitado = 'desquitado' in estadoCivilCondutor.lower()
                    divorciado = 'divorciado' in estadoCivilCondutor.lower()
                    EF.click(driver,'/html/body/div[3]/div/div[2]/div[1]/div/div/button[2]', insist)
                    sleep(1)
                    EF.fill(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[1]/div[1]/div/input',nomeCondutor,insist,1)
                    elem = driver.find_element(By.NAME,"DocumentoIteCond_Idade")
                    elem.send_keys(idadeCondutor)
                    if casado:
                        EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[4]/div/div/label[1]/input',insist)
                    elif solteiro:
                        EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[4]/div/div/label[2]/input',insist)
                    elif viuvo:
                        EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[4]/div/div/label[3]/input',insist)
                    elif desquitado:
                        EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[4]/div/div/label[4]/input',insist)
                    elif divorciado:
                        EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[4]/div/div/label[5]/input',insist)
                    else:
                        EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[4]/div/div/label[6]/input',insist)
                    sleep(2)
                    if nomeCondutor.lower() != segurado.lower():
                        EF.fill(driver, '/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[4]/div[1]/div/input', 'Não Informado', insist, 1)
                    elif nomeCondutor.lower() == segurado.lower():
                        EF.fill(driver, '/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[4]/div[1]/div/input', 'O Próprio', insist, 1)
                    #diz o genero do condutor
                    EF.fill(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[3]/div[1]/div/input',cpfCondutor,insist)
                    if 'masculino' in genero:
                            EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[3]/div/div/label[1]/input',insist)
                    elif 'feminino' in genero:
                        EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[3]/div/div/label[2]/input',insist)
                    sleep(2)
                   
            else:
                verificatexto = EF.getText(driver,'/html/body/form/div[3]/div/span[1]/span/div/div[5]/div/table/tbody/tr/td[1]/div')
                if verificatexto == 'Nenhum registro':
                    driver.switch_to.default_content()
                    driver.switch_to.frame('ZonaInterna')
                    driver.switch_to.frame('ZonaInterna')
                    driver.switch_to.frame('Documento')
                    driver.switch_to.frame('ZonaInterna')
                    driver.execute_script("incluirDocsItensAutoCond();")
                    if segurado.lower() == nomeCondutor.lower():
                        driver.switch_to.default_content()
                        EF.click(driver,'/html/body/div[4]/div/div[2]/div[1]/div/div/button[1]',insist)
                    else:
                        driver.switch_to.default_content()
                        EF.click(driver,'/html/body/div[3]/div/div[2]/div[1]/div/div/button[2]',insist)
                else:
                    sleep(5)
                    editar = driver.find_element(By.XPATH,'/html/body/form/div[3]/div/span[1]/span/div/div[3]/div[3]/div/table/tbody/tr[2]/td[1]/a')
                    driver.execute_script("arguments[0].click();", editar)
                    sleep(2)
                driver.switch_to.default_content()
                EF.acessarframe(driver, 'ZonaInterna',2)
                driver.switch_to.frame('Documento')
                driver.switch_to.frame('ZonaInterna')
                driver.switch_to.frame('ZonaInterna')
                sleep(2)
                wait.until(EC.visibility_of_element_located((By.XPATH,'/html/body/form/div[7]/div/div[2]/div/div')))
                sleep(2)
                #Nome Condutor
                EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[1]/div[2]/div/div/input',insist)

                sleep(2)
                driver.execute_script(f"document.getElementById('DocumentoIteCond_Nome').setAttribute('value', '{nomeCondutor}')")

               
                #Idade
                sleep(2)
                
                EF.click(driver, '/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[2]/div/input', insist)
                sleep(2)
                EF.fill(driver, '/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[2]/div/input', idadeCondutor, insist, 1)
                sleep(2)

                #Estado Civil
                allEstadoCivil = driver.find_elements(By.XPATH, '/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[4]/div/div/label')
                lenEstadoCivil = len(allEstadoCivil)+1
                for x in range(1,lenEstadoCivil):
                    valueOption = EF.getText(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[4]/div/div/label[{}]/label'.format(x),insist,'innerText')
                    if valueOption.lower() in estadoCivilCondutor:
                        sleep(2)
                        EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[4]/div/div/label[{}]/input'.format(x), insist)
                        break
                #Cpf
                sleep(2)
                driver.execute_script(f"document.getElementById('DocumentoIteCond_CnpjCpf').setAttribute('value', '{cpfCondutor}')")
                sleep(3)

                #Parentesco ou relacionamento com o condutor principal
                if nomeCondutor.lower() != segurado.lower():
                    EF.fill(driver, '/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[4]/div[1]/div/input', 'Não Informado', insist, 1)
                elif nomeCondutor.lower() == segurado.lower():
                    EF.fill(driver, '/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[4]/div[1]/div/input', 'O Próprio', insist, 1)
                #diz o genero do condutor
                sleep(2)
                if 'masculino' in genero.lower():
                    EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[3]/div/div/label[1]/input',insist)
                elif 'feminino' in genero.lower():
                    EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div/div[2]/div[2]/div[3]/div/div/label[2]/input',insist)

            #Clica em gravar
            sleep(2)
            driver.execute_script("evento('Gravar')")
            sleep(3)
            driver.execute_script("evento('Incluir')")

            driver.execute_script("evento('Voltar')")

            driver.switch_to.default_content()
            
            alert = '/html/body/div[3]/div'
            alertnovo = '/html/body/div[4]/div'
            sleep(3)
            try:
                if driver.find_elements(By.XPATH,alertnovo):
                    EF.click(driver,'/html/body/div[4]/div/div[3]/button[1]',insist)
                if driver.find_elements(By.XPATH,alert):
                    EF.click(driver, '/html/body/div[3]/div/div[3]/button[1]', insist)
            except:pass
        except ValueError as err:
            raise ValueError(err)
        except Exception as err:
            raise Exception('Erro - Incluir condutor')

    def CadastroProduto(driver,wait,rs):
        try:
            if 'cancalemento' in tipoEndosso.lower():
                return 1
            if 'itens' not in subtipoEndosso.lower():
                return 1

            sleep(2)
            driver.execute_script('incluirDocumentoItemVeiculo();')
            driver.switch_to.default_content()
            alertnovo = '/html/body/div[4]/div'
            try:
                if driver.find_elements(By.XPATH,alertnovo):
                    EF.click(driver,'/html/body/div[4]/div/div[3]/button[1]',insist)
            except:pass
            try:
                ExperienciaCliente = '/html/body/div[3]'
                if driver.find_element(By.XPATH,ExperienciaCliente):
                    EF.click(driver,'/html/body/div[3]/div[1]/a/span',insist)
            except:pass
            EF.acessarframe(driver, 'ZonaInterna',2)
            driver.switch_to.frame('Documento')
            driver.switch_to.frame('ZonaInterna')

            sleep(2)    
            EF.click(driver, '//*[@id="TEMIMPORTA"]', insist)
            sleep(5)

            try:
                obj = driver.find_element_by_xpath('/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[1]/div/div/span/div[1]/div[3]/div[3]/div/table/tbody/tr[2]')
                actions = ActionChains(driver)
                actions.double_click(obj).perform()
            except:
                pass

            ##Desmarcar 0km enquanto não temos numero da nota fiscal
            sleep(3)
            ValidarVeiculo = False
            try:
                if driver.find_element(By.XPATH,'/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[3]/div[3]/div[2]/div/input').is_selected():
                    EF.click(driver, '//*[@id="DocumentoIte_ZeroKm"]', insist)
                    ValidarVeiculo = True
                if '0km' in Modelo.lower():
                    sleep(2)
                    EF.click(driver, '//*[@id="DocumentoIte_ZeroKm"]', insist)
                    ValidarVeiculo = True
            except:
                 if ValidarVeiculo == False:
                    if '0km' in Modelo.lower():
                        sleep(2)
                        EF.click(driver, '//*[@id="DocumentoIte_ZeroKm"]', insist)
                        ValidarVeiculo = True
                 pass

            #Fabricante
            EF.click(driver, '//*[@id="select2-DocumentoIte_Fabricante-container"]', insist)
            sleep(2)
            EF.fill(driver, '/html/body/span[2]/span/span[1]/input',Fabricante,insist)
            sleep(2)
            EF.click(driver, '/html/body/span[2]/span/span[2]/ul/li', insist)

            #Modelo
            sleep(2)
            EF.click(driver, '//*[@id="DocumentoIte_DescrModelo"]', insist)
            sleep(2)
            EF.fill(driver, '//*[@id="DocumentoIte_DescrModelo"]',Modelo,insist)
            sleep(1)
           
            flex = 'flex' in Modelo.lower()
            gasolina = combustivel.lower()== 'gasolina'
            diesel = combustivel.lower() == 'diesel' 
            alcool = combustivel.lower() == 'álcool'
            eletrico = combustivel.lower() == 'elétrico' 
            tetrafuel = combustivel.lower() == 'tetrafuel' 
            gas = combustivel.lower() == 'gas' 
            sleep(2)
            if '0km' in Modelo.lower():
                EF.fill(driver,'/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[4]/div[6]/div/input',nota_fiscal,insist)
            if flex:
                driver.execute_script("arguments[0].click();",wait.until(EC.element_to_be_clickable((By.XPATH,'/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[3]/div[6]/div/div/label[5]/input'))))
            elif gasolina:
               driver.execute_script("arguments[0].click();",wait.until(EC.element_to_be_clickable((By.XPATH,'/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[3]/div[6]/div/div/label[1]/input'))))
            elif diesel:
                driver.execute_script("arguments[0].click();",wait.until(EC.element_to_be_clickable((By.XPATH,'/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[3]/div[6]/div/div/label[3]/input'))))
            elif alcool:
                driver.execute_script("arguments[0].click();",wait.until(EC.element_to_be_clickable((By.XPATH,'/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[3]/div[6]/div/div/label[2]/input'))))
            elif eletrico:
                driver.execute_script("arguments[0].click();",wait.until(EC.element_to_be_clickable((By.XPATH,'/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[3]/div[6]/div/div/label[10]/input'))))
            elif tetrafuel:
                driver.execute_script("arguments[0].click();",wait.until(EC.element_to_be_clickable((By.XPATH,'/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[3]/div[6]/div/div/label[11]/input'))))
            elif gas:
                driver.execute_script("arguments[0].click();",wait.until(EC.element_to_be_clickable((By.XPATH,'/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[3]/div[6]/div/div/label[4]/input'))))
            else:
               driver.execute_script("arguments[0].click();",wait.until(EC.element_to_be_clickable((By.XPATH,'/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[3]/div[6]/div/div/label[1]/input'))))
            sleep(1)
            try:
                anoFab, anoMod = ano.split(',')
            except:
                anoFab, anoMod = ano.split('/')
            validacaoanofab = driver.find_element(By.XPATH,'/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[3]/div[4]/div/input')
            validacaoanofab1 = validacaoanofab.get_attribute('value')
            sleep(2)
            driver.execute_script(f"document.getElementById('DocumentoIte_AnoFabricacao').setAttribute('value', '{anoFab}')")
            validacaoanofab2 = driver.find_element(By.XPATH,'/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[3]/div[4]/div/input')
            validacaoanofab3 = validacaoanofab2.get_attribute('value')
            

            if validacaoanofab1 == validacaoanofab3 and anoFab != anoMod or validacaoanofab3 != anoFab:
                try:
                    driver.find_element(By.XPATH,'/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[3]/div[4]/div/input').clear()
                    sleep(2)
                    EF.click(driver,'/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[3]/div[4]/div/input',anoFab,insist)
                    sleep(2)
                except:
                    driver.find_element(By.XPATH,'/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[3]/div[4]/div/input').clear()
                    sleep(2)
                    driver.execute_script(f'document.getElementById("DocumentoIte_AnoFabricacao").value="{anoFab}" ')
                    sleep(3)
                    validacaoanofab4 = driver.find_element(By.XPATH,'/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[3]/div[4]/div/input')
                    validacaoanofab5 = validacaoanofab4.get_attribute('value')
                    if validacaoanofab5 == '0':
                        driver.find_element(By.XPATH,'/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[3]/div[4]/div/input').clear()
                        EF.fill(driver,'/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[3]/div[4]/div/input',anoFab,insist)
            sleep(2)

            EF.click(driver, '//*[@id="DocumentoIte_AnoModelo"]', insist)

            time.sleep(defaultsleep)

            EF.fill(driver, '//*[@id="DocumentoIte_AnoModelo"]',anoMod, insist,1)
            sleep(2)
            EF.fill(driver, '//*[@id="DocumentoIte_Placa"]',Placa, insist)
            sleep(2)
            EF.fill(driver, '//*[@id="DocumentoIte_Chassi"]',Chassi, insist)
            sleep(2)   
            driver.find_element(By.XPATH,'//*[@id="DocumentoIte_Renavam"]').clear()
            sleep(1)
            if seguradora == 'Allianz':
                driver.find_element(By.XPATH,'//*[@id="DocumentoIte_Renavam"]').clear()
            else:
                if renavam != '':
                    driver.find_element(By.XPATH,'//*[@id="DocumentoIte_Renavam"]').clear()
                    sleep(2)
                    driver.find_element(By.XPATH,'//*[@id="DocumentoIte_Renavam"]').send_keys(renavam)

            EF.fill(driver, '//*[@id="DocumentoIte_CodFipe"]',FIPE, insist)
            sleep(2)

            sleep(1)
            if '0km' in Modelo.lower():
                EF.fill(driver,'//*[@id="DocumentoIte_DataSaida"]',data_saida,insist)

            if seguradora == 'Allianz':
                EF.fill(driver, '//*[@id="DocumentoIte_NotaFiscal"]', 'A/C', insist)

            #Categoria
            EF.click(driver, '/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[5]/div[5]/div/span', insist)
          
            EF.fill(driver, '/html/body/span[2]/span/span[1]/input', categoria,insist)
            EF.click(driver, '/html/body/span[2]/span/span[2]/ul/li', insist)

            EF.fill(driver, '//*[@id="DocumentoIte_Cep"]',CEPRISCO, insist)
                                
            #Cidade
            driver.execute_script("acionaJanelaDocumentoIte_AuxCodCidade();")
            sleep(5)
            driver.switch_to.frame('SearchCidades3')
    
            chavelidacao = False
            chave = False
            wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="Nome"]'))).send_keys(cidade)
            sleep(1)
            EF.click(driver, '//*[@id="SpanToolBarRightB"]/button', insist)
            sleep(1)
            alertnovo = '/html/body/div[4]/div'
            alert = '/html/body/div[3]/div'
            try:
                driver.switch_to.default_content()
                if driver.find_element(By.XPATH,'/html/body/div[3]/div'):
                    sleep(1)
                    try:
                        driver.find_element(By.XPATH,'/html/body/div[3]/div/div[3]/button[1]').click()
                        chave = True
                        sleep(2)
                        EF.acessarframe(driver, 'ZonaInterna',2)
                        driver.switch_to.frame('Documento')
                        driver.switch_to.frame('ZonaInterna')
                        EF.click(driver,'/html/body/form/div[11]/div/div/div[1]/button/span',insist)
                        chavelidacao = True
                    except:
                        if driver.find_element(By.XPATH,'/html/body/div[4]/div'):
                            sleep(1)
                            driver.find_element(By.XPATH,'/html/body/div[4]/div/div[3]/button[1]').click()
                            chave = True
                            sleep(2)
                            EF.acessarframe(driver, 'ZonaInterna',2)
                            driver.switch_to.frame('Documento')
                            driver.switch_to.frame('ZonaInterna')
                            EF.click(driver,'/html/body/form/div[11]/div/div/div[1]/button/span',insist)
                            chavelidacao = True
            except:pass
            if chave == False:
                EF.acessarframe(driver, 'ZonaInterna',2)
                driver.switch_to.frame('Documento')
                driver.switch_to.frame('ZonaInterna')
                driver.switch_to.frame('SearchCidades3')
            if chavelidacao == False:
                xpathTr = '/html/body/form/div[5]/div[2]/div[2]/div/div[3]/div[3]/div/table/tbody/tr'
                wait.until(EC.visibility_of_element_located((By.XPATH,'/html/body/form/div[5]/div[2]/div[2]/div/div[3]/div[3]/div/table')))
                teste = driver.find_elements(By.XPATH,xpathTr)
                lenTeste = len(teste)+1
                for x in range(1,lenTeste):
                    estadoRegistro = driver.find_element(By.XPATH,'/html/body/form/div[5]/div[2]/div[2]/div/div[3]/div[3]/div/table/tbody/tr[{}]/td[2]'.format(x)).get_attribute('innerText').lower()

                    if estadoRegistro == estado.lower():
                        EF.click(driver,'/html/body/form/div[5]/div[2]/div[2]/div/div[3]/div[3]/div/table/tbody/tr[{}]/td[1]/a'.format(x),insist)
                        break

            driver.switch_to.default_content()
            EF.acessarframe(driver, 'ZonaInterna',2)
            driver.switch_to.frame('Documento')
            driver.switch_to.frame('ZonaInterna')
            sleep(4)
            driver.execute_script(f'document.getElementById("DocumentoIte_Cep").value="{cep_pernoite}" ')
            #Classe
            
            if 'classe' in classeBonus:
                classe = classeBonus
            else:
                classe = 'classe {}'.format(classeBonus)
            wait.until(EC.visibility_of_element_located((By.XPATH,'/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[7]/div[3]')))
            EF.click(driver, '/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[7]/div[3]/div/span/span[1]/span/span[1]', insist)
            EF.fill(driver, '/html/body/span[2]/span/span[1]/input', classe, insist)
            EF.click(driver, '/html/body/span[2]/span/span[2]/ul/li[1]', insist)

            #Cód. identificação
            EF.fill(driver, '/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[8]/div[5]/div/input', codCI, insist,1)

            #Data substituição 
            dataSubs = inicioVigencia
            EF.click(driver, '/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[10]/div[3]/div/div/input', insist)
            EF.fill(driver, '/html/body/form/div[7]/div/div[2]/div/div/div[2]/div/div[10]/div[3]/div/div/input', dataSubs, insist, 1)
            sleep(3)

            #Clica em gravar
            sleep(2)
            driver.execute_script("evento('Gravar')")
            sleep(3)

            CadastroCoberturas(driver,wait,rs)

            CadastroCondutor(driver,wait,rs)

            driver.switch_to.default_content()
            try:
                ExperienciaCliente = '/html/body/div[3]'
                if driver.find_element(By.XPATH,ExperienciaCliente):
                    EF.click(driver,'/html/body/div[3]/div[1]/a/span',insist)
            except:pass
            driver.switch_to.frame('ZonaInterna')
            driver.switch_to.frame('ZonaInterna')
            driver.switch_to.frame('Documento')
            driver.switch_to.frame('ZonaInterna')
            sleep(2)
            driver.execute_script("evento('Incluir')")

            driver.execute_script("evento('Voltar')")

            driver.switch_to.default_content()
            alert = '/html/body/div[3]/div'
            alertnovo = '/html/body/div[4]/div'
            sleep(3)
            try:
                try:
                    if driver.find_elements(By.XPATH,alertnovo):
                        EF.click(driver,'/html/body/div[4]/div/div[3]/button[1]',insist)
                except:pass
                if driver.find_elements(By.XPATH,alert):
                    EF.click(driver, '/html/body/div[3]/div/div[3]/button[1]', insist)
            except:pass
            driver.switch_to.frame('ZonaInterna')
            driver.switch_to.frame('ZonaInterna')
            driver.switch_to.frame('Documento')
            pass

            return chave

        except ValueError as err:
            raise ValueError(err)
     
        except Exception as err:
            raise Exception('Erro - Substituir Veiculo')

    def cadastraConsultor(driver):
        driver.switch_to.default_content()  
        driver.switch_to.frame('ZonaInterna')
        driver.switch_to.frame('ZonaInterna')
        driver.switch_to.frame('Documento')

        driver.execute_script('AbreProdutores();')
        driver.switch_to.default_content()
        driver.switch_to.frame('DocumentoRepasse')

        ############### Início do script do LIRA ###################
        ultima_pag = driver.find_element_by_xpath('//*[@id="sp_1_BAR_GridCadastro2"]').get_attribute("innerHTML")
        ultima_pag = int(ultima_pag)
        pag_atual = 1
        limite = 4
        sleep(2)
        while (pag_atual <= ultima_pag):
            for x in range(2, 6):
                try:
                    sleep(2)
                    td_consultor = driver.find_element_by_xpath(f'/html/body/form/div[3]/div/span[2]/span/div/div[3]/div[3]/div/table/tbody/tr[{x}]/td[1]')
                    td_consultor = td_consultor.get_attribute("innerHTML")
        
                    if(td_consultor == 'CONSULTOR'):
                        sleep(2)
                        xpath = f'/html/body/form/div[3]/div/span[2]/span/div/div[3]/div[3]/div/table/tbody/tr[{x}]/td[4]/a'
                        btn_excluir = EF.click(driver,xpath,insist)
                        sleep(2)
                        driver.switch_to.default_content()
                        alertnovo = '/html/body/div[6]/div'
                        alertantigo = '/html/body/div[5]/div'
                        try:
                            if driver.find_elements(By.XPATH,alertnovo):
                                EF.click(driver,'/html/body/div[6]/div/div[3]/button[1]',insist)
                        except:pass
                        try:
                            if driver.find_elements(By.XPATH,alertantigo):
                                EF.click(driver,'/html/body/div[5]/div/div[3]/button[1]',insist)
                        except:pass
                        try:
                            xpath = 'body > div.swal2-container.swal2-center.swal2-fade.swal2-shown'
                            alerta_registro = driver.find_element_by_css_selector(xpath)
                            sleep(2)
                        except:
                            pass
        
                        if(alerta_registro != ''):
                            #print('apareceu')
                            sleep(2)
                            xpath = '/html/body/div[5]/div/div[3]/button[1]'
                            btn_fechar = EF.click(driver,xpath,insist)
                            sleep(2)
                        
                        driver.switch_to.frame('DocumentoRepasse')
                        break
                except:
                    pass
        
            if(pag_atual != ultima_pag):
                xpath = '//*[@id="next_BAR_GridCadastro2"]'
                btn_prox =  EF.click(driver,xpath,insist)
                sleep(1)
                pag_atual += 1
            else:
                break   

        #Incluir Consultor
        driver.switch_to.default_content()
        driver.switch_to.frame('DocumentoRepasse')
        driver.execute_script("eventoAjax('Incluir');")
        driver.execute_script('acionaJanelaNivel1();')
        driver.switch_to.frame('SearchNivelHierarq')

        ############### Fim do script do LIRA ###################
        EF.fill(driver, '//*[@id="Descricao"]', 'CONSULTOR', insist)
        sleep(2)
        driver.execute_script("eventoAjax('EXECUTA');")
        try:
            Alert(driver).accept()
        except:pass
        driver.execute_script('RowDblClick("11");')
        sleep(2)
        driver.switch_to.default_content()
        driver.switch_to.frame('DocumentoRepasse')
        sleep(2)

        # Segunda parte
        sleep(2)
        btnPesquisaLupa = driver.find_element_by_xpath('//*[@id="Divisao1_Bt"]')
        driver.execute_script("arguments[0].click();", btnPesquisaLupa)
        sleep(2)
        driver.switch_to.frame('SearchDivisao')
        sleep(2)
        EF.fill(driver, '//*[@id="Nome"]', consultor, insist)
        sleep(2)
        xpath = '//*[@id="Nome"]'
        EF.click(driver, '//*[@id="SpanToolBarRightB"]/button')
        sleep(2)
        # Click no nome do consultor 
        sleep(2)
        EF.click(driver, '/html/body/form/div[5]/div[2]/div[2]/div/div[3]/div[3]/div/table/tbody/tr[2]/td[1]/a/span', insist)
        sleep(2)
        EF.click(driver, '//*[@id="BtEdiReg"]', insist)
        driver.switch_to.default_content()
        driver.switch_to.frame('DocumentoRepasse')
        EF.click(driver, '//*[@id="BtGravar"]', insist)
        sleep(3)
        driver.execute_script("evento('Voltar')")
        sleep(3)
        driver.switch_to.default_content()
        sleep(3)
        if driver.find_elements(By.XPATH,'/html/body/div[3]/div/div[3]/button[1]'):
            EF.click(driver,'/html/body/div[3]/div/div[3]/button[1]',insist)

    def importarquivos(driver):

        EF.acessarframe(driver, 'ZonaInterna',2)
        driver.switch_to.frame('Documento')
        EF.click(driver,'/html/body/form/div[7]/div/div[1]/ul/li[2]/div[2]/div[1]',insist)

        driver.switch_to.default_content()
        time.sleep(defaultsleep)
        driver.switch_to.frame('ScanImagem')

        EF.fill(driver, '//*[@id="ScanImagem_Descricao"]', 'Proposta de Endosso', insist)
        EF.click(driver, '//*[@id="select2-ScanImagem_TipoImagem-container"]', insist)
        EF.fill(driver, '/html/body/span[2]/span/span[1]/input', 'apolice', insist)
        EF.click(driver, '/html/body/span[2]/span/span[2]/ul/li[1]', insist)
        driver.execute_script("AbrirJanelaMultiplosArquivos();")
        driver.switch_to.frame('ScanImagem')
        filepath = driver.find_element(By.ID,'files')

        caminho_doc = GC.caminho_pdf + nome_arq
        
        filepath.send_keys(caminho_doc)   

        EF.click(driver,'/html/body/form/div[2]/input',insist)
        driver.switch_to.default_content()
        driver.switch_to.frame('ScanImagem')
        
        #Clica em grava
        EF.gravar(driver, '/html/body/div[5]/div/div[3]/button[1]', insist)
        EF.click(driver, '//*[@id="BtGravar"]', insist)
        EF.gravar(driver, '//*[@id="BtVoltar"]', insist)
        
    def preenchepremio(driver):
        try:
            #Verifica qual subtipo do endosso
            if 'endosso sem movimento de prêmio' in tipoEndosso.lower():
                return 1
            try:
                ExperienciaCliente = '/html/body/div[3]'
                if driver.find_element(By.XPATH,ExperienciaCliente):
                    EF.click(driver,'/html/body/div[3]/div[1]/a/span',insist)
            except:pass
            EF.acessarframe(driver, 'ZonaInterna',2)
            driver.switch_to.frame('Documento')
            
            EF.click(driver, '//*[@id="TitPremios"]', insist)

            if 'endosso cancelamento sem restituição' in tipoEndosso.lower():
                EF.fill(driver, '//*[@id="Documento_PremioLiquido"]', '0,00',insist,0)
                EF.gravar(driver, '//*[@id="BtGravar"]', insist)
                return 1
            
            EF.click(driver, '//*[@id="select2-Documento_MeioPagto-container"]', insist)
            #Faz a validação do tipo de pagamento do Endosso
            #pagamento em débito
            if 'débito' in formarPgto:    
                PagamentoDebito = 'DEBITO EM CONTA'
                EF.fill(driver, '/html/body/span[2]/span/span[1]/input', PagamentoDebito.lower(),insist)
                  #pagamento em boleto/Ficha
            elif 'boleto' in formarPgto and not 'débito' in formarPgto or 'ficha' in formarPgto:
                PagamentoBoleto = 'BOLETO / FICHA COMPENSACAO'
                EF.fill(driver,'/html/body/span[2]/span/span[1]/input',PagamentoBoleto.lower(), insist)
                #Pagamento em crédito
            elif 'cartão' and 'demais bandeiras' in formarPgto.lower():
                sleep(1)
                PagamentoCredito = 'CREDITO EM CONTA'
                EF.fill(driver,'/html/body/span[2]/span/span[1]/input',PagamentoCredito.lower(), insist)

                #pagamento em cartão porto
            elif 'cartão' and 'porto' in formarPgto.lower():
                PagamentoPorto = 'CARTAO CREDITO - PORTO'
                EF.fill(driver,'/html/body/span[2]/span/span[1]/input',PagamentoPorto.lower(), insist)

            sleep(1)

             #confirma o tipo do pagamento
            EF.click(driver, '/html/body/span[2]/span/span[2]/ul/li', insist)
            
            #numero de parcelas
            EF.fill(driver, '/html/body/form/div[7]/div/div[2]/div[10]/div/div[2]/div/div/div/div[1]/span/fieldset/div[2]/div[1]/div/input', qtdParcelas, insist)
            
            #Data de vencimento da primeira parcela
            #EF.fill(driver, '/html/body/form/div[7]/div/div[2]/div[10]/div/div[2]/div/div/div/div[1]/span/fieldset/div[3]/div[1]/div/div/input',dataVencimento, insist)
            
            #premio
            if seguradora.lower() == 'porto seguro cia' or seguradora.lower() == 'porto':
                sleep(2)
                 #Preenche Premio Total
                driver.execute_script(f'document.getElementById("Documento_PremioTotal2").value="{premioTotal}" ')
                sleep(1)
                EF.click(driver,'/html/body/form/div[7]/div/div[2]/div[10]/div/div[2]/div/div/div/div[1]/span/fieldset/div[4]/div/div/input',insist)
                sleep(1)
                EF.click(driver,'/html/body/form/div[7]/div/div[2]/div[10]/div/div[2]/div/div/div/div[1]/span/fieldset/div[3]/div[2]/div/input',insist)
                sleep(1)
            else:
                EF.fill(driver, '//*[@id="Documento_PremioLiqDesc"]', premio_liquido,insist, 1)
                EF.fill(driver, '//*[@id="Documento_Adicional"]', adicional,insist, 0)
                EF.fill(driver, '//*[@id="Documento_Custo"]', custo,insist, 0)
            
            EF.click(driver, '//*[@id="Documento_Juros"]',insist)
            EF.gravar(driver, '//*[@id="BtGravar"]', insist)
            driver.switch_to.default_content()
            alertmoeda = '/html/body/div[3]/div'
            alert = '/html/body/div[4]/div'
            try:
                try:
                    if driver.find_element(By.XPATH,alert):
                        EF.click(driver,'/html/body/div[4]/div/div[3]/button[1]',insist)
                except:pass
                try:
                    if driver.find_element(By.XPATH,alertmoeda):
                        EF.click(driver,'/html/body/div[3]/div/div[3]/button[1]',insist)
                except:pass
            except:pass
        except ValueError as err:
            raise ValueError(err)
        except Exception as err:
            raise Exception('Erro - Preencher prêmio')

    def primeiratela(driver,wait,rs):
        try:

            driver.switch_to.frame('Documento')
            
            fieldUsuario = '/html/body/form/div[7]/div/div[2]/div[1]/div/div[1]/div[2]/div[11]/div[3]/div/div[2]/div[1]/div[2]/div/span/span[1]/span/span[1]'
            usuario = EF.getText(driver,fieldUsuario,1,'innerText')

            #Selecionar tipo de endosso
            EF.click(driver, '//*[@id="DIVDocumento_TipoDocumento"]/div/span/span[1]/span', insist)
            EF.fill(driver, '/html/body/span[2]/span/span[1]/input',tipoEndosso, insist)
            EF.click(driver, '/html/body/span[2]/span/span[2]/ul/li', insist)
            
            #Selecionar subtipo de endosso
            EF.click(driver, '//*[@id="select2-Documento_SubTipo-container"]', insist)
            EF.fill(driver, '/html/body/span[2]/span/span[1]/input',subtipoEndosso, insist)
            EF.click(driver, '/html/body/span[2]/span/span[2]/ul/li', insist)

            AjusteCadastral(driver,wait)

            driver.switch_to.default_content()
            driver.switch_to.frame('ZonaInterna')
            driver.switch_to.frame('ZonaInterna')
            driver.switch_to.frame('Documento')
            
            #Inicio da vigencia
            sleep(2)
            EF.fill(driver, '//*[@id="Documento_InicioVigencia"]', inicioVigencia, insist)
            
            sleep(2)
            #colocar data protocolo cia
            EF.fill(driver, '/html/body/form/div[7]/div/div[2]/div[1]/div/div[1]/div[2]/div[7]/div[3]/div/div/input',inicioVigencia, insist)

            #data proposta
            inputDataProposta = '/html/body/form/div[7]/div/div[2]/div[1]/div/div[1]/div[2]/div[7]/div[2]/div/div/input'
            valueInputDataProposta = EF.getText(driver, inputDataProposta,1,'value')
            dataCadastro = date.today()
            dataCadastro = datetime.strftime(dataCadastro, '%d/%m/%Y')
            sleep(2)

            if valueInputDataProposta == '':
                EF.fill(driver,inputDataProposta,dataCadastro,clear=1)
            else:
                EF.fill(driver,inputDataProposta,dataCadastro,clear=1)
            
            #Numero proposta cia
            sleep(2)
            EF.fill(driver, '//*[@id="Documento_PropostaCia"]', nupropostacia, insist)
            
            #Abrir dados complementares
            EF.click(driver, '//*[@id="TitDadosComplementares"]', insist)
            
            #Escrever dados complementares
            EF.fill(driver, '/html/body/form/div[7]/div/div[2]/div[2]/div/div[2]/div[2]/div/div/textarea', 'Número do registro: {}'.format(rs), insist)
            
            #Clica em gravar
            sleep(2)
            driver.execute_script("evento('Gravar')")
            sleep(2)

            driver.switch_to.default_content()
            #faz a verificação dos Alerts 
            alert = '/html/body/div[3]/div'
            novoalert = '/html/body/div[4]/div'
            textoAlert = '/html/body/div[3]/div/div[2]/div[1]/div'
            textonovoAlert =  '/html/body/div[4]/div/div[2]/div[1]/div'
            sleep(3)
            ChaveAlert = False
            try:
                if driver.find_elements(By.XPATH,alert) or driver.find_elements(By.XPATH,novoalert):
                    try:
                        textAlert = driver.find_element(By.XPATH,textoAlert).get_attribute('innerText').lower()
                        ChaveAlert = True
                    except:pass
                    if ChaveAlert == False:
                        textonovoAlert = driver.find_element(By.XPATH,textonovoAlert).get_attribute('innerText').lower()
                    try:
                        if 'já possui uma proposta' in textAlert:    
                            EF.click(driver,'/html/body/div[3]/div/div[2]/div[1]/div/div/button[1]',insist)
                    
                    except:pass
                    try:
                        if 'já possui uma proposta' in textonovoAlert:
                            EF.click(driver,'/html/body/div[4]/div/div[2]/div[1]/div/div/button[1]',insist)
                    except:pass
                    try:
                        if 'nível corretora' in textAlert:
                            EF.click(driver,'/html/body/div[3]/div/div[2]/div[1]/div/div/button[1]',insist)
                        
                            sleep(3)
                    except:pass
                    try:
                        if 'nível corretora' in textonovoAlert:
                            EF.click(driver,'/html/body/div[4]/div/div[2]/div[1]/div/div/button[1]',insist)
                    except:pass
            except:pass
            #Fim da verificação
            driver.switch_to.frame('ZonaInterna')
            driver.switch_to.frame('ZonaInterna')
            driver.switch_to.frame('Documento')
            sleep(5)
            #Pega o numero da proposta do Quiver
            nupropostaquiver = EF.getText(driver, '/html/body/form/div[7]/div/div[2]/div[1]/div/div[1]/div[2]/div[7]/div[1]/div/input', insist, 'value')
            conn = pymssql.connect(server=GC.server, user=GC.user, password=GC.password, database=GC.database)
            cursor = conn.cursor()
            updatenuPropostaQuiver = f""" UPDATE cockpit_mds_endosso SET noPropostaQuiver = '{nupropostaquiver}' WHERE rs = {rs} """
            cursor.execute(updatenuPropostaQuiver)
            conn.commit()
            conn.close()
        except ValueError as err:
            raise ValueError(err)
        except Exception as err:
            raise Exception('Erro - Sistema Quiver')

    def findClient(driver,wait,rs):
        try:
            #Começa a procurar cliente
            driver.execute_script("SelecionaModuloJQuery('ConsultaEmissaoERecusa;Fast/FrmCadastroNovo.aspx?pagina=Documento','EMISSOESRECUSAS','Professional','EMISSOESRECUSAS','Propostas/Apólices'); ")

            driver.switch_to.frame('ZonaInterna')
            
            EF.click(driver, '/html/body/form/div[5]/div[1]/div/div/div[2]/div[1]/div[2]/div/div/span/span[1]/span/span[1]', insist)
            
            #Pesquisa por NºApólice
            EF.click(driver, '/html/body/span[2]/span/span[2]/ul/li[2]', insist)
            
            #Insere nºApólice
            fieldNumApolice = '/html/body/form/div[5]/div[1]/div/div/div[2]/div[1]/div[4]/div[1]/div[1]/div/input'
            EF.fill(driver, fieldNumApolice, nuapolice, insist)
        
            #Selecionar Situação - O PADRÃO É SELECINAR AS APÓLICES COM A SITUAÇÃO ATIVA
            fieldSituacaoAtiva = '/html/body/form/div[5]/div[1]/div/div/div[2]/div[1]/div[4]/div[2]/div/div/div/label[2]/input'
            EF.click(driver, fieldSituacaoAtiva, insist)

            # #Aperta o botão de pesquisar
            btnPesquisar = '/html/body/form/div[5]/div[1]/div/div/div[2]/div[4]/div[2]/span[1]/button'
            EF.click(driver, btnPesquisar, insist)

            #Volta frame para validar existência de alert
            driver.switch_to.default_content()

            
            alert = '/html/body/div[3]/div'
            try:
                wait.until(EC.visibility_of_element_located((By.XPATH,alert)))
                findAlert = True
            except:
                findAlert = False

            if findAlert == True:
                raise ValueError('Erro - Apólice não Localizada')
            else:
                print('Apólice Localizada')
         

            EF.acessarframe(driver,'ZonaInterna',1)

            time.sleep(2)
            EF.click(driver, '/html/body/form/div[5]/div[2]/div[2]/div/div[3]/div[3]/div/table/tbody/tr[2]/td[1]/a/span',insist)
            sleep(5)
            driver.switch_to.frame('ZonaInterna')

            driver.execute_script('AbreEndosso()')
            time.sleep(5)

        except ValueError as err:
            raise ValueError(err)
        except Exception as err:
            print(err)
         

    def terminaRobo(driver,rs,chave):
        #Validação que verifica se a cidade foi preenchida ou não
        dataCadastro = date.today()
        dataCadastro = datetime.strftime(dataCadastro, '%d/%m/%Y')

        if chave == True:
            conn = pymssql.connect(server=GC.server, user=GC.user, password=GC.password, database=GC.database)
            cursor = conn.cursor()
            updateStatus = f""" UPDATE cockpit_mds_endosso SET status = 'Concluido - S/ Cidade Veículo', statusRPA = 'Proposta cadastrada',dataCadastroQuiver = '{dataCadastro}' WHERE rs = {rs} """
            cursor.execute(updateStatus)
            conn.commit()
            conn.close()
            driver.quit()
        else:
            conn = pymssql.connect(server=GC.server, user=GC.user, password=GC.password, database=GC.database)
            cursor = conn.cursor()
            updateStatus = f""" UPDATE cockpit_mds_endosso SET status = 'Concluido', statusRPA = 'Proposta cadastrada' ,dataCadastroQuiver = '{dataCadastro}' WHERE rs = {rs} """
            cursor.execute(updateStatus)
            conn.commit()
            conn.close()
            driver.quit()
    def initWebDriver():
        try:
            driver = webdriver.Chrome(service=Service('./chromedriver.exe'))
            wait = WebDriverWait(driver, 10, poll_frequency=5)
            sleep(2)
            try:
                driver.get(GC.quiver)
            except:pass
            sleep(2)
            driver.maximize_window()
            return driver, wait
        except Exception as err:
            print(err)
            raise ValueError('Erro - Não entrou no Quiver')

    def main(rs):
        try:
            conn = pymssql.connect(server=GC.server, user=GC.user, password=GC.password, database=GC.database)
            cursor = conn.cursor()

            select = """ SELECT seguradora, tipoDeEndosso, subtipoDeEndosso, rs FROM cockpit_mds_endosso WHERE rs = {} """.format(rs)
            cursor.execute(select)
            row = cursor.fetchone()

            seguradora = row[0]
            tipoEndosso = row[1]
            subtipoEndosso = row[2]

            if tipoEndosso == 'Endosso de Cobrança':
                if subtipoEndosso == 'Substituição de itens':

                    driver, wait = initWebDriver()

                    findClient(driver,wait,rs)

                    primeiratela(driver,wait,rs)

                    chave = CadastroProduto(driver,wait,rs)

                    preenchepremio(driver)

                    cadastraConsultor(driver)

                    importarquivos(driver)

                    terminaRobo(driver,rs,chave)
                    
        except Exception as err:
            EF.saveErro(driver,rs,err)
    
    main(rs)



