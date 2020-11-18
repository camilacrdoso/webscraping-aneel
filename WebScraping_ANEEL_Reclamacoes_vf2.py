from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from bs4 import BeautifulSoup
import time
import csv
import pandas as pd
import sys, os

### 0. Entrada de parâmetros
url = 'https://www2.aneel.gov.br/aplicacoes_liferay/relatorios_de_qualidade_v2/'

# 1: RECLAMAÇÕES
# 2: QUALIDADE DO ATENDIMENTO TELEFÔNICO
# 3: QUALIDADE DO ATENDIMENTO COMERCIAL
# 4: IASC - ÍNDICE DE SATISFAÇÃO DO CONSUMIDOR
# 5: INADIMPLÊNCIA
# 6: TARIFA SOCIAL
# 7: UNIVERSALIZAÇÃO
tipo_relatorio = '1'

# 5_1_v2: Inadimplência por distribuidora por mês
# 5_2_v2: Inadimplência por distribuidora - evolução no ano
# 5_3_v2: Inadimplência e Suspensão de Fornecimento - Evolução por Classe por Distribuidora
# 5_4_v2: Inadimplência média e Suspensão de Fornecimento - Brasil
# 5_5_v2: Inadimplência média e Suspensão de Fornecimento por Classe - Brasil
relatorio = '1_7_v2'

# C: Concessionária
# P: Permissionária
# T: Todas
tipo_distribuidora = 'T'

anos_selecionados = ['2020']
meses_selecionados = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12']

nome_arquivo_saida = '\DadosANEEL_Reclamacoes_' + anos_selecionados[0] + '_' + meses_selecionados[0] + '-' + meses_selecionados[len(meses_selecionados) - 1] + '.csv'
caminho_saida_csv = r'C:\Users\Camila\Google Drive\Idec\3. Atividades\3. Dados\5. Dados ANEEL\2. Reclamacoes\Dados_WebScrapping' + nome_arquivo_saida
tabela_csv = []
df_saida = pd.DataFrame(columns = ['distribuidora', 'ano', 'mes', 'descricao', 'reclamacoes_recebidas', 'reclamacoes_procedentes', 'prazo_medio_encerramento_horas', 'prazo_medio_encerramento_dias', 'reclamacoes_improcedentes', 'reclamacoes_totais' ])
dict_meses = {"1": "Janeiro",
              "2": "Fevereiro",
              "3": "Março",
              "4": "Abril",
              "5": "Maio",
              "6": "Junho",
              "7": "Julho",
              "8": "Agosto",
              "9": "Setembro",
              "10": "Outubro",
              "11": "Novembro",
              "12": "Dezembro"}

### 1. Fazendo a requisição HTML para acessar a página da ANEEL
def requisicaoHTML():
  chrome_options = Options()
  chrome_options.add_argument("--window-size=900,600")
  driver = webdriver.Chrome(r'C:\Users\Camila\Downloads\chromedriver_win32\chromedriver.exe', chrome_options = chrome_options)
  driver.get(url)
  return driver

### 2. Acessando o conteúdo HTML da página e buscando as informações necessárias
# 2.1 Seleciona o tipo de relatório
def selecionaTipoRelatorio(driver):
  try:
    select_relatorio_pai = Select(driver.find_element_by_xpath("//select[@id='select_tipo_relatorio_pai']"))
    select_relatorio_pai.select_by_value(tipo_relatorio)
    selecionou = verificaSelecao(select_relatorio_pai,tipo_relatorio)
    while selecionou != True:
      selecionou = verificaSelecao(select_relatorio_pai,tipo_relatorio)
    return selecionou
  except Exception as e:
    erroStaleElement(driver,'select_tipo_relatorio_pai')
    selecionaTipoRelatorio(driver)

# 2.2 Seleciona o relatório
def selecionaRelatorio(driver):
  try:
    select_relatorio = Select(driver.find_element_by_xpath("//select[@id='select_tipo_relatorio']"))
    select_relatorio.select_by_value(relatorio)
    selecionou = verificaSelecao(select_relatorio,relatorio)
    while selecionou != True:
      selecionou = verificaSelecao(select_relatorio,relatorio)
    return selecionou
  except Exception as e:
    erroStaleElement(driver,'select_tipo_relatorio')
    selecionaRelatorio(driver)

# 2.3 Seleciona o tipo de distribuidora
def selecionaTipoDistribuidora(driver):
  try:
    select_tipo_distribuidora = Select(driver.find_element_by_xpath("//select[@id='select_classificacao_agente']"))
    select_tipo_distribuidora.select_by_value(tipo_distribuidora)
    selecionou = verificaSelecao(select_tipo_distribuidora,tipo_distribuidora)
    while selecionou != True:
      selecionou = verificaSelecao(select_tipo_distribuidora,tipo_distribuidora)
    return selecionou
  except Exception as e:
    erroStaleElement(driver,'select_classificacao_agente')
    selecionaTipoDistribuidora(driver)

# 2.4 Gera lista de distribuidoras sobre a qual iremos iterar
def geraListaDistribuidoras(driver):
  try:
    select_distribuidora = Select(driver.find_element_by_xpath("//select[@id='select_agente']"))
    select_distribuidora_opcoes = select_distribuidora.options
    lista_nomes = [opcao.get_attribute("value") for opcao in select_distribuidora_opcoes]
    return lista_nomes
  except Exception as e:
    erroStaleElement(driver,'select_agente')
    return None

# 2.5 Seleciona distribuidora
def selecionaDistribuidora(distribuidora):
  try:
    select_distribuidora = Select(driver.find_element_by_xpath("//select[@id='select_agente']"))    
    select_distribuidora.select_by_value(distribuidora)
    selecionou = verificaSelecao(select_distribuidora,distribuidora)
    while selecionou != True:
      selecionou = verificaSelecao(select_distribuidora,distribuidora)
    return selecionou
  except Exception as e:
    print(e)
    if e == StaleElementReferenceException:
      erroStaleElement(driver,'select_agente')
      selecionaDistribuidora(driver)
  
# 2.6 Gera lista de anos
def geraListaAnos(driver):
  try:
    select_ano = Select(driver.find_element_by_xpath("//select[@id='select_ano']"))
    select_ano_opcoes = select_ano.options
    lista_anos = [opcao.get_attribute("value") for opcao in select_ano_opcoes]
    return lista_anos
  except Exception as e:
    erroStaleElement(driver,'select_ano')
    return None

# 2.7 Seleciona ano
def selecionaAno(ano):
  try:
    select_ano = Select(driver.find_element_by_xpath("//select[@id='select_ano']"))   
    select_ano.select_by_value(ano)
    selecionou = verificaSelecao(select_ano,ano)
    while selecionou != True:
      selecionou = verificaSelecao(select_ano,ano)
    return selecionou
  except Exception as e:
    erroStaleElement(driver,'select_ano')
    selecionaDistribuidora(driver)

# 2.8 Gera lista de meses
def geraListaMeses(driver):
  try:
    select_mes = Select(driver.find_element_by_xpath("//select[@id='select_mes']"))
    select_mes_opcoes = select_mes.options
    lista_meses = [opcao.get_attribute("value") for opcao in select_mes_opcoes]
    return lista_meses
  except Exception as e:
    erroStaleElement(driver,'select_mes')
    return None

# 2.9 Seleciona mes
def selecionaMes(mes):
  try:
    select_mes = Select(driver.find_element_by_xpath("//select[@id='select_mes']"))   
    select_mes.select_by_value(mes)
    selecionou = verificaSelecao(select_mes,mes)
    while selecionou != True:
      selecionou = verificaSelecao(select_mes,mes)
    return selecionou
  except Exception as e:
    erroStaleElement(driver,'select_mes')
    return None

# 2.10 Armazenando os dados
def armazenaDados(driver, distribuidora, ano, mes):
  try:
    #print(mes, "entrou no armazenaDados")
    indice = 0 if df_saida.empty else df_saida.index.max() + 1
    
    tabela_pagina = driver.find_element_by_xpath("//div[@id='div_relatorio']")
    linhas_tabela_pagina = tabela_pagina.find_element_by_tag_name("table").find_elements_by_tag_name("tr")
    #print(mes, len(linhas_tabela_pagina))
    count_linhas_armazenadas = 0
    for linha in linhas_tabela_pagina:
      if linha.get_attribute("class") == 'tr-cabecalho-tabela':
        for cabecalho in linha.find_elements_by_tag_name("td"):
          if cabecalho.get_attribute("colspan") == "7":
            if cabecalho.text != distribuidora + ' - ' + ano + ' - ' + dict_meses[mes]:
              #print("não atualizou a tabela ainda")
              return False

      if linha.get_attribute("class") in ["odd", "even"]:
        colunas_linha = linha.find_elements_by_tag_name("td")
        dados_csv = [distribuidora, ano, mes]
        for coluna in colunas_linha:   
          dado = coluna.text.replace(".","").replace(",", ".")
          if dado =="TOTAL":
            break
          else:
            try:
              dados_csv.append(int(dado))
            except:
              try:
                dados_csv.append(float(dado))
              except:
                dados_csv.append(dado)

        # Adicionando linha de dados à tabela
        if len(dados_csv) > 3:
          df_saida.loc[indice] = dados_csv
          count_linhas_armazenadas += 1
          #print("append", mes, count_linhas_armazenadas)
          indice += 1
          if count_linhas_armazenadas == 15:
            #print("entrou no return")
            return True
            

  except Exception as e:
    exc_type, exc_obj, exc_tb = sys.exc_info()
    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
    #print(exc_type, fname, exc_tb.tb_lineno)
    erroStaleElement(driver,'div_relatorio')
    armazenaDados(driver, distribuidora, ano, mes) 

### 3. Escrevendo CSV
def escrevendoCSV():
  df_saida.to_csv(caminho_saida_csv, mode = 'a', sep = ';', quoting = csv.QUOTE_NONE, index = False, header = None)
  


def erroStaleElement(driver, id_elemento):
  #print("deu erro StaleElement",id_elemento)
  my_element_id = id_elemento
  ignored_exceptions=(NoSuchElementException,StaleElementReferenceException)
  your_element = WebDriverWait(driver, 10,ignored_exceptions=ignored_exceptions).until(expected_conditions.presence_of_element_located((By.ID, my_element_id)))

def verificaSelecao(select, valor):
  opcao_selecionada = select.first_selected_option.get_attribute("value")
  if opcao_selecionada == valor:
    return True
  else:
    return False

def primeirosSelects(driver):  
  if selecionaTipoRelatorio(driver):
    selecionaRelatorio(driver)
  if selecionaRelatorio(driver):
    selecionaTipoDistribuidora(driver)
  if selecionaTipoDistribuidora:
    return True


driver = requisicaoHTML()
time.sleep(1)
primeiros_selects = primeirosSelects(driver)
if primeiros_selects:
  lista_distribuidoras = geraListaDistribuidoras(driver)
  while lista_distribuidoras == None:
    lista_distribuidoras = geraListaDistribuidoras(driver)

  lista_distribuidoras_unicos = sorted([elemento for elemento in set(lista_distribuidoras)])
  
  for distribuidora in lista_distribuidoras_unicos[22:len(lista_distribuidoras_unicos)]:
    selecionou_distribuidora = selecionaDistribuidora(distribuidora)
    if selecionou_distribuidora:
      lista_anos = geraListaAnos(driver)
      while lista_anos == None:
        lista_anos = geraListaAnos(driver)
      
      for ano in lista_anos[1:len(lista_anos)]:
        selecionou_ano = selecionaAno(ano)
        if selecionou_ano and ano in anos_selecionados:
          lista_meses = geraListaMeses(driver)
          while lista_meses == None:
            lista_meses = geraListaMeses(driver)
          # print(lista_meses)
          
          for mes in lista_meses[1:len(lista_meses)]:
            selecionou_mes = selecionaMes(mes)
            if selecionou_mes and mes in meses_selecionados:
              armazenou = armazenaDados(driver, distribuidora, ano, mes)
              while armazenou == False:
                armazenou = armazenaDados(driver, distribuidora, ano, mes)

              # print("saiu do armazonaDados")
              if len(df_saida[(df_saida['distribuidora'] == distribuidora) & (df_saida['ano'] == ano) & (df_saida['mes'] == mes)].index) != 15:
                print(distribuidora, ano, mes, len(df_saida[(df_saida['distribuidora'] == distribuidora) & (df_saida['ano'] == ano) & (df_saida['mes'] == mes)].index))
              
          
    if df_saida.empty != True:
      escrevendoCSV()
      print(lista_distribuidoras_unicos.index(distribuidora), distribuidora, "ok")
      df_saida = pd.DataFrame(columns = ['distribuidora', 'ano', 'mes', 'descricao', 'reclamacoes_recebidas', 'reclamacoes_procedentes', 'prazo_medio_encerramento_horas', 'prazo_medio_encerramento_dias', 'reclamacoes_improcedentes', 'reclamacoes_totais' ])



      


