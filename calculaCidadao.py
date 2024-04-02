from tkinter import *
from tkinter import filedialog
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from UliPlot.XLSX import auto_adjust_xlsx_column_width

'''
FUNCIONALIDADES:

!!!para um bom funcionamento a planilha deve ser do mesmo modelo da planilha de exemplo(ex01.xlsx)!!!

--Realiza consultas no site calculadora do cidadao com base nos dados da planilha selecionada e cria uma nova planilha com os dados obtidos

--As consultas sao: POUPANCA, SELIC, CDI 100%, IGPM e IPCA

--Tambem funciona como executavel

--Parametros para busca Tkinter: Botao para buscar o arquivo .xlsx no browser, opcao de selecao para data final a ser filtrada(data inicial ja inclusa na planilha), 
Botao executar para ser executado o progama com a data final e a planilha selecionada

'''

driver = webdriver.Chrome()

# Função para abrir o explorador de arquivos e selecionar o arquivo .xlsx
def browse_file():
    filename = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if filename:
        arquivo_entry.delete(0, END)
        arquivo_entry.insert(0, filename)

# Função para executar a consulta e fechar a janela
def execute_and_close():
    global hoje, valorCorrigidoPoup, lista_data, lista_lancamento, finalPoup, finalSelic, finalCdi, finalIgpm, finalIpca

    # Colsulta os dados input do Tkinter
    filename = arquivo_entry.get()
    selected_date = cal.get_date()
    # Formata a data input para o formato dia mes ano
    data_final = selected_date.strftime("%d%m%Y")

    # Consulta a data atual e a guarda em uma variavel
    mesAtras = datetime.today() - timedelta(days=30)
    # Formata a data atual para o formato dia mes ano
    mesatrasFormat = mesAtras.strftime("%d%m%Y")

    # Le o arquivo Excel e cria um dataframe
    df = pd.read_excel(filename)
    df1 = df.copy()

    # Converte os dados para serem usados na pesquisa
    df['Data'] = pd.to_datetime(df['Data'], format='%Y-%m-%d').dt.strftime('%d%m%Y')
    df1['Data'] = pd.to_datetime(df1['Data'], format='%Y-%m-%d').dt.strftime('%d%m%Y')

    # Formata para ser usado 2 numeros apos a virgula nos dados da coluna: "Lançamento R$"
    df["Lançamento R$"] = df["Lançamento R$"].apply(lambda x: format(float(x), ".2f"))
    df['Lançamento R$'] = df['Lançamento R$'].astype(str)

    # Fecha a janela do Tkinter
    root.destroy()

    # Acessa o site para a consulta
    url = f"https://www3.bcb.gov.br/CALCIDADAO/publico/exibirFormCorrecaoValores.do?method=exibirFormCorrecaoValores"
    driver.get(url)
    time.sleep(2)

    # Formata as datas
    # datamy: Formata o imput da data final para ser usada na pesquisa no formato mes ano
    dataMy = datetime.strptime(data_final, '%d%m%Y').date().strftime("%m%Y")
    # datamy1: Formata a data de 30 dias atras para ser usado nas Exceptions, pesquisa no formato mes ano
    dataMy1 = datetime.strptime(mesatrasFormat, '%d%m%Y').date().strftime("%m%Y")

    # dataFinal: Formata o imput da data final para ser usada na pesquisa no formato ano-mes-dia
    dataFinal = datetime.strptime(data_final, '%d%m%Y').date().strftime("%Y-%m-%d")

    # dataFinal: Formata o imput da data final para ser usada na pesquisa no formato ano-mes
    dataFinalMy = datetime.strptime(dataFinal, '%Y-%m-%d').date().strftime("%Y-%m")

    # datamy1: Formata a data de 30 dias atras para ser usado nas Exceptions, pesquisa no formato ano-mes-dia
    mesAtrasSeparado = datetime.strptime(mesatrasFormat, '%d%m%Y').date().strftime("%Y-%m-%d")

    # datamy1: Formata a data de 30 dias atras para ser usado nas Exceptions, pesquisa no formato ano-mes
    mesAtrasSeparadoMy = datetime.strptime(mesatrasFormat, '%d%m%Y').date().strftime("%Y-%m")


    # Dados poup Nova
    # Clica na aba poupanca
    driver.find_element(By.XPATH, '//*[@id="oTab"]/td[5]/a').click()

    valorCorrigidoPoup = []
    # iterating over rows using iterrows() function
    for i, j in df.iterrows():

        # Formata a data inicial para ser usada na condicional de pesquisa
        dataInicial = datetime.strptime(j[0], '%d%m%Y').date().strftime("%Y-%m-%d")


        # Clica na opcao nova da aba poupanca
        driver.find_element(By.XPATH,
                            '/html/body/div[6]/table/tbody/tr[3]/td/div/form/div[1]/table/tbody/tr[6]/td[2]/input[1]').click()
        time.sleep(5)

        if dataFinal > dataInicial:

            # Seleciona a data inicial
            driver.find_element(By.XPATH,
                                '/html/body/div[6]/table/tbody/tr[3]/td/div/form/div[1]/table/tbody/tr[3]/td[2]/input').send_keys(j)

            # Seleciona a data final
            driver.find_element(By.XPATH,
                                '/html/body/div[6]/table/tbody/tr[3]/td/div/form/div[1]/table/tbody/tr[4]/td[2]/input').send_keys(data_final)

            # Seleciona o valor a ser corrigido
            driver.find_element(By.XPATH,
                                '/html/body/div[6]/table/tbody/tr[3]/td/div/form/div[1]/table/tbody/tr[5]/td[2]/input').send_keys(j[2])

            # Clica no botao corrigir valor
            driver.find_element(By.XPATH, '/html/body/div[6]/table/tbody/tr[3]/td/div/form/div[2]/input[1]').click()
            time.sleep(2)

            # Armazena o valor cirrigido em um array
            valor = driver.find_element(By.XPATH,
                                        '/html/body/div[6]/table/tbody/tr/td/form/div[2]/table[1]/tbody/tr[9]/td[2]').text
            # Aplicando a expressão regular para encontrar apenas os números e a vírgula
            numeros_com_virgula = re.findall(r'\d|,', valor)
            # Substituindo a vírgula por ponto
            numeros_str = ''.join(numeros_com_virgula).replace(',', '.')
            numero_float = float(numeros_str)
            valorCorrigidoPoup.append(numero_float)

            # Armazena a data final usada em uma varialvel
            finalPoup = driver.find_element(By.XPATH, '/html/body/div[6]/table/tbody/tr/td/form/div[2]/table[1]/tbody/tr[3]/td[2]').text

            # Clica no botao fazer nova pesquisa
            driver.find_element(By.XPATH, '/html/body/div[6]/table/tbody/tr/td/form/div[2]/table[2]/tbody/tr/td[1]/input').click()
        else:
            valorCorrigidoPoup.append('-')


    # Dados Selic
    # Clica na aba Selic
    driver.find_element(By.XPATH, '//*[@id="oTab"]/td[7]/a').click()
    time.sleep(5)

    valorCorrigidoSelic = []
    # iterating over rows using iterrows() function
    for i, j in df.iterrows():

        # Formata a data inicial para ser usada na condicional de pesquisa
        dataInicial = datetime.strptime(j[0], '%d%m%Y').date().strftime("%Y-%m-%d")

        if dataFinal > dataInicial:

            # Seleciona a data inicial
            driver.find_element(By.XPATH,
                                '/html/body/div[6]/table/tbody/tr[3]/td/div/form/div[1]/table/tbody/tr[4]/td[2]/input').send_keys(j)

            # Seleciona a data final
            driver.find_element(By.XPATH,
                                '/html/body/div[6]/table/tbody/tr[3]/td/div/form/div[1]/table/tbody/tr[5]/td[2]/input').send_keys(data_final)

            # Seleciona o valor a ser corrigido
            driver.find_element(By.XPATH,
                                '/html/body/div[6]/table/tbody/tr[3]/td/div/form/div[1]/table/tbody/tr[6]/td[2]/input').send_keys(j[2])
            time.sleep(2)

            # Clica no botao corrigir valor
            driver.find_element(By.XPATH, '/html/body/div[6]/table/tbody/tr[3]/td/div/form/div[2]/input[1]').click()

            # Armazena o valor cirrigido em um array
            valor = driver.find_element(By.XPATH,
                                        '/html/body/div[6]/table/tbody/tr/td/form/div[2]/table[1]/tbody/tr[8]/td[2]').text
            # Aplicando a expressão regular para encontrar apenas os números e a vírgula
            numeros_com_virgula = re.findall(r'\d|,', valor)
            # Substituindo a vírgula por ponto
            numeros_str = ''.join(numeros_com_virgula).replace(',', '.')
            numero_float = float(numeros_str)
            valorCorrigidoSelic.append(numero_float)

            # Armazena a data final usada em uma varialvel
            finalSelic = driver.find_element(By.XPATH, '/html/body/div[6]/table/tbody/tr/td/form/div[2]/table[1]/tbody/tr[3]/td[2]').text

            # Clica no botao fazer nova pesquisa
            driver.find_element(By.XPATH, '/html/body/div[6]/table/tbody/tr/td/form/div[2]/table[2]/tbody/tr/td[1]/input').click()
        else:
            valorCorrigidoSelic.append('-')


    # Dados Cdi
    # Clica na aba Cdi
    driver.find_element(By.XPATH, '//*[@id="oTab"]/td[9]/a').click()
    time.sleep(5)

    valorCorrigidoCdi = []
    # iterating over rows using iterrows() function
    for i, j in df.iterrows():

        # Formata a data inicial para ser usada na condicional de pesquisa
        dataInicial = datetime.strptime(j[0], '%d%m%Y').date().strftime("%Y-%m-%d")

        if dataFinal > dataInicial:

            # Seleciona a data inicial
            driver.find_element(By.XPATH,
                                '/html/body/div[6]/table/tbody/tr[3]/td/div/form/div[1]/table/tbody/tr[4]/td[2]/input').send_keys(j)

            # Seleciona a data final
            driver.find_element(By.XPATH,
                                '/html/body/div[6]/table/tbody/tr[3]/td/div/form/div[1]/table/tbody/tr[5]/td[2]/input').send_keys(data_final)

            # Seleciona o valor a ser corrigido
            driver.find_element(By.XPATH,
                                '/html/body/div[6]/table/tbody/tr[3]/td/div/form/div[1]/table/tbody/tr[6]/td[2]/input').send_keys(j[2])

            # Seleciona o % do CDi (padrao 100%)
            driver.find_element(By.XPATH,
                                '/html/body/div[6]/table/tbody/tr[3]/td/div/form/div[1]/table/tbody/tr[7]/td[2]/input').send_keys(10000)

            # Clica no botao corrigir valor
            driver.find_element(By.XPATH, '/html/body/div[6]/table/tbody/tr[3]/td/div/form/div[2]/input[1]').click()
            time.sleep(2)

            try:
                # Armazena o valor cirrigido em um array
                valor = driver.find_element(By.XPATH,
                                            '/html/body/div[6]/table/tbody/tr/td/form/div[2]/table[1]/tbody/tr[9]/td[2]').text
                # Aplicando a expressão regular para encontrar apenas os números e a vírgula
                numeros_com_virgula = re.findall(r'\d|,', valor)
                # Substituindo a vírgula por ponto
                numeros_str = ''.join(numeros_com_virgula).replace(',', '.')
                numero_float = float(numeros_str)
                valorCorrigidoCdi.append(numero_float)

                # Armazena a data final usada em uma varialvel
                finalCdi = driver.find_element(By.XPATH,
                                               '/html/body/div[6]/table/tbody/tr/td/form/div[2]/table[1]/tbody/tr[3]/td[2]').text
                # Clica no botao fazer nova pesquisa
                driver.find_element(By.XPATH,
                                    '/html/body/div[6]/table/tbody/tr/td/form/div[2]/table[2]/tbody/tr/td[1]/input').click()

            except Exception:
                if mesAtrasSeparado > dataInicial:
                    # Deleta o campo da data final
                    driver.find_element(By.XPATH,
                                        '/html/body/div[6]/table/tbody/tr[3]/td/div/form/div[1]/table/tbody/tr[5]/td[2]/input').send_keys(
                        Keys.CONTROL, 'a', Keys.BACKSPACE)

                    # Seleciona a data final
                    driver.find_element(By.XPATH,
                                        '/html/body/div[6]/table/tbody/tr[3]/td/div/form/div[1]/table/tbody/tr[5]/td[2]/input').send_keys(
                        mesatrasFormat)

                    # Clica no botao corrigir valor
                    driver.find_element(By.XPATH, '/html/body/div[6]/table/tbody/tr[3]/td/div/form/div[2]/input[1]').click()
                    time.sleep(2)

                    # Armazena o valor cirrigido em um array
                    valor = driver.find_element(By.XPATH,
                                                '/html/body/div[6]/table/tbody/tr/td/form/div[2]/table[1]/tbody/tr[9]/td[2]').text
                    # Aplicando a expressão regular para encontrar apenas os números e a vírgula
                    numeros_com_virgula = re.findall(r'\d|,', valor)
                    # Substituindo a vírgula por ponto
                    numeros_str = ''.join(numeros_com_virgula).replace(',', '.')
                    numero_float = float(numeros_str)
                    valorCorrigidoCdi.append(numero_float)


                    # Armazena a data final usada em uma varialvel
                    finalCdi = driver.find_element(By.XPATH,
                                                   '/html/body/div[6]/table/tbody/tr/td/form/div[2]/table[1]/tbody/tr[3]/td[2]').text
                    # Clica no botao fazer nova pesquisa
                    driver.find_element(By.XPATH,
                                        '/html/body/div[6]/table/tbody/tr/td/form/div[2]/table[2]/tbody/tr/td[1]/input').click()
                else:
                    # Limpa todos os campos de pesquisa do cdi
                    driver.find_element(By.XPATH,
                                        '//*[@id="oTab"]/td[9]/a').click()

                    valorCorrigidoCdi.append('-')
        else:
            valorCorrigidoCdi.append('-')


    # Passa a data para mes/ano seguindo os padroes de pesquisa
    df['Data'] = pd.to_datetime(df['Data'], format='%d%m%Y').dt.strftime('%m%Y')

    # Dados IgpM
    # Clica na aba IgpM
    driver.find_element(By.XPATH, '//*[@id="oTab"]/td[1]/a').click()
    time.sleep(5)

    valorCorrigidoIgpM = []
    # iterating over rows using iterrows() function
    for i, j in df.iterrows():

        # solucao: https://sqa.stackexchange.com/questions/1355/what-is-the-correct-way-to-select-an-option-using-seleniums-python-webdriver
        # Seleciona o tipo de pesquisa
        el = driver.find_element(By.ID, 'selIndice')
        for option in el.find_elements(By.TAG_NAME, 'option'):
            if option.text == 'IGP-M (FGV) - a partir de 06/1989':
                option.click()  # select() in earlier versions of webdriver
                break

        # Formata a data inicial para ser usada na condicional de pesquisa
        dataInicial = datetime.strptime(j[0], '%m%Y').date().strftime("%Y-%m")


        if dataFinalMy > dataInicial:

            # Seleciona a data inicial
            driver.find_element(By.XPATH,
                                '//*[@id="corrigirPorIndiceForm"]/div[1]/table/tbody/tr[4]/td[2]/input').send_keys(j)

            # Seleciona a data final
            driver.find_element(By.XPATH,
                                '//*[@id="corrigirPorIndiceForm"]/div[1]/table/tbody/tr[5]/td[2]/input').send_keys(f'{dataMy}')

            # Seleciona o valor a ser corrigido
            driver.find_element(By.XPATH,
                                '//*[@id="corrigirPorIndiceForm"]/div[1]/table/tbody/tr[6]/td[2]/input').send_keys(j[2])

            # Clica no botao corrigir valor
            driver.find_element(By.XPATH, '//*[@id="corrigirPorIndiceForm"]/div[2]/input[1]').click()
            time.sleep(2)

            try:
                # Armazena o valor cirrigido em um array
                valor = driver.find_element(By.XPATH,
                                            '/html/body/div[6]/table/tbody/tr/td/div[2]/table[1]/tbody/tr[8]/td[2]').text
                # Aplicando a expressão regular para encontrar apenas os números e a vírgula
                numeros_com_virgula = re.findall(r'\d|,', valor)
                # Substituindo a vírgula por ponto
                numeros_str = ''.join(numeros_com_virgula).replace(',', '.')
                numero_float = float(numeros_str)
                valorCorrigidoIgpM.append(numero_float)

                # Armazena a data final usada em uma varialvel
                finalIgpm = driver.find_element(By.XPATH,
                                                '/html/body/div[6]/table/tbody/tr/td/div[2]/table[1]/tbody/tr[3]/td[2]').text

                # Clica no botao fazer nova pesquisa
                driver.find_element(By.XPATH,
                                    '/html/body/div[6]/table/tbody/tr/td/div[2]/table[2]/tbody/tr/td[1]/input').click()

            except Exception:
                if mesAtrasSeparadoMy > dataInicial:

                    driver.find_element(By.XPATH,
                                        '//*[@id="corrigirPorIndiceForm"]/div[1]/table/tbody/tr[5]/td[2]/input').send_keys(Keys.CONTROL, 'a', Keys.BACKSPACE)

                    # Seleciona a data final
                    driver.find_element(By.XPATH,
                                        '//*[@id="corrigirPorIndiceForm"]/div[1]/table/tbody/tr[5]/td[2]/input').send_keys(f'{dataMy1}')

                    # Clica no botao corrigir valor
                    driver.find_element(By.XPATH, '//*[@id="corrigirPorIndiceForm"]/div[2]/input[1]').click()
                    time.sleep(2)

                    # Armazena o valor cirrigido em um array
                    valor = driver.find_element(By.XPATH,
                                                '/html/body/div[6]/table/tbody/tr/td/div[2]/table[1]/tbody/tr[8]/td[2]').text
                    # Aplicando a expressão regular para encontrar apenas os números e a vírgula
                    numeros_com_virgula = re.findall(r'\d|,', valor)
                    # Substituindo a vírgula por ponto
                    numeros_str = ''.join(numeros_com_virgula).replace(',', '.')
                    numero_float = float(numeros_str)
                    valorCorrigidoIgpM.append(numero_float)

                    # Armazena a data final usada em uma varialvel
                    finalIgpm = driver.find_element(By.XPATH,
                                                    '/html/body/div[6]/table/tbody/tr/td/div[2]/table[1]/tbody/tr[3]/td[2]').text

                    # Clica no botao fazer nova pesquisa
                    driver.find_element(By.XPATH,
                                        '/html/body/div[6]/table/tbody/tr/td/div[2]/table[2]/tbody/tr/td[1]/input').click()

                else:
                    # Limpa todos os campos de pesquisa do igpm
                    driver.find_element(By.XPATH, '//*[@id="oTab"]/td[1]/a').click()

                    valorCorrigidoIgpM.append('-')
        else:
            valorCorrigidoIgpM.append('-')


    # Dados Ipca
    # Clica na aba Ipca
    driver.find_element(By.XPATH, '//*[@id="oTab"]/td[1]/a').click()
    time.sleep(5)

    valorCorrigidoIpcA = []
    # iterating over rows using iterrows() function
    for i, j in df.iterrows():

        # solucao: https://sqa.stackexchange.com/questions/1355/what-is-the-correct-way-to-select-an-option-using-seleniums-python-webdriver
        # Seleciona o tipo de pesquisa
        el = driver.find_element(By.ID, 'selIndice')
        for option in el.find_elements(By.TAG_NAME, 'option'):
            if option.text == 'IPCA (IBGE) - a partir de 01/1980':
                option.click()  # select() in earlier versions of webdriver
                break

        # Formata a data inicial para ser usada na condicional de pesquisa
        dataInicial = datetime.strptime(j[0], '%m%Y').date().strftime("%Y-%m")

        if dataFinalMy > dataInicial:

            # Seleciona a data inicial
            driver.find_element(By.XPATH,
                                '//*[@id="corrigirPorIndiceForm"]/div[1]/table/tbody/tr[4]/td[2]/input').send_keys(j)

            # Seleciona a data final
            driver.find_element(By.XPATH,
                                '//*[@id="corrigirPorIndiceForm"]/div[1]/table/tbody/tr[5]/td[2]/input').send_keys(f'{dataMy}')

            # Seleciona o valor a ser corrigido
            driver.find_element(By.XPATH,
                                '//*[@id="corrigirPorIndiceForm"]/div[1]/table/tbody/tr[6]/td[2]/input').send_keys(j[2])

            # Clica no botao corrigir valor
            driver.find_element(By.XPATH, '//*[@id="corrigirPorIndiceForm"]/div[2]/input[1]').click()
            time.sleep(2)

            try:
                # Armazena o valor cirrigido em um array
                valor = driver.find_element(By.XPATH,
                                            '/html/body/div[6]/table/tbody/tr/td/div[2]/table[1]/tbody/tr[8]/td[2]').text
                # Aplicando a expressão regular para encontrar apenas os números e a vírgula
                numeros_com_virgula = re.findall(r'\d|,', valor)
                # Substituindo a vírgula por ponto
                numeros_str = ''.join(numeros_com_virgula).replace(',', '.')
                numero_float = float(numeros_str)
                valorCorrigidoIpcA.append(numero_float)

                # Armazena a data final usada em uma varialvel
                finalIpca = driver.find_element(By.XPATH,
                                                '/html/body/div[6]/table/tbody/tr/td/div[2]/table[1]/tbody/tr[3]/td[2]').text

                # Clica no botao fazer nova pesquisa
                driver.find_element(By.XPATH,
                                    '/html/body/div[6]/table/tbody/tr/td/div[2]/table[2]/tbody/tr/td[1]/input').click()

            except Exception:
                if mesAtrasSeparadoMy > dataInicial:

                    # Seleciona a data final
                    driver.find_element(By.XPATH,
                                        '//*[@id="corrigirPorIndiceForm"]/div[1]/table/tbody/tr[5]/td[2]/input').send_keys(Keys.CONTROL, 'a', Keys.BACKSPACE)

                    # Seleciona a data final
                    driver.find_element(By.XPATH,
                                        '//*[@id="corrigirPorIndiceForm"]/div[1]/table/tbody/tr[5]/td[2]/input').send_keys(f'{dataMy1}')

                    # Clica no botao corrigir valor
                    driver.find_element(By.XPATH, '//*[@id="corrigirPorIndiceForm"]/div[2]/input[1]').click()
                    time.sleep(2)

                    # Armazena o valor cirrigido em um array
                    valor = driver.find_element(By.XPATH,
                                                '/html/body/div[6]/table/tbody/tr/td/div[2]/table[1]/tbody/tr[8]/td[2]').text
                    # Aplicando a expressão regular para encontrar apenas os números e a vírgula
                    numeros_com_virgula = re.findall(r'\d|,', valor)
                    # Substituindo a vírgula por ponto
                    numeros_str = ''.join(numeros_com_virgula).replace(',', '.')
                    numero_float = float(numeros_str)
                    valorCorrigidoIpcA.append(numero_float)

                    # Armazena a data final usada em uma varialvel
                    finalIpca = driver.find_element(By.XPATH,
                                                    '/html/body/div[6]/table/tbody/tr/td/div[2]/table[1]/tbody/tr[3]/td[2]').text

                    # Clica no botao fazer nova pesquisa
                    driver.find_element(By.XPATH,
                                        '/html/body/div[6]/table/tbody/tr/td/div[2]/table[2]/tbody/tr/td[1]/input').click()
                else:
                    # Limpa os campos de busca do Ipca
                    driver.find_element(By.XPATH, '//*[@id="oTab"]/td[1]/a').click()

                    valorCorrigidoIpcA.append('-')
        else:
            valorCorrigidoIpcA.append('-')


    # Adiciona os novos dados ao dataframe
    df1[f"Pouo.Nova {finalPoup}"] = valorCorrigidoPoup
    df1[f"Selic {finalSelic}"] = valorCorrigidoSelic
    df1[f"100% CDI {finalCdi}"] = valorCorrigidoCdi
    df1[f"IGP-M {finalIgpm}"] = valorCorrigidoIgpM
    df1[f"IPCA {finalIpca}"] = valorCorrigidoIpcA

    # Formata a data para forma brasileira para ser extraida para o excel
    df1["Data"] = pd.to_datetime(df1['Data'], format='%d%m%Y').dt.strftime('%d/%m/%Y')

    print(df1.to_string())


    # Adiciona os dados na planilha
    with pd.ExcelWriter(filename) as writer:
        df1.to_excel(writer, sheet_name="MySheet")
        auto_adjust_xlsx_column_width(df1, writer, sheet_name="MySheet", margin=0)


# Configuração da interface Tkinter
root = Tk()
root.title("Consulta Calculadora do Cidadão")

# Label e campo de entrada para o arquivo
arquivo_label = Label(root, text="Arquivo:")
arquivo_label.grid(row=0, column=0, padx=10, pady=5, sticky=W)

arquivo_entry = Entry(root, width=30)
arquivo_entry.grid(row=0, column=1, padx=10, pady=5)

arquivo_button = Button(root, text="Buscar", command=browse_file)
arquivo_button.grid(row=0, column=2, padx=5)

# Label e campo de entrada para a data
data_label = Label(root, text="Data final:")
data_label.grid(row=1, column=0, padx=10, pady=5, sticky=W)

cal = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
cal.grid(row=1, column=1, padx=10, pady=5)

# Botão de Executar e Sair
execute_button = Button(root, text="Executar", command=execute_and_close)
execute_button.grid(row=2, column=0, columnspan=3, padx=10, pady=10)

root.mainloop()