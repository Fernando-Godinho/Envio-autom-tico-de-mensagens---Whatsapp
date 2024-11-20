import win32com.client as client
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
from selenium.webdriver.common.by import By
import time
import datetime as dt
from datetime import datetime
from PIL import Image
import os
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
import win32gui
import win32con
import pandas as pd
import warnings
import smartsheet
import traceback
import re
import pyautogui


# Funções de automação
def escrever(campo, texto):
    time.sleep(3)
    try:
        element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, campo)))
        element.clear()
        element.send_keys(texto)
    except TimeoutException:
        print("Timeout: Elemento de entrada não encontrado. Ignorando e continuando.")
def dois_click(campo):
    try:
        time.sleep(5)
        element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, campo)))
        actions = ActionChains(driver)
        actions.double_click(element)
        actions.perform()
    except TimeoutException:
        print("Timeout: Elemento não encontrado. Ignorando e continuando.")
def carrega_dados(sheet):
    col_names = [col.title for col in sheet.columns]
    rows = []
    for row in sheet.rows:
        cells = []
        for cell in row.cells:
            cells.append(cell.value)
            rows.append(cells)
    data_frame = pd.DataFrame(rows, columns=col_names)
    return data_frame

def acessar_relatorio(element):
    iframe_busc = driver.find_element(By.XPATH, element)
    iframe_busc.click()
    driver.switch_to.frame(iframe_busc)   

def usuario():
    escrever("/html/body/form/div[1]/div/div[1]/div/div/div/div/div[4]/div/div[1]/div/div/div[2]/div[1]/div/div/input", "edinei.l.silva")
    escrever("/html/body/form/div[1]/div/div[1]/div/div/div/div/div[4]/div/div[1]/div/div/div[3]/div[1]/div/div/input", "Campseg2024@")
    click("/html/body/form/div[1]/div/div[1]/div/div/div/div/div[4]/div/div[2]/div/div/a/span/span/span[2]")
    dois_click("/html/body/div[1]/div[2]/div[4]/div/div/a[6]/span/span/span[2]")

def click(campo):
    try:
        time.sleep(5)
        element = WebDriverWait(driver, 50).until(EC.presence_of_element_located((By.XPATH, campo)))
        element.click()
    except TimeoutException:
        print("Timeout: Elemento não encontrado. Ignorando e continuando.")
    
def click_webelement(xpath, indice):
    lista_campos = driver.find_elements(By.XPATH, xpath)
    campo_a_ser_clicado = lista_campos[indice]
    try:
        time.sleep(5)
        WebDriverWait(driver, 50).until(EC.visibility_of(campo_a_ser_clicado))
        campo_a_ser_clicado.click()
    except TimeoutException:
        print("Timeout: Elemento não encontrado. Ignorando e continuando.")
def get_latest_file_path(folder_path):
    file_paths = [os.path.join(folder_path, filename) for filename in os.listdir(folder_path)]
    file_paths = [path for path in file_paths if os.path.isfile(path)]
    latest_file_path = max(file_paths, key=os.path.getmtime)
    return latest_file_path
    
def is_recent_file(file_path, minutes=4):
    file_mod_time = datetime.fromtimestamp(os.path.getmtime(file_path))
    return datetime.now() - file_mod_time < timedelta(minutes=minutes)

def envio_de_data(campo_de_data,competencia_formatada):
    try:
        # Encontre o elemento de entrada de data pelo ID
        input_element = driver.find_element(By.ID, campo_de_data)
        #Execute um script JavaScript para definir o valor do campo de entrada
        driver.execute_script(f"arguments[0].value = {competencia_formatada}", input_element)
        # Pressione a tecla Enter para confirmar a entrada (opcional)
        input_element.send_keys(Keys.ENTER)
        # Imprima uma mensagem de sucesso
        print("Data enviada com sucesso!")
    except Exception as e:
        # Em caso de erro, imprima a mensagem de erro\n",
        print(f"Erro ao enviar a data: {e}")

def limpa_pasta(caminho_pasta):
    
    itens_pasta = os.listdir(r'C:\\Users\\hugo.clemente\\Downloads')
    if len(itens_pasta) > 0:
        for item in itens_pasta:
            os.remove(os.path.join(caminho_pasta, item))

def update_smartsheet(column_title, cell_value, row_index):
    global smart
    global sheet_id
    global max_rows

    if row_index < max_rows:
        try:
            sheet = smart.Sheets.get_sheet(sheet_id)
            rows = sheet.rows
            row_id = rows[row_index].id
            column_id = [col.id for col in sheet.columns if col.title == column_title][0]
            cell = smartsheet.models.Cell()
            cell.column_id = column_id
            cell.value = cell_value
            updated_row = smartsheet.models.Row()
            updated_row.id = row_id
            updated_row.cells.append(cell)
            response = smart.Sheets.update_rows(sheet_id, [updated_row])
    
            if response.message == 'SUCCESS':
                print(f'Valor \"{cell_value}\" atualizado na coluna \"{column_title}\" com sucesso.')
            else:
                print('Falha ao atualizar o valor na coluna.')
        except Exception as e:
            print(f'Erro: {e}')
    else:
        print('Índice da linha fora do intervalo.')
    return row_index       
    
import requests
from selenium import webdriver

def download_pdf_from_blob(blob_url, output_file_path):
    # Use requests to directly download the PDF
    response = requests.get(blob_url)
    if response.status_code == 200:
        with open(output_file_path, 'wb') as file:
            file.write(response.content)
            print(f"PDF saved at: {output_file_path}")
    else:
        print(f"Failed to download PDF. Status code:{response.status_code}")

#token
access_token = 'qGPQ1qbfWCpw07VytHUObAjZj5h2cAvHOBUet'
# Conecta à API
smart = smartsheet.Smartsheet(access_token)
# Id da planilha Smartsheet
sheet_id = '4739764502613892'
#Assume valores da Planilha
sheet = smart.Sheets.get_sheet(sheet_id)
#Carrega para o dataframe df
df = carrega_dados(sheet)
row_index = 0
max_rows = len(df)
# Define a duração do loop em segundos (8 horas)
duration = 480 * 60
# Loop que roda por 8 horas
start_time = time.time()

while time.time() - start_time < duration:
    try: 
        print("--------------INICIOU O PROCESSO!--------------")
        # Recarregar o DataFrame a cada iteração
        sheet = smart.Sheets.get_sheet(sheet_id)
        df = carrega_dados(sheet)
        max_rows = len(df)

        if row_index >= max_rows:
            row_index = 0
    
        print("--------------INICIANDO AUTOMAÇÃO--------------")
        # Assume nome das variaveis
        cpf = str(df['CPF'].iloc[row_index])
        cpf = cpf.replace('.', '').replace('-', '')
        print(cpf)
        competencia = str(df['Periodo'].iloc[row_index]) 
        ano, mes, dia = competencia.split('-')
        competencia_formatada = mes+"/"+ano
        competencia_formatada = str(competencia_formatada)
        link_pdf = str(df['Link PDF'].iloc[row_index])
        cr = str(df['CR'].iloc[row_index])
        print(cpf, competencia_formatada, link_pdf, cr)
        time.sleep(2)

        if cr != "Enviado":
            #Define the URL and output file path
            pdf_url = link_pdf
            output_path = r'C:\\Users\\hugo.clemente\\Downloads\\downloaded_pdf.pdf'
            # Download the PDF
            download_pdf_from_blob(pdf_url, output_path)
            driver = webdriver.Edge()
            time.sleep(5)
            driver.get("https://portal.gpssa.com.br/GPS/Login.aspx")
            driver.maximize_window()
            usuario()
            time.sleep(5)
            print("--------------INICIANDO PORTAL GPS--------------")
            dois_click("//*[text()='Gestão de Pessoas']")
            dois_click("//*[text()='4. Documentos e Contratos']")
            dois_click("//*[text()='4.1 - Card - Gestão do Colaborador']")
            time.sleep(5)
            acessar_relatorio("/html/body/form/div[4]/div/div/div/div/div[3]/div/div[2]/div[2]/div[1]/div/iframe")
            print("ok iframe acessado")
            browser_tabs = driver.window_handles
            driver.switch_to.window(browser_tabs[0])
            acessar_relatorio("/html/body/form/div[4]/div/div/div/div/div[3]/div/div[2]/div[2]/div[1]/div/iframe")
            time.sleep(5)
            escrever("//*[@ID='txtPesquisa-inputEl']", cpf)
            click("//*[@ID='btnPesquisar-btnWrap']")
            time.sleep(8)
            click("/html/body/div[1]/div[2]/div[3]/div/div[2]/table/tbody/tr/td[1]/div/div")
            time.sleep(8) 
            browser_tabs = driver.window_handles
            driver.switch_to.window(browser_tabs[-1])
            time.sleep(2)
            click("//*[@ID='ext-element-7']")
            click("//*[text()='5.80 - INCLUIR DOCUMENTO']")
            print("debug1")
            acessar_relatorio("/html/body/div[2]/div/div/div/div/div/div/div/div/div/div[2]/div[2]/div[1]/div/iframe")
            print("debug2")
            click("/html/body/div[5]/div[2]/div/div/div/div/div/div/div/div/div[1]/div[1]/div[1]/div[2]")
            print("debug3")
            click("//*[text()='FOLHA DE PONTO']")
            print("debug4")
            click("//*[@ID='cbbTiposDocumentos-trigger-picker']")
            print("debug5")
            click("//*[text()='FOLHA DE PONTO MANUAL COM COMPETENCIA']")
            print("debug6")
            click('//*[@id=\"dtCompetencia2627-inputEl\"]')
            print("debug7")
            input_element = driver.find_element(By.ID, "dtCompetencia2627-inputEl")
            # Execute um script JavaScript para definir o valor do campo de entrada
            driver.execute_script(f"arguments[0].value = '{competencia_formatada}';", input_element)
            time.sleep(5)
            print("colocou o doc no form")
            print("finalizou")
            time.sleep(3)
            downloads_path = os.path.join(os.path.expanduser("~"),"Downloads")
            ultimo_arquivo = get_latest_file_path(downloads_path)
            input_element = driver.find_element(By.XPATH, "//*[@id='fupNovoDocumento-button-fileInputEl']")
            input_element.send_keys(ultimo_arquivo)
            click('//*[@id="btnSalvar-btnEl\"]')
            time.sleep(27)
            update_smartsheet("CR", "Enviado", row_index)
            print("Enviado")
            limpa_pasta('C:\\Users\\hugo.clemente\\Downloads')
            driver.quit()
            row_index += 1
    except Exception as e:
            print("erro:", e)
            update_smartsheet("CR","Erro", row_index)
            time.sleep(3)
            row_index += 1
            driver.quit()
           