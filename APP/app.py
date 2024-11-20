import streamlit as st
from selenium import webdriver
from selenium.webdriver.common.by import By
from PIL import Image
from io import BytesIO
import pandas as pd
import time
import base64
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from defs_TIP import *  # Certifique-se de que o módulo defs_TIP está acessível e correto

st.set_page_config(page_title="Automação de Envio de Mensagens", page_icon="📲")

def conecta_whats(driver):
    time.sleep(10)
    driver.get("https://web.whatsapp.com/")
    time.sleep(10)
    qr_code_element = driver.find_element(By.XPATH, '//*[@id="app"]/div/div[2]/div[2]/div[1]/div/div/div[2]/div[2]/div[1]/canvas')
    qr_code_base64 = qr_code_element.screenshot_as_base64
    qr_code_data = base64.b64decode(qr_code_base64)
    qr_code_image = Image.open(BytesIO(qr_code_data))
    st.image(qr_code_image, caption="QR Code para WhatsApp Web")

def execute_process(sheet_df, attachment_file=None):
    driver = chama_driver_chrome()
    time.sleep(5)
    conecta_whats(driver)
    time.sleep(30)

    # Configuração da barra de progresso
    progress_bar = st.progress(0)
    total = len(sheet_df)
    count = 0  # Contador para atualizações de progresso

    for index, row in sheet_df.iterrows():
        telefone = row["Telefone"]
        status = row.get("Status", "")
        
        if status != "Enviado":
            st.write(f"Iniciando processo de envio para {telefone}")
            driver.get(f'https://wa.me/{telefone}')
            time.sleep(5)
            click(driver, '//*[@id="action-button"]')
            time.sleep(2)
            click(driver, '//*[@id="fallback_block"]/div/div/h4[2]/a/span')
            time.sleep(15)
            mensagem = 'teste'
            escrever(driver, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/div/div[1]/p', mensagem)
            time.sleep(1)
            click(driver, '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[2]/button/span')
            time.sleep(3)
            
            if attachment_file:
                attachment_icon = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//div[@title="Anexar"]')))
                time.sleep(1)
                attachment_icon.click()
                time.sleep(1)
                image_option = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')))
                time.sleep(1)
                image_option.send_keys(r"C:\Users\fernando.galves\Desktop\OIP.jpg")
                time.sleep(2)
                click(driver, '/html/body/div[1]/div/div/div[3]/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div/span')
                time.sleep(3)

            st.success(f"Enviado para {telefone}")
            sheet_df.at[index, "Status"] = "Enviado"

            # Atualiza o progresso
            count += 1
            progress_bar.progress(count / total)

            # Atualiza o link de download a cada envio
            output = BytesIO()
            sheet_df.to_excel(output, index=False, engine='xlsxwriter')
            output.seek(0)
            b64 = base64.b64encode(output.read()).decode()
            download_link = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="planilha_atualizada.xlsx">📥 Baixar planilha atualizada</a>'
            st.markdown(download_link, unsafe_allow_html=True)

    driver.close()
    return sheet_df

# Interface Streamlit
st.title("📲 Envio Automático de Mensagens no WhatsApp")

# Descrição
st.markdown("""
Bem-vindo! Este aplicativo permite enviar mensagens automaticamente via WhatsApp usando uma planilha com números de telefone.
Por favor, siga as instruções abaixo para realizar o envio.
""")

# Divisor
st.divider()

# Link de download do arquivo padrão
st.markdown("### 📄 Baixe o arquivo padrão")
st.markdown(
    '[Clique aqui para baixar o arquivo padrão](https://docs.google.com/spreadsheets/d/113DeuY-EMd2aE2_5Yu3qo-QEDa_09IHP/export?format=xlsx)',
    unsafe_allow_html=True
)

# Divisor
st.divider()

# Upload do arquivo de contatos
st.markdown("### 📂 Insira a planilha com os números de telefone")
file = st.file_uploader(
    label="Envie um arquivo Excel com uma coluna chamada 'Telefone' contendo os números",
    type=["xlsx"],
    help="Somente arquivos Excel com uma coluna 'Telefone' são aceitos."
)

# Opção de envio de arquivo adicional
arquivo = st.checkbox(label="📎 Deseja enviar um arquivo adicional?")
attachment_file = None

# Mostrar uploader de arquivo adicional se a opção for selecionada
if arquivo:
    attachment_file = st.file_uploader(label="Selecione o arquivo adicional para envio")

# Divisor
st.divider()

# Processamento do arquivo
if file:
    sheet_df = pd.read_excel(file)
    
    # Validação da coluna 'Telefone'
    if "Telefone" not in sheet_df.columns:
        st.error("⚠️ A coluna 'Telefone' é obrigatória na planilha. Por favor, verifique e tente novamente.")
    else:
        # Adicionar coluna 'Status' se não existir
        if "Status" not in sheet_df.columns:
            sheet_df["Status"] = ""
        
        # Botão de início
        if st.button("🚀 Iniciar Automação"):
            with st.spinner("Processando..."):
                sheet_df = execute_process(sheet_df, attachment_file)
            st.success("✅ Processo concluído com sucesso!")
