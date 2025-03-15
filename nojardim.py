import os
import time
import pandas as pd
from dotenv import load_dotenv
from openpyxl import load_workbook
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager

load_dotenv("cripto.env")

hoje = datetime.today()
primeiro_dia_mes = hoje.replace(day=1)
ultimo_dia_mes = (primeiro_dia_mes.replace(month=hoje.month % 12 + 1, day=1) - timedelta(days=1))

email = os.getenv("emailtakeat")
password = os.getenv("senhatakeat")

data = datetime.now()
mes_atual = f"{data.month:02}"
print(mes_atual)
planilha = fr"C:\Users\sigab\Downloads\Produtos(01-{mes_atual}_31-{mes_atual}).xlsx"
caminho = f'C:\\Users\\sigab\\OneDrive - Siga Financeiro e Controladoria\\AUTOMACAO\\produtos.xlsx'

def configurar_driver():
    edge_options = webdriver.EdgeOptions()
    prefs = {
        "download.default_directory": r"C:\Users\sigab\Downloads",
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    edge_options.add_experimental_option("prefs", prefs)
    edge_options.add_argument("--start-maximized")
    
    driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options=edge_options)
    return driver

def realizar_login(driver):
    driver.get("https://gestor.takeat.app/login")
    
    # Aguarda o campo de e-mail estar presente
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//input[@placeholder="E-mail"]')))
    email_field = driver.find_element(By.XPATH, '//input[@placeholder="E-mail"]')
    email_field.send_keys(email)

    senha_field = driver.find_element(By.XPATH, '//input[@placeholder="Senha"]')
    senha_field.send_keys(password)

    button_acesso = driver.find_element(By.XPATH, '//button[@type="button"]//span[text()="Acessar"]')
    button_acesso.click()

    wait = WebDriverWait(driver, 20)
    relatorio_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[span[text()='Relatórios']]")))
    relatorio_link.click()

    wait = WebDriverWait(driver, 20)
    elemento = wait.until(EC.visibility_of_element_located((By.XPATH, "//*[contains(text(), 'Vendidos')]")))
    elemento.click()

    time.sleep(60)

    wait = WebDriverWait(driver, 20)
    botao_baixar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Baixar']]")))
    botao_baixar.click()

def tratar_planilha(planilha, data):

    global arquivo_saida

    df = pd.read_excel(planilha, sheet_name="Relatório Produtos")

    df.insert(0, "Período", data)
    df.insert(1,"Tipo","")
    df.insert(len(df.columns), "PIZZA de Dois Sabores", "")

    arquivo_saida = r"C:\Users\sigab\Downloads\meu_arquivo.xlsx"
    df.to_excel(arquivo_saida, index=False, engine='openpyxl')

def concatenar_planilhas(arquivo_saida, caminho):

    global df_concatenado

    planilha1 = pd.read_excel(arquivo_saida)
    planilha2 = pd.read_excel(caminho)

    planilha2.rename(columns={'Tipo 1': 'Tipo'}, inplace=True)
    

    print("Colunas da planilha 1:", planilha1.columns)
    print("Colunas da planilha 2:", planilha2.columns)

    ultima_linha_df = planilha2['Período'].last_valid_index() + 2
    print(f"Última linha válida na aba 'produtos': {ultima_linha_df}")

    df_concatenado = pd.concat([planilha2.iloc[: ultima_linha_df], planilha1], ignore_index=True)

    with pd.ExcelWriter(caminho, engine='openpyxl') as writer:
        df_concatenado.to_excel(writer, index=False, sheet_name="Relatório Produtos")

def gerar_arquivo(df_concatenado):

    wb = load_workbook(caminho, data_only=False)

    ws = wb["Relatório Produtos"] 

    for row in range(2, len(df_concatenado) + 2):  
        valor_c = ws[f'C{row}'].value
        valor_d = ws[f'D{row}'].value
        
        if valor_c == ws["N1"].value: 
            ws[f'B{row}'] = "Normal"
        elif valor_d and valor_d[:4] == "  - ": 
            ws[f'B{row}'] = "Complemento"
        else:
            ws[f'B{row}'] = "Normal"


        data = ws[f'A{row}'].value  
        if isinstance(data, datetime):  
            periodo = data.strftime("%B-%Y").lower()  
            periodo = periodo.replace("january", "janeiro").replace("february", "fevereiro").replace("march", "março") \
                .replace("april", "abril").replace("may", "maio").replace("june", "junho").replace("july", "julho") \
                .replace("august", "agosto").replace("september", "setembro").replace("october", "outubro") \
                .replace("november", "novembro").replace("december", "dezembro")
            
            ws[f'A{row}'] = periodo 
                       
    time.sleep(60)

    wb.save(caminho)

def deletar_arquivo(planilha, arquivo_saida):
    if os.path.exists(planilha):
        os.remove(planilha)
        print(f"Arquivo '{planilha}' deletado com sucesso.")
    elif os.path.exists(arquivo_saida):
        os.remove(arquivo_saida)
        print(f"Arquivo '{arquivo_saida}' deletado com sucesso.")
    else:
        print("Nenhum dos arquivos foi encontrado para deletar.")


driver = configurar_driver()
configurar_driver()
realizar_login(driver)
tratar_planilha(planilha, data)
concatenar_planilhas(arquivo_saida, caminho)
gerar_arquivo(df_concatenado)
deletar_arquivo(planilha, arquivo_saida)

print("Processo concluído.")
driver.quit()

