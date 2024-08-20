# Macro criada para preenchimento de informações em uma plataforma online. 
import openpyxl
import tkinter as tk
from tkinter import messagebox, simpledialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import threading
import json
import logging
from selenium.webdriver.edge.options import Options
import time

# Configuração do logger
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Função para mostrar a mensagem de conclusão
def mostrar_mensagem():
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal
    root.focus_force()  # Força o foco na janela do Tkinter
    root.attributes("-topmost", True)  # Garante que a janela esteja no topo
    messagebox.showinfo("Informação", "Tarefa executada com sucesso!=)")
    root.destroy()

# Função para carregar o caminho do arquivo do arquivo de configuração
def carregar_configuracao():
    try:
        with open('config.json', 'r', encoding='utf-8') as f:
            config = json.load(f)
        return config.get('caminho_arquivo_excel', ''), config.get('usuario', ''), config.get('senha', '')
    except Exception as e:
        logging.error(f"Erro ao carregar configuração: {e}")
        raise

try:
    # Carregar o caminho do arquivo Excel e credenciais do arquivo de configuração
    caminho_arquivo, valor_usuario, valor_senha = carregar_configuracao()
    if not caminho_arquivo:
        raise Exception("Nenhum caminho de arquivo especificado no arquivo de configuração")
    if not valor_usuario ou not valor_senha:
        raise Exception("Usuário ou senha não fornecidos")

    # Carregar a planilha
    workbook = openpyxl.load_workbook(caminho_arquivo)
    sheet = workbook.active

    logging.info(f'Usuário: {valor_usuario}')

    # Configurar opções do navegador
    options = Options()
    options.add_argument("start-maximized")  # Iniciar o navegador maximizado
    options.add_argument("disable-infobars")  # Desabilitar infobars
    options.add_argument("disable-extensions")  # Desabilitar extensões

    # Abrir navegador com as opções configuradas
    navegador = webdriver.Edge(options=options)
    navegador.get("URL_DO_SITE")

    # Login
    try:
        navegador.find_element(By.XPATH, '//*[@id="Conteudo_ctrLogin_UserName"]').send_keys(valor_usuario)
        navegador.find_element(By.XPATH, '//*[@id="Conteudo_ctrLogin_Password"]').send_keys(valor_senha)
        navegador.find_element(By.XPATH, '//*[@id="Conteudo_ctrLogin_Login"]').click()
    except Exception as e:
        logging.error(f"Erro ao fazer login: {e}")

    # Esperar a página carregar
    WebDriverWait(navegador, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="elemento"]')))

    # Localizar o elemento dropdown
    dropdown_element = WebDriverWait(navegador, 15).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="elemento"]'))
    )

    # Criar uma instância da classe Select
    select = Select(dropdown_element)

    # Loop pelas células da coluna
    for row in range(3, sheet.max_row + 1):
        valor_CD = sheet[f'C{row}'].value
        valor_SKU = sheet[f'A{row}'].value
        valor_PEN = sheet[f'D{row}'].value
        valor_PRAZO = sheet[f'E{row}'].value
        
        # Adicionar mensagens de depuração
        print(f"linha {row}: valor_CD={valor_CD}, valor_SKU={valor_SKU}, valor_PEN={valor_PEN}, valor_PRAZO={valor_PRAZO}")
        
        if valor_CD:  # Verifica se a célula não está vazia
            time.sleep(1)  # Tempo de espera para carregar a página
            select.select_by_visible_text(valor_CD)
            WebDriverWait(navegador, 5).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="idsku"]'))
            )
            
            # Preencher os campos
            time.sleep(1)
            if valor_SKU is not None:
                navegador.find_element(By.XPATH, '//*[@id="idsku"]').send_keys(str(valor_SKU))
            else:
                print(f"valor_SKU na linha {row} está vazio ou é None")
            time.sleep(1)
            if valor_PRAZO is not None:
                navegador.find_element(By.XPATH, '//*[@id="prazoabastecimento"]').send_keys(str(valor_PRAZO))
            else:
                print(f"valor_PRAZO na linha {row} está vazio ou é None")
            time.sleep(1)
            if valor_PEN is not None:
                navegador.find_element(By.XPATH, '//*[@id="quantidade"]').send_keys(str(valor_PEN))
            else:
                print(f"valor_PEN na linha {row} está vazio ou é None")

            navegador.find_element(By.XPATH, '//*[@id="collapse1"]/div[2]/div/div/a[2]/span').click()
            WebDriverWait(navegador, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idsku"]')))
            time.sleep(3)

            # Limpar os campos (LIMPA OS CAMPOS PREENCHIDOS PELO PY)
            navegador.find_element(By.XPATH, '//*[@id="idsku"]').clear()
            navegador.find_element(By.XPATH, '//*[@id="prazoabastecimento"]').clear()
            navegador.find_element(By.XPATH, '//*[@id="quantidade"]').clear()

finally:
    # Fechar o navegador
    navegador.quit()
    
    print("Todas as tarefas foram executadas.")

    # Mostrar a mensagem de conclusão em uma thread separada
    thread = threading.Thread(target=mostrar_mensagem)
    thread.start()
