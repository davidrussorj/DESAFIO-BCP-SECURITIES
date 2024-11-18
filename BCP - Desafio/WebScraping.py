import os
import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from datetime import date

# Lista de feriados fixos
HOLIDAYS = [
    date(2024, 11, 15),  
    date(2024, 4, 21), 
    date(2024, 5, 1),  # Dia do Trabalho
    # Adicione outros feriados fixos e móveis aqui
]

# Função modificada para excluir feriados
def get_last_business_days(n=5):
    business_days = []
    current_date = datetime.today() - timedelta(days=1)
    while len(business_days) < n:
        if current_date.weekday() < 5 and current_date.date() not in HOLIDAYS:
            business_days.append(current_date)
        current_date -= timedelta(days=1)
    return business_days

# Nome da pasta onde os arquivos serão salvos
FOLDER_NAME = "Daily Prices"
URL_TEMPLATE = "https://www.anbima.com.br/informacoes/merc-sec-debentures/resultados/mdeb_{date}_di_percentual.asp"

# Função para criar a pasta se ela não existir
def create_folder(folder_name):
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

# Função para obter os últimos 'n' dias úteis
# Função modificada para excluir feriados
def get_last_business_days(n=5):
    business_days = []
    current_date = datetime.today() - timedelta(days=1)
    while len(business_days) < n:
        if current_date.weekday() < 5 and current_date.date() not in HOLIDAYS:
            business_days.append(current_date)
        current_date -= timedelta(days=1)
    return business_days

# Função para formatar a data no estilo "ddmmmaaaa"
def format_date_for_url(date):
    return date.strftime("%d%b%Y").lower()

# Função para baixar o arquivo .xls de uma data específica
def download_xls_from_page(url, date):
    try:
        response = requests.get(url)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.content, "html.parser")
        xls_link = soup.find("a", class_="linkinterno", href=lambda href: href and href.endswith(".xls"))
        
        if xls_link:
            xls_url = "https://www.anbima.com.br/informacoes/merc-sec-debentures/" + xls_link['href'][2:]
            file_name = f"{date.strftime('%Y%m%d')}.xls"
            file_path = os.path.join(FOLDER_NAME, file_name)
            
            xls_response = requests.get(xls_url)
            xls_response.raise_for_status()
            
            with open(file_path, "wb") as file:
                file.write(xls_response.content)
            print(f"Arquivo salvo: {file_name}")
        else:
            print(f"Link para o arquivo .xls não encontrado na página de {date.strftime('%Y-%m-%d')}")
    
    except requests.exceptions.RequestException as e:
        print(f"Erro ao acessar a página para {date.strftime('%Y-%m-%d')}: {e}")

#Cria a pasta onde os arquivos serão salvos
create_folder(FOLDER_NAME)
    
# Obtém os últimos 5 dias úteis
dates = get_last_business_days()
# Baixa os arquivos de cada data
for date in dates:
    formatted_date = format_date_for_url(date)
    page_url = URL_TEMPLATE.format(date=formatted_date)
    download_xls_from_page(page_url, date)
