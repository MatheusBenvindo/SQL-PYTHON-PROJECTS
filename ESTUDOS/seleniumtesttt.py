from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
import time

# Caminho para o executável do Pale Moon
palemoon_path = r"C:\Users\matheus.ribeiro\Desktop\palemoon.exe"
geckodriver_path = r"C:\Users\matheus.ribeiro\Documents\geckodriver.exe"

# Configurações do Pale Moon
options = Options()
options.binary_location = palemoon_path
options.add_argument("--disable-web-security")
options.add_argument("--allow-running-insecure-content")
options.add_argument("--cors=no-cors")
service = Service(geckodriver_path)

# Inicializa o navegador
driver = webdriver.Firefox(service=service, options=options)

# Acessa o site que requer Flash
driver.get("https://lotr.creaction-network.com")

# Realiza outras operações necessárias
# ...

# Mantém o navegador aberto até ser interrompido manualmente
try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    # Fecha o navegador quando interrompido manualmente
    driver.quit()
