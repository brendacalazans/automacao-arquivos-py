import chromedriver_autoinstaller
from selenium import webdriver
from urllib.request import urlopen
import ssl
import json
ssl._create_default_https_context = ssl._create_unverified_context
from selenium.webdriver.firefox.options import Options
import json
import sys
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import secret
import time
from selenium.webdriver.firefox.options import FirefoxProfile
from datetime import datetime
from selenium.webdriver.common.action_chains import ActionChains
from bs4 import BeautifulSoup
import pdfkit


#chromedriver_autoinstaller.install(cwd=True)

class Browser:
    def resource_path(relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.dirname(__file__)
        return os.path.join(base_path, relative_path)

    def __init__(self, cod6: str, n_file: str):
        name_file = n_file + " - " + cod6 + ".pdf"
        final_file = "/Users/brendacalazans/Documents/novos documentos/" + name_file

        firefox_options = webdriver.FirefoxOptions()
        firefox_options.add_argument("--disable-infobars")
        firefox_options.add_argument("--disable-extensions")
        firefox_options.add_argument("--disable-popup-blocking")
        
        # Start printing the page without needing to press "print"
        firefox_options.set_preference("print.always_print_silent", True)
        firefox_options.set_preference("print.show_print_progress", False)
        firefox_options.set_preference('print.save_as_pdf.links.enabled', True)
        firefox_options.set_preference("pdfjs.disabled", True)
        firefox_options.set_preference("dom.disable_open_during_load", False)  # Desativa bloqueio de pop-ups


        # configurações da página
        firefox_options.set_preference("print.print_background", True)  # Imprime com fundo
        firefox_options.set_preference("print.margin.top", 0)  # Ajusta margens
        firefox_options.set_preference("print.margin.bottom", 0)
        firefox_options.set_preference("print.margin.left", 0)
        firefox_options.set_preference("print.margino w.right", 0)

        # Using print_printer
        profile = webdriver.FirefoxProfile()
        profile.set_preference("print_printer", "Microsoft Print to PDF")
        profile.set_preference("print.printer_Microsoft_Print_to_PDF.print_to_file", True)
        profile.set_preference('print.printer_Microsoft_Print_to_PDF.print_to_filename', final_file)

        firefox_options.profile = profile
        self.browser = webdriver.Firefox(options=firefox_options)
    
    def open_page(self, url: str):
        self.browser.get(url)
        
    def close_browser(self):
        self.browser.close()
        
    def add_input(self, by: By, value: str, text: str):
        field = self.browser.find_element(by=by, value = value)
        field.send_keys(text)
        time.sleep(1)
        
    def click_button(self, by: By, value: str):
        WebDriverWait(self.browser, 20).until(EC.element_to_be_clickable(
            self.browser.find_element(by=by, value = value))).click()
        time.sleep(1)

    def login_cmr(self, username: str, password: str):
        self.add_input(by=By.ID, value='username', text=username)
        self.add_input(by=By.ID, value='password', text=password)
        self.click_button(by=By.XPATH, value='/html/body/div[5]/div[4]/div/div[2]/div/div/div/div[1]/div[2]/div[1]/form/div/div/p/input')


    def search_cmr(self, cod6: str):
        self.click_button(by=By.XPATH, value='/html/body/div[5]/div[4]/div[1]/ul/li[2]/a')
        self.add_input(by=By.ID, value='customerNumber_multi', text=cod6)
        self.add_input(by=By.ID, value='cmr_cust_country', text='BR-Brazil')
        self.click_button(by=By.XPATH, value='/html/body/div[5]/div[5]/div/div[2]/div/form/div/div/div/div/div[1]/div/div/div[9]/div/div[2]/input[1]')

    def download_cmr(self):
            
        tds = self.browser.find_elements(By.TAG_NAME, 'td')

        for td in tds:
        # Encontre todos os links dentro do <td>
            links = td.find_elements(By.TAG_NAME, 'a')
            
            for link in links:
                # Obter o URL do link
                href = link.get_attribute('href')
                    
                    
                if href:
                   
                    original_window = self.browser.current_window_handle
                    link.click()

                        # Alterna para a nova janela
                    new_window = [window for window in self.browser.window_handles if window != original_window][0]
        
                    self.browser.switch_to.window(new_window)
                    time.sleep(1)
                    self.browser.execute_script('window.print()')

                    cust_number1 = WebDriverWait(self.browser,30).until(EC.presence_of_element_located((By.ID, "data.customerLegalName1"))).get_attribute("innerHTML")
                    cust_number2 = WebDriverWait(self.browser,30).until(EC.presence_of_element_located((By.ID, "data.customerLegalName2"))).get_attribute("innerHTML")

                    cnpj = WebDriverWait(self.browser,30).until(EC.presence_of_element_located((By.ID, "data.vat"))).get_attribute("innerHTML")
                    
                    cust_number = cust_number1 + " " + cust_number2
                    
                    dados = [cnpj, cust_number]

                    time.sleep(1)
                    self.browser.close()
                    self.browser.switch_to.window(original_window)

                    return dados
                else:
                    print("Erro. Não há links!")

                

    def login_w3(self, username: str, password: str):
        WebDriverWait(self.browser,20).until(EC.element_to_be_clickable((By.ID, 'credsDiv')))
        self.click_button(by=By.ID, value='credsDiv')
        time.sleep(1)
        self.add_input(by=By.ID, value='user-name-input', text=username)
        time.sleep(1)
        self.add_input(by=By.ID, value='password-input', text=password)
        time.sleep(1)
        self.click_button(by=By.ID, value='login-button')


    def search_gcs(self, cod6: str):
        element = WebDriverWait(self.browser,20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="SEARCH_METHOD"]')))
        element.send_keys('Legacy customer number')
        self.add_input(by=By.ID, value='SEARCH_DATA', text=cod6)
        self.click_button(by=By.ID, value='SUBMIT')

        original_window = self.browser.current_window_handle

        id_client = "CUSTOMER_RESULT_SELECTION_0"
        
        # Encontre todos os links dentro do <td>
    
        bolinha = WebDriverWait(self.browser,20).until(EC.element_to_be_clickable(self.browser.find_element(By.ID, id_client)))
        bolinha.click()


        enviar = WebDriverWait(self.browser,20).until(EC.element_to_be_clickable(self.browser.find_element(By.ID, "CUSTOMER_DETAILS_BUTTON")))
        enviar.click()

        new_window = [window for window in self.browser.window_handles if window != original_window][0]
        self.browser.switch_to.window(new_window)
        time.sleep(1)
        self.browser.execute_script('window.print()')

        time.sleep(1)
        self.browser.close()
        self.browser.switch_to.window(original_window)
            
       

    def send_ero(self, cod6: str, oppt: str, name: str):
        
        el = WebDriverWait(self.browser,20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/article/div/div/div/div[2]/div/div/div/div/div[7]/div/div')))
        el.click()
        
        div = self.browser.find_element(By.CLASS_NAME, "form-dropdown-field__options")

        self.browser.execute_script("arguments[0].setAttribute('style','overflow-y:scroll;overflow-x:hidden;height:1000px;')", div)

        element = WebDriverWait(self.browser,20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="Technology Lifecycle Services (TLS)-option"]')))
        element.click()
     


        self.add_input(by=By.ID, value="form-33ae06bf-f6ee-4ce2-89b1-f5d05f9f9f28-field-5", text=name)
        self.add_input(by=By.ID, value="form-33ae06bf-f6ee-4ce2-89b1-f5d05f9f9f28-field-6", text=cod6)
        self.add_input(by=By.ID, value="form-33ae06bf-f6ee-4ce2-89b1-f5d05f9f9f28-field-7", text=oppt)
        descr = "HWMA TSS - " + name
        self.add_input(by=By.ID, value="form-33ae06bf-f6ee-4ce2-89b1-f5d05f9f9f28-field-8", text=descr)


        respostas = self.browser.find_elements(By.CLASS_NAME, 'form-radio-field__input-control')

        cont = 0
        for resposta in respostas:
            cont += 1
            if(cont % 2 == 0):
                resposta.click()
        
        self.click_button(by=By.CLASS_NAME, value="form-checkbox-field__input-control")

        botao = WebDriverWait(self.browser,20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/article/div/div/div/div[2]/div/div/div/div/button')))
        botao.click()
    
    def login_outlook(self, username: str, password: str):
        valcon = True
        while (valcon):
            try:
                if (self.browser.find_element(By.ID, 'idSIButton9') is None):
                    time.sleep(1)
                    valcon = True
                else:
                    time.sleep(1)
                    valcon = False
            except:
                valcon = True

        usu = self.browser.find_element(By.ID, 'i0116')
        usu.clear()
        usu.send_keys(username)
        ing = self.browser.find_element(By.ID, 'idSIButton9')
        ing.click()
    
    def pegar_email(self, output_path, cod6):
        self.browser.maximize_window()
        element = WebDriverWait(self.browser,20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div/div/div/div[3]/div/div/div[1]/div[2]/div/div/div/div/div/div/div/div[2]/div/div/div/div/div[2]/div[2]/div[1]/div[1]")))
        element.click()
        original_window = self.browser.current_window_handle
        time.sleep(5)
        self.browser.execute_script("document.body.style.zoom='50%'")
        botao = WebDriverWait(self.browser,20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div/div/div/div[3]/div/div/div[3]/div/div/div/div[1]/div/div[2]/div/div/div[1]/div/div/div/div/div/div[3]/div[1]/div/div/div/table/tbody/tr/td/table[1]/tbody/tr[3]/td/center/table/tbody/tr/td")))
        time.sleep(2)
        botao.click()

        time.sleep(2)
        new_window = [window for window in self.browser.window_handles if window != original_window][0]
        self.browser.switch_to.window(new_window)
        time.sleep(10)
        self.browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")

        time.sleep(2)
        self.browser.execute_script('window.print()')

        time.sleep(1)
        self.browser.close()
        
        self.browser.switch_to.window(original_window)

if __name__ == '__main__':
    output_path = '/Users/brendacalazans/Documents/novos documentos'
    cod6 = "277602"
    oppt = "006Ka00000NV699IAD"

    # # ------------------------ CMR ------------------------

    print("CMR - Começar")

    # browser = Browser(cod6, "FindCMR")

    # browser.open_page('https://findcmr.epm2-prod.not-for-users.ibm.com/FindCMR/showLoginPage')
    # time.sleep(2)

    # browser.login_cmr(secret.email, secret.password)
    # time.sleep(1)
 
    # browser.search_cmr(cod6)
    
    # time.sleep(1)
    
    # dados = browser.download_cmr()
    # time.sleep(1)

    # browser.close_browser()

    # # ------------------------ GCS ------------------------

    # print("GCS - Começar")

    # browser = Browser(cod6, "GCS")
    # time.sleep(1)

    # browser.open_page('https://gcs-web-prod.dal1a.cirrus.ibm.com/financing/credit/protect/GCSGateway.wss?jadeAction=CUSTOMER_INQUIRY_INITIAL_ACTION_HANDLER&SessionTimeout=YES')
    # time.sleep(1)

    # browser.login_w3(secret.email, secret.password)
    
    # time.sleep(2)

    # browser.search_gcs(cod6)
    # time.sleep(1)

    # browser.close_browser()

    # # ------------------------ ERO ------------------------

    print("ERO - Começar")

    # browser = Browser(cod6, "ERO")
    # time.sleep(1)

    # browser.open_page('https://w3.ibm.com/w3publisher/ibm-export-regulation-office/delivery-of-client-solutions/client-services-evaluation-guide')
    # time.sleep(1)

    # browser.login_w3(secret.email, secret.password)
    # time.sleep(1)

    # name = str(dados[1])

    # browser.send_ero(cod6, oppt, name)

    # browser.close_browser()

    # ------------------------ Outlook ------------------------

    print("Outlook - Começar")

    browser = Browser(cod6, "ERO")
    time.sleep(1)

    browser.open_page('https://outlook.office.com/mail/')
    time.sleep(1)

    browser.login_outlook(secret.email, secret.password)
    time.sleep(1)

    browser.login_w3(secret.email, secret.password)
    time.sleep(1)

    browser.pegar_email(output_path, cod6)
    time.sleep(1)

    browser.close_browser()