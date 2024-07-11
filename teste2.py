from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
import secret
import time
import pdfkit



class Browser:
    browser, service = None,  None

    def __init__(self, driver: str):
        options = Options()
        options.headless = True
        self.browser = webdriver.Firefox()
        

    def open_page(self, url: str):
        self.browser.get(url)
    
    def close_browser(self):
        self.browser.close()
    
    def add_input(self, by: By, value: str, text: str):
        field = self.browser.find_element(by=by, value = value)
        field.send_keys(text)
        time.sleep(1)
    
    def click_button(self, by: By, value: str):
        WebDriverWait(browser, 20).until(EC.element_to_be_clickable(
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

    def download_cmr(self, cod6: str, output_path: str, wkhtmltopdf_path: str):
        
        tds = self.browser.find_elements(By.TAG_NAME, 'td')

        for td in tds:
        # Encontre todos os links dentro do <td>
            links = td.find_elements(By.TAG_NAME, 'a')
        
            for link in links:
                # Obter o URL do link
                href = link.get_attribute('href')
                
                
                if href:
                    # Navegar para o link
                    #self.browser.get(href)
                    #print(self.browser.get(href))
                    original_window = self.browser.current_window_handle

                    link.click()
                    time.sleep(5)

                    # Espera até que uma nova janela seja abert

                    # Alterna para a nova janela
                    new_window = [window for window in self.browser.window_handles if window != original_window][0]
                    self.browser.switch_to.window(new_window)

                    name_file = "FindCmr - " + cod6 + ".pdf"
                    # Aguarde um pouco para a página carregar (ajuste conforme necessário)
                    time.sleep(1)

                    # Enviar comando Cmd + P para abrir a caixa de diálogo de impressão
                    
                    final_file = output_path + "/" + name_file

                    page_source = self.browser.page_source
                    try:
                        config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)
                        pdfkit.from_string(page_source, final_file, configuration=config)
                        
                    except IOError as e:
                        print(f"Erro de IO: {e}")
                        if 'Done' not in str(e):
                            raise e
                    except Exception as e:
                        print(f"Erro inesperado: {e}")
                    finally:
                        time.sleep(3)
                    
                else:
                    print("Erro. Não há links!")

if __name__ == '__main__':
    output_path = '/Users/brendacalazans/Documents/novos documentos'
    wkhtmltopdf_path = '/usr/local/bin/wkhtmltopdf'
    browser = Browser('drivers/chromedriver')

    browser.open_page('https://findcmr.epm2-prod.not-for-users.ibm.com/FindCMR/showLoginPage')
    time.sleep(2)

    browser.login_cmr(secret.email, secret.password)
    time.sleep(1)
 
    browser.search_cmr('095076')
    time.sleep(1)
    
    browser.download_cmr('095076', output_path, wkhtmltopdf_path)
    time.sleep(1)

    print("Funcionando!")

    browser.open_page('https://gcs-web-prod.dal1a.cirrus.ibm.com/financing/credit/protect/GCSGateway.wss?jadeAction=CUSTOMER_INQUIRY_INITIAL_ACTION_HANDLER')