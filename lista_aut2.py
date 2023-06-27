from configparser import ConfigParser
import logging
import keyring
import pandas as pd
# import openpyxl
import os
import time
from selenium import webdriver
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.proxy import Proxy, ProxyType

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
#from selenium.webdriver.firefox.service import Service as FirefoxService
#from webdriver_manager.firefox import GeckoDriverManager

ROOT_PATH = r"C:\Users\FernandoSimoes\OneDrive - kyndryl\Área de Trabalho\code\lista_autorizados"
LOG_PATH = os.path.join(ROOT_PATH, "Logs")
CONFIG_FILE = os.path.join(ROOT_PATH, "config.ini")


def get_browser():
    config = ConfigParser()
    config.read(CONFIG_FILE)
    geckodriver_path = config['settings']['geckodriver_path']
    firefox_profile_path = config['settings']['firefox_profile_path']
    options =webdriver.FirefoxOptions()
    proxy_server_url = "fw1.bes.gbes:8080"

    """   
    prox = Proxy()
    prox.proxy_type = ProxyType.MANUAL
    prox.http_proxy = proxy_server_url
    prox.socks_proxy = proxy_server_url
    prox.ssl_proxy = proxy_server_url
    prox.socksVersion=5

    capabilities = webdriver.DesiredCapabilities.FIREFOX
    prox.add_to_capabilities(capabilities)
    """
#    driver = webdriver.Firefox(desired_capabilities=capabilities)
    
#    options.add_argument('--headless')
    options.add_argument("--window-size=1920,1080")
    options.profile=firefox_profile_path
    options.set_preference("security.enterprise_roots.enabled", True)
#    fp = webdriver.FirefoxProfile(firefox_profile_path)
#    fp.set_preference("security.enterprise_roots.enabled", True)
    browser = webdriver.Firefox(
        executable_path=geckodriver_path,
#        desired_capabilities=capabilities
#        firefox_profile=fp,
#        service=FirefoxService(GeckoDriverManager().install()),

        options=options,
#        proxy=proxy
    )
    
    return browser


def refresh_groups(browser):
    WebDriverWait(browser,15).until(EC.visibility_of_all_elements_located((By.XPATH, '//img[@data-automation-id="HeroImage"]')))
    element = browser.find_element(By.XPATH,'(//img[@data-automation-id="HeroImage"])[1]')
    browser.execute_script("return arguments[0].scrollIntoView();", element)
    time.sleep(2)
    element = browser.find_element(By.XPATH,'(//img[@data-automation-id="HeroImage"])[6]')
    browser.execute_script("return arguments[0].scrollIntoView();", element)
    groups = browser.find_elements(By.XPATH, '//img[@data-automation-id="HeroImage"]')
    return groups


def refresh_areas(browser):
    areas = WebDriverWait(browser,15).until(EC.visibility_of_all_elements_located((By.XPATH, '//div[@data-automation-id="quick-links-item-title"]')))
    browser.execute_script("window.scrollTo(0,900);")
    time.sleep(2) 
#    browser.execute_script("window.scrollTo(451,1000)")  
#    time.sleep(2)
    areas = WebDriverWait(browser,15).until(EC.visibility_of_all_elements_located((By.XPATH, '//div[@data-automation-id="quick-links-item-title"]')))
    return areas

def refresh_nucleos(browser):
    nucleos = WebDriverWait(browser,25).until(EC.visibility_of_all_elements_located((By.XPATH, '//div[@data-automation-id="quick-links-item-title"]')))
    browser.execute_script("window.scrollTo(0,900);")
    time.sleep(2) 
#    browser.execute_script("window.scrollTo(451,721)")
    nucleos = WebDriverWait(browser,15).until(EC.visibility_of_all_elements_located((By.XPATH, '//div[@data-automation-id="quick-links-item-title"]')))
    return nucleos

def main():
    
    logging.basicConfig(filename=LOG_PATH + '\list_aut.log', filemode='a', \
    format='%(asctime)s : %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p', level=logging.INFO)
    logging.info('--------------------------------------------------------------------------')
    
    df = pd.DataFrame(columns=['Direcao','Equipa','PoolSuporte', 'ITSMassetID', 'CodigoAplicacao','Autorizado','EMail','User'])

    logging.info("GET CREDENTIALS")
    username = "t04292@novobanco.pt"    #"T04292"
    password = keyring.get_password("system", username)


    logging.info('GET WEBDRIVER')
    browser = get_browser()
    browser.maximize_window()
    url = r'https://bdso.sharepoint.com/sites/sup_aut'
    browser.get(url)

    user_ini=WebDriverWait(browser, 15).until(EC.element_to_be_clickable((By.XPATH, '//input[@id= "i0116"]')))
    user_ini.send_keys(username)
    btn_next = browser.find_element(By.XPATH, '//input[@id="idSIButton9"]')
    btn_next.click()

    pass_ini=WebDriverWait(browser, 15).until(EC.element_to_be_clickable((By.XPATH, '//input[@id="passwordInput"]')))
    pass_ini.send_keys(password)

    btn_login=browser.find_element(By.XPATH, '//span[@id="submitButton"]')
    btn_login.click()

    wind_btn=WebDriverWait(browser, 15).until(EC.element_to_be_clickable((By.XPATH, '//input[@id="idSIButton9"]')))
    wind_btn.click()
    logging.info('LOGIN SUCCESSFUL')
  
    

    conta=0
    cntg=0
    groups= refresh_groups(browser)
    while cntg < len(groups):  
        browser.execute_script("return arguments[0].scrollIntoView();",groups[cntg]) 
#        print(groups)     
        groups[cntg].click()
        group = WebDriverWait(browser, 15).until(EC.element_to_be_clickable((By.XPATH,'//a[contains(@class,"title-")]'))).text

        cnta=0
        areas=refresh_areas(browser)
        while cnta < len(areas):
            browser.execute_script("return arguments[0].scrollIntoView();",areas[cnta])
            area =areas[cnta].text
            areas[cnta].click()
            cntn=0
            nucleos = refresh_nucleos(browser)
            while cntn < len(nucleos):
                nucleos = refresh_nucleos(browser)
                nucleo = nucleos[cntn].text
                nucleos[cntn].click()

                try:
                    aut = WebDriverWait(browser, 15).until(EC.element_to_be_clickable((By.XPATH, '//span[text()="Autorizados"]')))
                    aut.click()
                except NoSuchElementException:
                    print('Menu Autorizados inexistente.')
                    cntn+=1
                    WebDriverWait(browser, 15).until(EC.element_to_be_clickable((By.XPATH, '//a[@title="Área"]/span'))).click()
                    time.sleep(2)
                    continue

                try:
                    auts=WebDriverWait(browser,10).until(EC.visibility_of_all_elements_located((By.XPATH, '//button[contains(@class,"od-FieldRender-lookup")]')))
                except TimeoutException:
                    print('Sem autorizados.')
                    cntn+=1
                    WebDriverWait(browser, 15).until(EC.element_to_be_clickable((By.XPATH, '//a[@title="Área"]/span'))).click()
                    time.sleep(2)
                    continue


                 
                if auts!=[]:
                    
                    WebDriverWait(browser, 25).until(EC.element_to_be_clickable((By.XPATH,'//span[contains(text(),"Aplicações Suportadas")]'))).click()
                    try:
                        cas=WebDriverWait(browser, 10).until(EC.visibility_of_all_elements_located((By.XPATH,'//button[contains(@data-automationid,"FieldRenderer-name")]')))
                        cod=WebDriverWait(browser, 10).until(EC.visibility_of_all_elements_located((By.XPATH,'//div[contains(@class,"od-FieldRenderer-text")]')))
                        pool=WebDriverWait(browser, 10).until(EC.visibility_of_all_elements_located((By.XPATH,'//button[contains(@class,"od-FieldRender-lookup")]')))
                    except TimeoutException:
                        cntn +=1
                        print('Sem aplicações!')
                        WebDriverWait(browser, 15).until(EC.element_to_be_clickable((By.XPATH, '//a[@title="Área"]/span'))).click()
                        continue
#                    print(len(cas),len(cod),len(pool))
                    tups=[]
                    for index in range(len(cas)):
                        tups.append((cas[index].text,cod[index].text,pool[index].text))


                    WebDriverWait(browser, 15).until(EC.element_to_be_clickable((By.XPATH,'//a[@title="Contactos"]'))).click()
                    nome=WebDriverWait(browser, 15).until(EC.visibility_of_all_elements_located((By.XPATH,'//button[@data-automationid="FieldRenderer-name"]')))
                    try:
                        email=WebDriverWait(browser, 10).until(EC.visibility_of_all_elements_located((By.XPATH,'//div[@class="sp-field-customFormatter"]/a')))
                    except TimeoutException:
                        email=[]
                        ema=WebDriverWait(browser, 15).until(EC.visibility_of_all_elements_located((By.XPATH,'//div[@class="od-FieldRenderer-text fieldText_875b1af5"]')))

                        for em in ema:
                            if em.location['x'] ==  668:
                                email.append(em)
                        pass
                    user=WebDriverWait(browser, 15).until(EC.visibility_of_all_elements_located((By.XPATH,'//div[contains(@class,"field_875b1af5")]')))
                    nomes=[]
#                    print(len(nome),len(email),len(user))
                    red=0
                    for index3 in range(len(nome)):
#                        print(index3,red)
                        for windex in range(len(user)):
                            if nome[index3].text.split(" ")[0].lower() == user[windex].text.split(" ")[0].lower():
                                nomes.append((nome[index3].text,email[index3].text,user[windex].text))
                                red=1
                                break
                            
                        if red == 0: 
                            nomes.append((nome[index3].text,email[index3].text,"")) 
                        else:
                            red = 0
                            
                    aut = WebDriverWait(browser, 15).until(EC.element_to_be_clickable((By.XPATH, '//span[text()="Autorizados"]'))).click()
                    auts=WebDriverWait(browser, 15).until(EC.visibility_of_all_elements_located((By.XPATH,'//button[contains(@class,"od-FieldRender-lookup")]')))
#                        print(auts)
#                        print('passa aqui')
                    

                    for index2 in range(len(auts)):
#                            XY=auts[index2].location
#                            print(str(XY['x']))
                        if auts[index2].location['x'] == 368:
                            apl=auts[index2].text
#                                print(apl)
                            continue
                        for tup in tups:
#                                print(tup[2]+"|"+apl+"|")

                            if tup[2] in apl:
                                conta+=1
                                aut_line=group+" | "+nucleo+" | "+tup[2]+" | "+tup[0]+" | "+tup[1]+" | "+auts[index2].text
                                df.at[conta,'Direcao'] = group
                                df.at[conta,'Equipa'] = nucleo
                                df.at[conta,'PoolSuporte'] = tup[2]
                                df.at[conta,'ITSMassetID'] = tup[0]
                                df.at[conta,'CodigoAplicacao'] = tup[1]
                                df.at[conta,'Autorizado'] = auts[index2].text
                                for nom in nomes:
                                    if nom[0] == auts[index2].text:
                                        aut_line = aut_line+" ! " + nom[1] + " ! "+nom[2]
                                        df.at[conta,'EMail'] = nom[1]
                                        df.at[conta,'User'] = nom[2]
                                    
                                print(aut_line)


#                        WebDriverWait(browser, 15).until(EC.element_to_be_clickable((By.XPATH, '//a[@title="Área"]/span'))).click()
                    
#                else:
#                        browser.back()
#                    print('sem autorizados')
#                        time.sleep(2)
#                except:
#                    browser.back()
#                    print('passou mais um')
#                    time.sleep(2)
#                time.sleep(4)
                try:
                    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.XPATH, '//a[@title="Área"]/span'))).click()
                except TimeoutException:
                    WebDriverWait(browser, 15).until(EC.element_to_be_clickable((By.XPATH, '(//span[@class="ms-Nav-linkText linkText_2f72ef6c"])[2]'))).click()
                    pass
                time.sleep(1)
                print(group,area,nucleo)
                nucleos = refresh_nucleos(browser) 
                cntn+=1
            WebDriverWait(browser, 15).until(EC.element_to_be_clickable((By.XPATH, '//a[@title="Direção"]/span'))).click()
#            time.sleep(3)
            areas=refresh_areas(browser)
            cnta+=1
    
#        time.sleep(2)        
        WebDriverWait(browser, 15).until(EC.element_to_be_clickable((By.XPATH, '//a[@title="Portal Home Page"]/span'))).click()
        groups=refresh_groups(browser)
        cntg+=1





    WebDriverWait(browser, 15).until(EC.element_to_be_clickable((By.XPATH, '//button[@id="O365_MainLink_Settings"]')))
    lout_btn=WebDriverWait(browser, 15).until(EC.element_to_be_clickable((By.XPATH, '//span[contains(text(),"FERNANDO MANUEL VELHA FRAGOSO FREITAS SIMOES")]')))
    lout_btn.click()
    sair_btn=WebDriverWait(browser, 15).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mectrl_body_signOut"]')))
    sair_btn.click()
    time.sleep(10)
    browser.quit()
    logging.info('LOGOUT SUCCESSFUL')
    print('5 stars')
    print(conta)
   
#    writer = pd.ExcelWriter("Autorizados.xlsx", engine='openpyxl', mode="a", if_sheet_exists="replace")
    df.to_excel("Autorizados.xlsx", sheet_name = "Autorizados", index=False)
#    writer.save() 

if __name__ == '__main__':
    main()




"""
(//div[contains(@class,"od-FieldRenderer-text")])[1]    name
//div[@class="sp-field-customFormatter"]/a    email
//div[contains(@class,"field_875b1af5")]    username
"""