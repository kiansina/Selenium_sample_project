from selenium import webdriver
from selenium.webdriver.common.by import By
import xml.etree.ElementTree as ET
import pandas as pd
import sinfun as sf
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time as t
def mil(amb):
    path=r'C:\Users\s.kian\OneDrive - Intrum Law\Desktop\MIL'
    path1=r'C:\Users\s.kian\OneDrive - Intrum Law\Desktop\MIL\results'
    #driver = webdriver.Chrome()
    #this silents the error
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument("--window-size=1920,1080")
    driver = webdriver.Chrome(options=options)
    driver.get("https://areautenti.masscreditcollection.eu/auth/login")
    com=driver.find_element(By.ID, "j_username")
    com.send_keys('sina.kian@i-law.it')
    com=driver.find_element(By.ID, "j_password")
    com.send_keys('')
    l=driver.find_elements(By.TAG_NAME, 'a')[2]
    l.click()
    while True:
        try:
            l=driver.find_elements(By.TAG_NAME, 'a')[2]
            l.send_keys(amb)
            dropdown_options = driver.find_elements(By.CSS_SELECTOR, 'ul.chosen-results li.active-result')
            desired_option_text = amb
            for option in dropdown_options:
                if option.text == desired_option_text:
                    option.click()
                    break
            break
        except:
            t.sleep(2)
            print(1)
    #
    l=driver.find_elements(By.ID, 'commessa_button')[0]
    l.click()


def tipology(amb,file):
    path=r'C:\Users\s.kian\OneDrive - Intrum Law\Desktop\MIL'
    path1=r'C:\Users\s.kian\OneDrive - Intrum Law\Desktop\MIL\results'
    df=pd.read_excel(path+'\\'+file+'.xlsx')
    #driver = webdriver.Chrome()
    #this silents the error
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument("--window-size=1500,1080")
    driver = webdriver.Chrome(options=options)
    driver.get("https://areautenti.masscreditcollection.eu/auth/login")
    com=driver.find_element(By.ID, "j_username")
    com.send_keys('sina.kian@i-law.it')
    com=driver.find_element(By.ID, "j_password")
    com.send_keys('')
    l=driver.find_elements(By.TAG_NAME, 'a')[2]
    l.click()
    driver.execute_script("window.scrollBy(0, -document.body.scrollHeight);")
    while True:
        try:
            l=driver.find_elements(By.TAG_NAME, 'a')[2]
            l.send_keys(amb.split()[0])
            dropdown_options = driver.find_elements(By.CSS_SELECTOR, 'ul.chosen-results li.active-result')
            desired_option_text = amb
            for option in dropdown_options:
                if option.text == desired_option_text:
                    option.click()
                    break
            break
        except:
            t.sleep(2)
            print(1)
    #
    l=driver.find_elements(By.ID, 'commessa_button')[0]
    l.click()
    ########################à
    while True:
        try:
            l=driver.find_elements(By.TAG_NAME, 'a')[63]
            ActionChains(driver).move_to_element(l).perform()
            l=driver.find_elements(By.TAG_NAME, 'a')[78]
            ActionChains(driver).move_to_element(l).perform()
            l=driver.find_elements(By.TAG_NAME, 'a')[81]
            l.click()
            break
        except:
            t.sleep(1)
            print(1)
    #
    ############################################
    ##############################################
    #############################################
    ##############################################
    driver.refresh()
    c=0
    for i in df.columns[2:]:
        print(i,c)
        c+=1
        try:
            l=driver.find_elements(By.TAG_NAME, 'a')[140]
            l.click()
            l=driver.find_elements(By.ID, 'nome')[1]
            l.send_keys(i.strip())
            l=driver.find_elements(By.ID, 'listaIntervalloTipoTipologiaSelect_chosen')[1]
            l.click()
            t.sleep(1)
            dropdown_options = driver.find_elements(By.CSS_SELECTOR, 'ul.chosen-results li.active-result')
            desired_option_text = 'Fisso'
            for option in dropdown_options:
                if option.text == desired_option_text:
                    option.click()
                    break
            #
            t.sleep(1)
            l=driver.find_elements(By.CLASS_NAME, 'btn')[-1]
            l.click()
            t.sleep(2)
            l=driver.find_elements(By.ID, 'button_callback')[0]
            l.click()
            driver.refresh()
            t.sleep(2)
        except Exception as e:
            # Handle the exception or error
            print(f"ERROR:{e}")
            # You can choose to log the error, take alternative actions, or simply pass
            # Continue to the next iteration of the loop
            lis.append((i,c))
            driver.get("https://areautenti.masscreditcollection.eu/secure/intervalloTipo/")
            t.sleep(3)
        continue





def Clusters(amb,file):
    path=r'C:\Users\s.kian\OneDrive - Intrum Law\Desktop\MIL'
    path1=r'C:\Users\s.kian\OneDrive - Intrum Law\Desktop\MIL\results'
    df=pd.read_excel(path+'\\'+file+'.xlsx')
    #driver = webdriver.Chrome()
    #this silents the error
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument("--window-size=1500,1080")
    driver = webdriver.Chrome(options=options)
    driver.get("https://areautenti.masscreditcollection.eu/auth/login")
    com=driver.find_element(By.ID, "j_username")
    com.send_keys('sina.kian@i-law.it')
    com=driver.find_element(By.ID, "j_password")
    com.send_keys('')
    l=driver.find_elements(By.TAG_NAME, 'a')[2]
    l.click()
    driver.execute_script("window.scrollBy(0, -document.body.scrollHeight);")
    while True:
        try:
            l=driver.find_elements(By.TAG_NAME, 'a')[2]
            l.send_keys(amb.split()[0])
            dropdown_options = driver.find_elements(By.CSS_SELECTOR, 'ul.chosen-results li.active-result')
            desired_option_text = amb
            for option in dropdown_options:
                if option.text == desired_option_text:
                    option.click()
                    break
            break
        except:
            t.sleep(2)
            print(1)
    #
    ##########fase 3
    l=driver.find_elements(By.ID, 'commessa_button')[0]
    l.click()
    ########################à
    while True:
        try:
            l=driver.find_elements(By.TAG_NAME, 'a')[63]
            ActionChains(driver).move_to_element(l).perform()
            l=driver.find_elements(By.TAG_NAME, 'a')[78]
            ActionChains(driver).move_to_element(l).perform()
            l=driver.find_elements(By.TAG_NAME, 'a')[80]
            l.click()
            break
        except:
            t.sleep(1)
            print(1)
    ############################################
    ##############################################
    #############################################
    ##############################################
    driver.refresh()
    c=1
    cc=1
    t.sleep(2)
    for i in df.columns[2:]:
        if c%3==0 or cc%10==0:
            t.sleep(60)
        print(i,c)
        c+=1
        for j in df.index:
            try:
                l=driver.find_elements(By.TAG_NAME, 'a')[140]
                l.click()
                l=driver.find_elements(By.CLASS_NAME, 'form-group')[9]
                t.sleep(3)
                l.click()
                wait = WebDriverWait(driver, 10)
                dropdown_options = wait.until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR, 'ul.chosen-results li.active-result')))
                desired_option_text = i.strip()
                for option in dropdown_options:
                    if option.text == desired_option_text:
                        option.click()
                        break
                #
                l=driver.find_elements(By.ID, 'intervalloDa')[1]
                l.send_keys(str(df['DA'].loc[j]))
                l=driver.find_elements(By.ID, 'intervalloA')[1]
                l.send_keys(str(df['A'].loc[j]))
                l=driver.find_elements(By.ID, 'valore')[1]
                l.send_keys(str(df[i].loc[j]))
                l=driver.find_elements(By.TAG_NAME, 'button')[5]
                l.click()
                t.sleep(2)
                l=driver.find_elements(By.ID, 'button_callback')[0]
                t.sleep(2)
                l.click()
                driver.refresh()
                t.sleep(2)
                cc+=1
                print(df['DA'].loc[j], df['A'].loc[j], df[i].loc[j])
            except Exception as e:
                # Handle the exception or error
                print(f"ERROR:{e}")
                # You can choose to log the error, take alternative actions, or simply pass
                # Continue to the next iteration of the loop
                lis.append((i,c,df['DA'].loc[j], df['A'].loc[j], df[i].loc[j]))
                driver.get("https://areautenti.masscreditcollection.eu/secure/intervallo/")
                t.sleep(10)
            continue


def Dettaglio(amb,file,excel=True):
    path=r'C:\Users\s.kian\OneDrive - Intrum Law\Desktop\MIL'
    path1=r'C:\Users\s.kian\OneDrive - Intrum Law\Desktop\MIL\results'
    df=pd.read_excel(path+'\\'+file+'.xlsx')
    #driver = webdriver.Chrome()
    #this silents the error
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument("--window-size=1500,1080")
    driver = webdriver.Chrome(options=options)
    #
    driver.get("https://areautenti.masscreditcollection.eu/auth/login")
    com=driver.find_element(By.ID, "j_username")
    com.send_keys('sina.kian@i-law.it')
    com=driver.find_element(By.ID, "j_password")
    com.send_keys('')
    l=driver.find_elements(By.TAG_NAME, 'a')[2]
    l.click()
    driver.execute_script("window.scrollBy(0, -document.body.scrollHeight);")
    while True:
        try:
            l=driver.find_elements(By.TAG_NAME, 'a')[2]
            l.send_keys(amb.split()[0])
            dropdown_options = driver.find_elements(By.CSS_SELECTOR, 'ul.chosen-results li.active-result')
            desired_option_text = amb
            for option in dropdown_options:
                if option.text == desired_option_text:
                    option.click()
                    break
            break
        except:
            t.sleep(2)
            print(1)
    #
    l=driver.find_elements(By.ID, 'commessa_button')[0]
    l.click()
    ########################à
    while True:
        try:
            l=driver.find_elements(By.TAG_NAME, 'a')[63]
            ActionChains(driver).move_to_element(l).perform()
            l=driver.find_elements(By.TAG_NAME, 'a')[82]
            ActionChains(driver).move_to_element(l).perform()
            l=driver.find_elements(By.TAG_NAME, 'a')[85]
            l.click()
            break
        except:
            t.sleep(1)
            print(1)
    #
    ############################################
    ##############################################
    #############################################
    ##############################################
    driver.refresh()
    t.sleep(2)
    c=0
    if excel==False:
        looplist=list(map(lambda x: x[0],lis))
    else:
        looplist=df.columns[2:]
    for i in looplist:
        print(i,c)
        c+=1
        try:
            l=driver.find_elements(By.TAG_NAME, 'a')[141]
            l.click()
            t.sleep(1)
            l=driver.find_elements(By.ID, 'nome')[1]
            l.send_keys(i.strip())
            l=driver.find_elements(By.ID, 'valoreDefault')[1]
            l.send_keys(0)
            l=driver.find_elements(By.ID, 'listaCostiTipo_chosen')[1]
            l.click()
            t.sleep(1)
            dropdown_options = driver.find_elements(By.CSS_SELECTOR, 'ul.chosen-results li.active-result')
            desired_option_text = 'Onorario'
            for option in dropdown_options:
                if option.text == desired_option_text:
                    option.click()
                    break
            #
            l=driver.find_elements(By.ID, 'listaTipologiaCalcoloSelect_chosen')[1]
            l.click()
            t.sleep(1)
            dropdown_options = driver.find_elements(By.CSS_SELECTOR, 'ul.chosen-results li.active-result')
            desired_option_text = 'Variabile'
            for option in dropdown_options:
                if option.text == desired_option_text:
                    option.click()
                    break
            #
            l=driver.find_elements(By.ID, 'listaIntervalloTipoSelect_chosen')[0]
            l.click()
            t.sleep(1)
            dropdown_options = driver.find_elements(By.CSS_SELECTOR, 'ul.chosen-results li.active-result')
            desired_option_text = i.strip()
            for option in dropdown_options:
                if option.text == desired_option_text:
                    option.click()
                    break
            #
            l=driver.find_elements(By.ID, 'listaValoreRiferimentoConfrontoSelect_chosen')[1]
            l.click()
            t.sleep(1)
            dropdown_options = driver.find_elements(By.CSS_SELECTOR, 'ul.chosen-results li.active-result')
            desired_option_text = 'Capitale residuo'
            for option in dropdown_options:
                if option.text == desired_option_text:
                    option.click()
                    break
            #
            l=driver.find_elements(By.CLASS_NAME, 'btn')[-1]
            l.click()
            t.sleep(2)
            l=driver.find_elements(By.ID, 'button_callback')[0]
            l.click()
            driver.refresh()
            t.sleep(2)
            driver.execute_script("window.scrollBy(0, -document.body.scrollHeight);")
        except Exception as e:
            # Handle the exception or error
            print(f"ERROR:{e}")
            # You can choose to log the error, take alternative actions, or simply pass
            # Continue to the next iteration of the loop
            if excel==True:
                lis.append((i,c))
            else:
                lis2.append((i,c))
            driver.get("https://areautenti.masscreditcollection.eu/secure/costoDescrizione/")
            t.sleep(10)
            driver.execute_script("window.scrollBy(0, -document.body.scrollHeight);")
        continue





def Eventi(amb,file):
    path=r'C:\Users\s.kian\OneDrive - Intrum Law\Desktop\MIL'
    path1=r'C:\Users\s.kian\OneDrive - Intrum Law\Desktop\MIL\results'
    df=pd.read_excel(path+'\\'+file+'.xlsx')
    #driver = webdriver.Chrome()
    #this silents the error
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument("--window-size=1500,1080")
    driver = webdriver.Chrome(options=options)
    driver.get("https://areautenti.masscreditcollection.eu/auth/login")
    com=driver.find_element(By.ID, "j_username")
    com.send_keys('sina.kian@i-law.it')
    com=driver.find_element(By.ID, "j_password")
    com.send_keys('')
    l=driver.find_elements(By.TAG_NAME, 'a')[2]
    l.click()
    driver.execute_script("window.scrollBy(0, -document.body.scrollHeight);")
    while True:
        try:
            l=driver.find_elements(By.TAG_NAME, 'a')[2]
            l.send_keys(amb.split()[0])
            dropdown_options = driver.find_elements(By.CSS_SELECTOR, 'ul.chosen-results li.active-result')
            desired_option_text = amb
            for option in dropdown_options:
                if option.text == desired_option_text:
                    option.click()
                    break
            break
        except:
            t.sleep(2)
            print(1)
    #
    l=driver.find_elements(By.ID, 'commessa_button')[0]
    l.click()
    ########################à
    while True:
        try:
            l=driver.find_elements(By.TAG_NAME, 'a')[63]
            ActionChains(driver).move_to_element(l).perform()
            l=driver.find_elements(By.TAG_NAME, 'a')[119]
            ActionChains(driver).move_to_element(l).perform()
            l=driver.find_elements(By.TAG_NAME, 'a')[120]
            l.click()
            break
        except:
            t.sleep(1)
            print(1)
    #
    while True:
        try:
            l=driver.find_elements(By.TAG_NAME, 'a')[143]
            l.click()
            t.sleep(1)
            break
        except:
            t.sleep(1)
            print(1)
    #
    ############################################
    ##############################################
    #############################################
    ##############################################
    driver.refresh()
    c=0
    for i in df.columns[2:]:
        c+=1
        print(i,c)
        try:
            l=driver.find_elements(By.TAG_NAME, 'a')[140]
            l.click()
            t.sleep(1)
            l=driver.find_elements(By.ID, 'nome')[1]
            l.send_keys(i.strip())
            l=driver.find_elements(By.ID, 'etichetta')[1]
            l.send_keys('fatto')
            l=driver.find_elements(By.NAME, 'scadenzaGiorni')[1]
            l.send_keys('10')
            l=driver.find_elements(By.ID, 'eventoPrioritaSelect_chosen')[1]
            l.click()
            t.sleep(1)
            dropdown_options = driver.find_elements(By.CSS_SELECTOR, 'ul.chosen-results li.active-result')
            desired_option_text = 'Media'
            for option in dropdown_options:
                if option.text == desired_option_text:
                    option.click()
                    break
            #
            l=driver.find_elements(By.ID, 'descrizione')[1]
            l.send_keys(i.strip())
            #?????????????????????????????????
            l = driver.find_elements(By.CSS_SELECTOR, '.label-true.btn.btn-primary.pulsanteLarghezza50.pulsanteSi')[1]
            l.click()
            #??????????????????????????????
            driver.execute_script("window.scrollBy(0, 500);")
            #driver.execute_script("document.body.style.zoom = '100%';")
            l=driver.find_elements(By.ID, 'azioneAutomatica_chosen')[1]
            l.click()
            t.sleep(1)
            dropdown_options = driver.find_elements(By.CSS_SELECTOR, 'ul.chosen-results li.active-result')
            desired_option_text = 'Aggiungi costo'
            for option in dropdown_options:
                if option.text == desired_option_text:
                    option.click()
                    break
            #
            t.sleep(2)
            l=driver.find_elements(By.ID, 'selezionaTipoCosto_chosen')[1]
            l.click()
            t.sleep(1)
            dropdown_options = driver.find_elements(By.CSS_SELECTOR, 'ul.chosen-results li.active-result')
            desired_option_text = 'Onorario'
            for option in dropdown_options:
                if option.text == desired_option_text:
                    option.click()
                    break
            #
            t.sleep(2)
            l=driver.find_elements(By.ID, 'selezionaCosto_2_chosen')[0]
            l.click()
            t.sleep(1)
            dropdown_options = driver.find_elements(By.CSS_SELECTOR, 'ul.chosen-results li.active-result')
            desired_option_text = i.strip()
            for option in dropdown_options:
                if option.text == desired_option_text:
                    option.click()
                    break
            #
            l=driver.find_elements(By.TAG_NAME, 'a')[-1]
            l.click()
            t.sleep(3)
            l=driver.find_elements(By.ID, 'button_callback')[0]
            l.click()
            driver.refresh()
            driver.execute_script("window.scrollBy(0, -document.body.scrollHeight);")
            t.sleep(2)
        except Exception as e:
            # Handle the exception or error
            print(f"ERROR:{e}")
            # You can choose to log the error, take alternative actions, or simply pass
            # Continue to the next iteration of the loop
            lis.append((i,c))
            driver.get("https://areautenti.masscreditcollection.eu/secure/evento")
            t.sleep(10)
            driver.execute_script("window.scrollBy(0, -document.body.scrollHeight);")
        continue
    print('FIRST ROUND FINISHED')
    for i in lis:
        i=i[0]
        c+=1
        print(i,c)
        try:
            l=driver.find_elements(By.TAG_NAME, 'a')[140]
            l.click()
            t.sleep(1)
            l=driver.find_elements(By.ID, 'nome')[1]
            l.send_keys(i.strip())
            l=driver.find_elements(By.ID, 'etichetta')[1]
            l.send_keys('fatto')
            l=driver.find_elements(By.NAME, 'scadenzaGiorni')[1]
            l.send_keys('10')
            l=driver.find_elements(By.ID, 'eventoPrioritaSelect_chosen')[1]
            l.click()
            t.sleep(1)
            dropdown_options = driver.find_elements(By.CSS_SELECTOR, 'ul.chosen-results li.active-result')
            desired_option_text = 'Media'
            for option in dropdown_options:
                if option.text == desired_option_text:
                    option.click()
                    break
            #
            l=driver.find_elements(By.ID, 'descrizione')[1]
            l.send_keys(i.strip())
            #?????????????????????????????????
            l = driver.find_elements(By.CSS_SELECTOR, '.label-true.btn.btn-primary.pulsanteLarghezza50.pulsanteSi')[1]
            l.click()
            #??????????????????????????????
            driver.execute_script("window.scrollBy(0, 500);")
            #driver.execute_script("document.body.style.zoom = '100%';")
            l=driver.find_elements(By.ID, 'azioneAutomatica_chosen')[1]
            l.click()
            t.sleep(1)
            dropdown_options = driver.find_elements(By.CSS_SELECTOR, 'ul.chosen-results li.active-result')
            desired_option_text = 'Aggiungi costo'
            for option in dropdown_options:
                if option.text == desired_option_text:
                    option.click()
                    break
            #
            t.sleep(2)
            l=driver.find_elements(By.ID, 'selezionaTipoCosto_chosen')[1]
            l.click()
            t.sleep(1)
            dropdown_options = driver.find_elements(By.CSS_SELECTOR, 'ul.chosen-results li.active-result')
            desired_option_text = 'Onorario'
            for option in dropdown_options:
                if option.text == desired_option_text:
                    option.click()
                    break
            #
            t.sleep(2)
            l=driver.find_elements(By.ID, 'selezionaCosto_2_chosen')[0]
            l.click()
            t.sleep(1)
            dropdown_options = driver.find_elements(By.CSS_SELECTOR, 'ul.chosen-results li.active-result')
            desired_option_text = i.strip()
            for option in dropdown_options:
                if option.text == desired_option_text:
                    option.click()
                    break
            #
            l=driver.find_elements(By.TAG_NAME, 'a')[-1]
            l.click()
            t.sleep(3)
            l=driver.find_elements(By.ID, 'button_callback')[0]
            l.click()
            driver.refresh()
            driver.execute_script("window.scrollBy(0, -document.body.scrollHeight);")
            t.sleep(2)
        except Exception as e:
            # Handle the exception or error
            print(f"ERROR:{e}")
            # You can choose to log the error, take alternative actions, or simply pass
            # Continue to the next iteration of the loop
            lis.append((i,c))
            driver.get("https://areautenti.masscreditcollection.eu/secure/evento")
            t.sleep(10)
            driver.execute_script("window.scrollBy(0, -document.body.scrollHeight);")
        continue
