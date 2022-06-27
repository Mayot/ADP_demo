
# Objectif : Se connecter à ADP pour ajouter les jours en TT depuis un fichier
# Réalisation des imports


from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains #pour créer des enchainements d'action, notamment mouseover
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select #pour gérer les différents tabs
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from unidecode import unidecode #for unicode conversion
from datetime import datetime, timedelta #For date
import time #for sleep
import re #for Regex
import logging #for logging
import os #for listing file
#from mail import send_mail_with_attm, send_mail #lib perso pour envoi de mail




# Variables initiales


app_name = "ADP"
workspace = "C:/Apps/"+app_name+"/"
download_folder = "C:\\Apps\\"+app_name


# Variables à modifier
url = os.getenv('url', default=None)
username = os.getenv('ADP_user', default=None)
pwd = os.getenv('ADP_pwd', default=None)





# Début du programme 


#filemode : a = append, w = overwrite every run
logging.basicConfig(filename=workspace+app_name+'_log.log', filemode='a', format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
#Let us Create an object 
logger=logging.getLogger() 
#Now we are going to Set the threshold of logger to INFO 
logger.setLevel(logging.INFO) #DEBUG, INFO, WARNING, ERROR, CRITICAL

logger.info("START START START START START START START START START")
#logger.info(datetime.now().strftime("%m/%d/%Y, %H:%M:%S"))

#Read date.txt file (contains all dates to add in ADP) and get all the line
with open(workspace+'date.txt') as f:
    #new collection for future
    myFile = []
    for readline in f:
        #Strip (delete \r\n) and replace to harmonize format
        line = readline.strip().replace("-","/")
        #Ignore 1st line
        if line == 'dd/MM/yyyy':
            pass
        else:
            #Append collection
            myFile.append(line)
logger.info('Fichier contenant '+str(len(myFile))+' lignes.')


#si pas de donnée on fait rien
if len(myFile) > 0:
    ### Démarage du programme
    try:
        #Set drivers (Browser = Chrome)
        #to remove irrelevant log error
        options = webdriver.ChromeOptions()
        
        preferences = {
            "download.default_directory" : download_folder
            }
        options.add_experimental_option("prefs", preferences)
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        driver = webdriver.Chrome(options=options)

        #Ouverture de l'url via le driver
        driver.implicitly_wait(5)
        driver.get(url)

        ### Connexion, changement de tab, alimentation

        # Store the ID of the original window (tab)
        original_window = driver.current_window_handle
        #input user + pwd
        driver.find_elements_by_id("login")[0].send_keys(username)
        driver.find_elements_by_id("login-pw")[0].send_keys(pwd)
        #wait then log in
        driver.implicitly_wait(5)
        driver.find_element_by_xpath("/html/body/div/div/div/div[1]/div[3]/div[4]/div/form/div[4]/button").click()
        #wait then open Activité page in new tab
        time.sleep(2)
        driver.find_element_by_xpath('//*[@id="menu_2_navItem"]').click()
        driver.implicitly_wait(5)
        driver.find_element_by_xpath('//*[@id="menu_2_ttd_0_0"]/span[2]').click()
        time.sleep(2)
        #switch to old tab
        driver.switch_to.window(original_window)
        #closing the old tab
        driver.close()
        # Wait for the new window or tab
        driver.implicitly_wait(5)
        # Loop through until we find a new window handle
        for window_handle in driver.window_handles:
            if window_handle != original_window:
                driver.switch_to.window(window_handle)
        time.sleep(2)
        # loop for every date in collection then add to TT
        for date in myFile:
            #open declaration event
            driver.implicitly_wait(5)
            driver.find_element_by_xpath('//*[@id="btn_acc_add_label"]').click()
            time.sleep(1)
            #select event family (Télétravail)
            select_TT = Select(driver.find_element_by_id('evt_famille'))
            select_TT.select_by_visible_text('Télétravail')
            time.sleep(2)
            #select event nature (Télétravail)
            select_EVT = Select(driver.find_element_by_id('evt_natSelect'))
            select_EVT.select_by_visible_text('Télétravail')
            #go on datepicker
            datepicker = driver.find_elements_by_id("evt_datedeb")[0]
            #clear pre-registred value
            datepicker.clear()
            #send date in datepicker
            #Pour info on ne peut logger du TT que pour le mois en cours
            datepicker.send_keys(date)
            time.sleep(1)
            #Save button
            driver.find_element_by_id('evt_save_label').click()
            time.sleep(1)
            #Get box title value in variable to testing if error 
            titleBox = driver.find_element_by_id('dialogBoxGen_1_title')
            #close title box button
            driver.find_element_by_id('btn1_label').click()
            #test title value
            if titleBox.text == 'Erreur':
                logger.error('Problème avec la date: '+date)
                #close saisir evenement for the following date
                driver.find_element_by_id('evt_close_label').click()
            #if no error
            else:
                logger.info('Date ajoutée à ADP en Télétravail: '+date)


            time.sleep(1)
        #on quitte à la fin de l'execution
        driver.quit()

        #Reset data from date file for next run
        resetFile = open(workspace+'date.txt', "w")
        resetFile.write('dd-MM-yyyy\n')
        resetFile.close()
        logger.info("Fichier effacé")

    #handle exception
    except Exception as err:
        logger.error(str(err))
        try: 
            #Destinaire du mail
            to = os.getenv('email', default=None)
            subject = "Erreur ADP"
            msg = str(err)
            #send_mail(to,subject,msg)
        except:
            pass
logger.info("END END END END END END END END END")