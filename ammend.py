from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC 
from selenium.common.exceptions import NoSuchElementException 
from selenium.webdriver.support.ui import Select 
from selenium.webdriver.common.alert import Alert
import pandas as pd 
import time 
import json 
import math as m 
from colorama import Fore, Style
from datetime import datetime
from dynamic import IdentifyPort

def Ammend_Fields(file_path,username,password):
    TIMEOUT = 30 
    filename = file_path; username=username; password=password 

    print(f"[INFO] Reading Data From {filename}")

    dtype_dict = {
    'numerical_column': float,
    'wildcard_column': str
    }
    CTN_TYPE = {
        "10X":"24x16x10-RS10",
        "12X":"24x16x12-RS12X",
        'EURO':"25x16x11-Euro",
        '6X':"24X16X6-RS6",
        '5X':"24x16x5-RS5X"
    }
    data = pd.read_excel(filename,dtype=dtype_dict)
    # print(data.columns)
    final_df = pd.DataFrame()
    final_df_failed = pd.DataFrame()

    driver = webdriver.Chrome()

    driver.get("https://auth.damco.com/adfs/ls/?wtrealm=https%3A%2F%2Fportal.damco.com&wctx=WsFedOwinState%3DclUNwkNUP4dpiVAR1vBCtUw6PO0n1IdLnSzMvQIIahHiM5jw7bnM1i2W9bgWqEHiqJv0vURH11fUgH0T2bI4Ldn9clIwiH17_zsNz9TN0eh1xHQLEkPZ7RW9GtRMhvKeVb78R-6cSvuxKdlD7UzvO_hdhYv_dP-c15p-EL_jLfk&wa=wsignin1.0")
    print("got browser")
    time.sleep(3)

    login_username = driver.find_element(value='ctl00_ContentPlaceHolder1_UsernameTextBox')
    login_username.send_keys(username)
    login_password = driver.find_element(value='ctl00_ContentPlaceHolder1_PasswordTextBox')
    login_password.send_keys(password)
    login_enter = driver.find_element(value = 'ctl00_ContentPlaceHolder1_SubmitButton')
    login_enter.send_keys(Keys.ENTER)



    try:

        element = WebDriverWait(driver,5).until(
            EC.presence_of_element_located((By.ID,'ctl00_ContentPlaceHolder1_ErrorTextLabel')))
        if 'Authentication Failed'.lower() in element.text.lower():
            return False 
    except:
        pass
    print("got login")

    for index,df in data.iterrows():
        try:
            
            PO = df['PO#'].split("-")[0]
            booking_id = int(df['Booking id'])
            L_NO = int(df['PO#'].split("-")[1])
            date = f"{df['Plan-HOD'].year}-{df['Plan-HOD'].month}-{df['Plan-HOD'].day}"
            parsed_date = datetime.strptime(date,'%Y-%m-%d')
            date = parsed_date.strftime("%Y-%m-%d")
            country = df['Country']
            quantity = str(df['Order Qty']) #quantity
            port_of_discharge = IdentifyPort(country) #port of discharge
            packages = str(m.floor(df['CARTON QTY'])) #packages and carton count
            weight = round(float(df['GROSS WT']),1) #gross weight 
            carton_cbm = df['CARTON CBM'] #measurement 
            ctn_type = str(df['CTN Type']) #ctn type
            
            print("[Extracted Info] \n ")
            print(f"<PO> : {PO}")
            print(f"<Line No> : {L_NO}")
            print(f"<date> : {date}")
            search = f'ViewSOAction.action?so_number={booking_id}&amp;searchByShipper_id'
            driver.get("https://booking.damco.com/ShipperPortalWeb/SearchAction.action")

            WebDriverWait(driver,TIMEOUT).until(EC.presence_of_element_located((By.ID,"tab_treetab2")))

            booking = driver.find_element(by = By.ID,value='tab_treetab2')
            booking.click()

            WebDriverWait(driver,TIMEOUT).until(EC.presence_of_element_located((By.ID,'searchSO_SO_NO')))
            po_number = driver.find_element(value='searchSO_SO_NO')
            po_number.send_keys(booking_id)
            #print("got search")
            
            search_po_number = driver.find_element(value='searchSObtn')
            search_po_number.send_keys(Keys.ENTER)
            
            time.sleep(2)
            a_tag = driver.find_element(by=By.XPATH, value= f'//a[text()="{booking_id}"]')
            a_tag.click()

            WebDriverWait(driver,TIMEOUT).until(EC.presence_of_element_located((By.ID,'soEdit')))
            so_edit = driver.find_element(by=By.ID, value= 'soEdit')
            so_edit.click()

            # <------------------------ Header Tab ------------------------>
            IDS = {
                "estmDlvrDtId":str(date), #date
                "portOfDischargeGrpId":str(port_of_discharge) #port of discharge
                }
            for key in IDS:
                WebDriverWait(driver,TIMEOUT).until(EC.presence_of_element_located((By.ID,key)))
                so_edit = driver.find_element(by=By.ID, value= key)
                value = so_edit.get_attribute('value')
                if str(value) != IDS.get(key):
                    for _ in range(len(value)):
                        so_edit.send_keys(Keys.BACK_SPACE)
                    so_edit.send_keys(IDS.get(key))

            WebDriverWait(driver,TIMEOUT).until(EC.presence_of_element_located((By.ID,'tab_treetab3')))
            so_edit = driver.find_element(by=By.ID, value= 'tab_treetab3')
            so_edit.click()


            # <------------------------ Details Tab ------------------------>

            time.sleep(6)
            
            # --> Quantity
            WebDriverWait(driver,TIMEOUT).until(EC.presence_of_element_located((By.ID,'bookedQtyId0')))
            so_edit = driver.find_element(by=By.ID, value= f'bookedQtyId0')
            value = so_edit.get_attribute('value')
            if str(value) != str(quantity):
                for _ in range(len(value)):
                    so_edit.send_keys(Keys.BACK_SPACE)
                so_edit.send_keys(str(quantity))



            # --> Packages
            WebDriverWait(driver,TIMEOUT).until(EC.presence_of_element_located((By.ID,'bookedPackagesId0')))
            so_edit = driver.find_element(by=By.ID, value= f'bookedPackagesId0')
            value = str(so_edit.get_attribute('value'))
            if value != packages:
                for _ in range(len(value)):
                    so_edit.send_keys(Keys.BACK_SPACE)
                so_edit.send_keys(packages)
            
            # --> weight
            WebDriverWait(driver,TIMEOUT).until(EC.presence_of_element_located((By.ID,'bookedWeightId0')))
            so_edit = driver.find_element(by=By.ID, value= f'bookedWeightId0')
            value = so_edit.get_attribute('value')
            value_ = float(value)
            value_ = str(round(value_,1))
            if value_ != weight:
                for _ in range(len(value)):
                    so_edit.send_keys(Keys.BACK_SPACE)
                so_edit.send_keys(str(weight))


            # --> Measurement / Carton CBM
            WebDriverWait(driver,TIMEOUT).until(EC.presence_of_element_located((By.ID,'bookedMeasurementId0')))
            so_edit = driver.find_element(by=By.ID, value= f'bookedMeasurementId0')
            value = so_edit.get_attribute('value')
            value_,carton_cbm = float(value), float(carton_cbm)
            value_,carton_cbm = str(round(value_,2)), str(round(carton_cbm,2))
            if value_ != carton_cbm:
                for _ in range(len(str(value))):
                    so_edit.send_keys(Keys.BACK_SPACE)
                so_edit.send_keys(str(carton_cbm))

            # --> packages
            WebDriverWait(driver,TIMEOUT).until(EC.presence_of_element_located((By.NAME,'soDto.soLineDtoList[0].soLineRefDtoList[4].refValue')))
            so_edit = driver.find_element(by=By.NAME,value="soDto.soLineDtoList[0].soLineRefDtoList[4].refValue")
            value = so_edit.get_attribute('value')
            if str(value) != str(packages):
                for _ in range(len(value)):
                    so_edit.send_keys(Keys.BACK_SPACE)
                so_edit.send_keys(str(packages))

            types = list(CTN_TYPE.keys())
            if str(df['CTN Type']) in types: 
                dropdown = Select(driver.find_element(by=By.ID,value=f"dynafield_0_0_refValue"))
                dropdown.select_by_visible_text(CTN_TYPE.get(str(df['CTN Type'])))  #no column in excel
                #print("got field1")
                
                dropdown = Select(driver.find_element(by=By.ID,value=f"dynafield_0_1_refValue"))
                dropdown.select_by_visible_text(CTN_TYPE.get(str(df['CTN Type'])))  #no column in excel
                # print("got field2")
                # print(CTN_TYPE.get(str(df['CTN Type'])))
            else:
                raise ValueError

            # --> Finishing the Booking 
            #input("Continue...")
            if True:
                #WebDriverWait(driver,TIMEOUT).until(EC.presence_of_element_located((By.ID,"tab_treetab1")))
                #print("got headers")
                WebDriverWait(driver,TIMEOUT).until(EC.presence_of_element_located((By.ID,'SaveAsMenuBtnId')))
                save_button = driver.find_element(by=By.ID,value="SaveAsMenuBtnId")
                save_button.click()
                #print("got save button")

                WebDriverWait(driver,TIMEOUT).until(EC.presence_of_element_located((By.ID,'SOStatusOption')))
                span_tag = driver.find_element(by=By.ID, value = 'SOStatusOption')
                draft_and_finished = span_tag.find_elements(by=By.TAG_NAME, value = 'a')
                # WebDriverWait(driver,TIMEOUT).until(EC.presence_of_element_located((By.XPATH,"//a[contains(@onclick, 'saveAsBookingOption('draft')')]")))
                # button = driver.find_element(by=By.XPATH,value="//a[contains(@onclick, 'saveAsBookingOption('draft')')]")
                
                #input("Enter to continue....!")
                draft_and_finished[1].click()
                #print(draft_and_finished[1].get_attribute("onclick"))
                time.sleep(3)
                try:
                    pop_up = Alert(driver)
                    pop_up.accept()
                    #print("got popup")
                except Exception as e:
                    print(Fore.CYAN+"Alert not found, continuing forward"+Style.RESET_ALL)
                WebDriverWait(driver, TIMEOUT).until(EC.invisibility_of_element_located((By.ID, 'progressStatusId')))
                WebDriverWait(driver,3).until(EC.text_to_be_present_in_element((By.ID,'MsgDivId'),f'{booking_id} saved successfully'))
            time.sleep(2)
            #time.sleep(2)
            #
            # WebDriverWait(driver,TIMEOUT).until(EC.presence_of_element_located((By.ID,'MsgDivId')))
            
            print(Fore.GREEN+'Booking Fininsed....!!!'+Style.RESET_ALL)
            print("<Booking ID> : ",Fore.GREEN+str(booking_id),Style.RESET_ALL)
            df['booking_status'] = "success" 
            
            final_df = pd.concat([final_df, df], axis=1)
            print(Fore.GREEN+"->"*3,Fore.GREEN+"Success"+Style.RESET_ALL)
            final_df.T.to_excel(f"AmmendSuccess.xlsx",index=False)
            print(Fore.MAGENTA+"->"*3,Fore.MAGENTA+"-"*10,Style.RESET_ALL)

        
        except NoSuchElementException:
            print("Record not found.")
            df['booking_status'] = "failed" 
            df['booking_id'] = '-'
            print(Fore.RED+"->"*3,Fore.RED+"Failed"+Style.RESET_ALL)
            final_df_failed = pd.concat([final_df_failed, df], axis=1)
            final_df_failed.T.to_excel("AmmendFailed.xlsx",index=False)
            print("-"*10)
            continue
        except Exception as e:
            print(str(e))
            df['booking_status'] = "failed" 
            df['booking_id'] = '-'
            print(Fore.RED+"->"*3,Fore.RED+"Failed"+Style.RESET_ALL)
            final_df_failed = pd.concat([final_df_failed, df], axis=1)
            final_df_failed.T.to_excel("AmmendFailed.xlsx",index=False)
            print("-"*10)
            continue
            