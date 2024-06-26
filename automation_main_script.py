import os
import sys
from os import path 
from glob import glob
from time import sleep
import PySimpleGUI as sg
from datetime import date
import mimetypes
import smtplib, ssl
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.multipart import MIMEMultipart
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC


def send_email(email_from, email_to, subject, text, html, p_filename=None):
    try:
        msg = MIMEMultipart('multipart')
        msg['Subject'] = subject
        msg['From'] = email_from
        msg['To'] = email_to
        text = text
        html = html
        part1 = MIMEText(text, 'plain')
        part2 = MIMEText(html, 'html')
        msg.attach(part1)
        msg.attach(part2)

        if p_filename is not None:
            if not os.path.isfile(p_filename):
                return
            ctype, encoding = mimetypes.guess_type(p_filename)
            if ctype is None or encoding is not None:
                ctype = 'application/octet-stream'
            maintype, subtype = ctype.split('/', 1)
            if maintype == 'text':
                with open(p_filename) as f:
                    mime = MIMEText(f.read(), _subtype=subtype)
            elif maintype == 'image':
                with open(p_filename, 'rb') as f:
                    mime = MIMEImage(f.read(), _subtype=subtype)
            elif maintype == 'audio':
                with open(p_filename, 'rb')  as f:
                    mime = MIMEAudio(f.read(), _subtype=subtype)
            else:
                with open(p_filename, 'rb') as f:
                    mime = MIMEBase(maintype, subtype)
                mime.set_payload(f.read())
                encoders.encode_base64(mime)
                mime.add_header('Content-Disposition', 'attachment', filename=p_filename.split("\\")[-1])
            msg.attach(mime)

        context = ssl.create_default_context()
        with smtplib.SMTP("email-smtp.sa-east-1.amazonaws.com", 25) as server:
            server.starttls(context=context)
            server.ehlo()
            server.login(email_from, 'email_password')
            response = server.sendmail(email_from, email_to, msg.as_string())
            server.quit()
        if not response:
            return {"msg": "E-mail sent successfully"}
        else:
            return response
    except Exception as e:
        return e


def get_exe_directory():
    if getattr(sys, 'frozen', False):  # Running as a bundled executable
        return os.path.dirname(sys.executable)
    else:  # Running as a script
        return os.path.dirname(os.path.abspath(__file__))


def open_site():
    try:
        chrome_options = Options()
        chrome_options.add_experimental_option("detach", True)
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": get_exe_directory(),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })

        chrome_service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=chrome_service, options=chrome_options)
        driver.set_window_size(1400, 800)
        print('Preparing to access the site .... ')
        driver.get("https://banco.bradesco/html/pessoajuridica/index.shtm")
        print('Site Accessed')


        return driver

    except Exception as error:
        
        print(error)


def login(driver, username, password):

    print("starting login")

    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="topBar"]/div[1]/div[2]/a[1]')))
    
    sleep(3)

    account_access = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="topBar"]/div[1]/div[2]/a[1]')))
    
    account_access.click()

    print("Found account access button")

    sleep(3)

    try:

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="modalOtherBrowser"]/a[2]/img[1]')))

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="modalOtherBrowser"]/a[2]/img[1]')))
        
        driver.find_element('xpath', '//*[@id="modalOtherBrowser"]/a[2]/img[1]').click()

        print("Found account access popup button")

    except Exception as e:
        # driver.quit()
        print("Popup not found")

        try:
            WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="topBar"]/div[1]/div[2]/a[1]')))
        
            account_access = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="topBar"]/div[1]/div[2]/a[1]')))
            
            account_access.click()

            print("Found account access button on retry")
            
        except Exception as e:
            pass

    sleep(3)

    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="identificationForm:txtUsuario"]')))
    driver.find_element('xpath', '//*[@id="identificationForm:txtUsuario"]').send_keys(username)

    print("Filled in username")
    sleep(1)

    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="identificationForm:txtSenha"]')))
    driver.find_element('xpath', '//*[@id="identificationForm:txtSenha"]').send_keys(password)

    print("Filled in password")

    sleep(2)

    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="identificationForm:botaoAvancar"]')))
    
    driver.find_element('xpath', '//*[@id="identificationForm:botaoAvancar"]').click()

    sleep(4)

    try:

        driver.find_element('xpath', '/html/body/div[3]/form[1]/div[2]/div[3]/div/div[2]/div[2]/div[2]/span[1]/div')

        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="identificationForm:txtUsuario"]')))
        driver.find_element('xpath', '//*[@id="identificationForm:txtUsuario"]').send_keys(username)

        print("Filled in username")
        sleep(1)

        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="identificationForm:txtSenha"]')))
        driver.find_element('xpath', '//*[@id="identificationForm:txtSenha"]').send_keys(password)

        print("Filled in password")

        sleep(2)

        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="identificationForm:botaoAvancar"]')))
        
        driver.find_element('xpath', '//*[@id="identificationForm:botaoAvancar"]').click()

    except Exception as e:
        print("Login worked on first try")
        pass

    print('Login Successful')

    sleep(5)

    sleep(5)

    try: 

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'modalBoxAlertPrincipal')))
        
        driver.refresh()

        sleep(5)
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="identificationForm:txtUsuario"]')))
        
        driver.find_element('xpath', '//*[@id="identificationForm:txtUsuario"]').send_keys(username)

        sleep(1)

        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="identificationForm:txtSenha"]')))
        
        driver.find_element('xpath', '//*[@id="identificationForm:txtSenha"]').send_keys(password)

        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="identificationForm:botaoAvancar"]')))
        
        driver.find_element('xpath', '//*[@id="identificationForm:botaoAvancar"]').click()
        
    except Exception as _:
        print("Popup not found")
        pass

    sleep(3)

    try:
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, 'jqmOverlay')))
        print("Body bug detected")

        driver.refresh()

    except Exception as _:
        print("Body bug not found")
        pass


    print("LOGIN FINISHED")


def change_account(driver):

    sleep(3)
    change_account_button = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, 'lnkGrupoEconomico')))
    
    print("Located button")
    driver.refresh()

    try:

        change_account_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, 'lnkGrupoEconomico')))
        print("Found clickable")
        change_account_button.click()

    except Exception as e:

        body = driver.find_element(By.XPATH, '/html/body[1]/div')
        body.click()

        sleep(2)

        change_account_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, 'lnkGrupoEconomico')))
        print("Found clickable")
        change_account_button.click()    


    print('tbody found')

    sleep(4)
    try:

        iframe = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'modal_infra_estrutura')))

        # Switch focus to the iframe
        driver.switch_to.frame(iframe)

        company_row = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="tabelaEmpresas"]/tbody/tr[2]')))

        company_row.click()

        # Switch back to the main window (outside the iframe)
        
    except Exception as e:
        print(e)
        pass
        sleep(20)

    driver.switch_to.default_content()

    try:
        body = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body[1]')))
        # driver.refresh()
        body.click()
        sleep(3)
        # driver.refresh()
    except Exception as e:

        pass


def navigate_to_voucher(driver):

    sleep(3)

    driver.refresh()

    body = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body[1]')))

    #  driver.refresh()

    body.click()

    
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, '_id74_1:_id76')))
    print("payments FOUND")

    payments = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.ID, '_id74_1:_id76')))
    print("payments CLICKABLE found")
    sleep(2)
    payments.click()

    print("payments clicked")

    sleep(2)
    
    iframe = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'paginaCentral')))

    # Switch focus to the iframe
    driver.switch_to.frame(iframe)

    print("searching for vouchers button....")

    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a[title='vouchers (2nd Copy)']")))

    print("vouchers FOUND")

    vouchers = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a[title='vouchers (2nd Copy)']")))

    print("vouchers CLICKABLE found")

    vouchers.click()

    print("vouchers clicked")

    try:           
        print("Searching for BRADESCO voucher")
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "frm:_id104")))

        print("BRADESCO voucher FOUND")

        bradesco_voucher = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.ID, "frm:_id104")))

        print("BRADESCO voucher CLICKABLE found")

        bradesco_voucher.click()

        print("BRADESCO voucher clicked.........")

        sleep(4)

    except Exception:
            driver.refresh()

            body = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body[1]')))

            #  driver.refresh()

            body.click()

            WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, '_id74_1:_id76')))
            print("payments FOUND")

            payments = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.ID, '_id74_1:_id76')))
            print("payments CLICKABLE found")

            payments.click()

            print("payments clicked")

            sleep(2)
            
            iframe = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'paginaCentral')))

            # Switch focus to the iframe
            driver.switch_to.frame(iframe)

            print("searching for vouchers button....")

            WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a[title='vouchers (2nd Copy)']")))

            print("vouchers FOUND")

            vouchers = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a[title='vouchers (2nd Copy)']")))

            print("vouchers CLICKABLE found")

            vouchers.click()

            print("vouchers clicked")

            print("Searching for BRADESCO voucher")
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "frm:_id104")))

            print("BRADESCO voucher FOUND")

            bradesco_voucher = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.ID, "frm:_id104")))

            print("BRADESCO voucher CLICKABLE found")

            bradesco_voucher.click()

            print("BRADESCO voucher clicked.........")            


def select_filters(driver, consultation_date):

    def select_account_and_day(driver):
        
        sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'contaDebito__sexyCombo')))

        select_accounts = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'contaDebito__sexyCombo')))
        
        select_accounts.click()

        sleep(2)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="frm:_id123"]')))

        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="frm:_id123"]')))
        sleep(2)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="contaDebito__sexyCombo"]')))

        input_account = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="contaDebito__sexyCombo"]')))
        input_account.click()

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="frm:outputComboContas"]/div/div/ul')))

        ul = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="frm:outputComboContas"]/div/div/ul')))

        lis = ul.find_elements(By.TAG_NAME, 'li')

        for li in lis:
            # Check if the element text is equal to '2367 | 0009826-4'
            if li.text.strip() == '2367 | 0009826-4':
                # Perform whatever action you want with this element, like clicking on it
                li.click()
                break 

        sleep(3)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="frm:_id151"]')))

        detailed_search = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="frm:_id151"]')))

        detailed_search.click()

        
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="frm:datasPeriodoSemMes_beginDia"]')))

        date_field = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="frm:datasPeriodoSemMes_beginDia"]')))
        
        date_field.click()

        date_field.send_keys(consultation_date)

        sleep(2)

        body = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))

        body.send_keys(Keys.RETURN)
        sleep(2)


    def open_filter_options(driver):

        print("SEARCHING FOR voucher OPTIONS......")
        
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'cmbOperacoes')))

        print("voucher OPTIONS FOUND")

        select_filters = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'cmbOperacoes')))
        
        select_filters.click()

        print("voucher OPTIONS CLICKED")


    def select_and_save(driver, filter_name, today_date):

        sleep(2)

        try:

            no_voucher_text = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="frm:divResultado"]/div/div/p')))   

            no_voucher_text = no_voucher_text.text 

            print(f"{no_voucher_text}: {filter_name}")

            return True

        except Exception as e:
            
            print("SEARCHING FOR SELECT")
            select = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="chkAll"]')))

            sleep(3)
            print("SELECT FOUND")

            select.click()
            sleep(3)

            print("SEARCHING FOR SAVE")

            try:
                save = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="_id14"]/form/input')))

            except Exception as e:
                save = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="frm:divResultado"]/ul[1]/li[4]/form/input')))

            print("SAVE FOUND")

            save.click()

            sleep(2)
            
            driver.switch_to.default_content()

            sleep(2)

            iframe2 = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'modal_infra_estrutura')))

            # Switch focus to iframe2
            driver.switch_to.frame(iframe2)

            sleep(2)
            print("SEARCHING FOR pdf")
            pdf = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="formSalvarComo:pdf"]')))

            print("pdf FOUND")
            pdf.click()

            sleep(2)
            print("SEARCHING FOR close button")
            close = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="_id29"]')))

            print("close FOUND")
            close.click()

            print("close CLICKED")

            return False


    def send_voucher(emails, filter_name):

        script_directory = get_exe_directory() # obtain the script directory path

        file_type = '*.pdf'  # if you do not want to filter by extension just leave *

        file_paths = glob(path.join(script_directory, file_type))
        
        try:

            if len(file_paths) == 1:

                pdf_path = file_paths[0]
                # print("Excel file found:", excel_path)


                for email in emails:

                    send_email(
                        p_from = "osirisdatantech@gmail.com" , p_to = email, 
                        p_subject = "vouchers from Bradesco Bank RPA", 
                        p_text = f'voucher related to the filter {filter_name}', 
                        p_html = '',
                        p_filename = pdf_path
                        )
            
              
                print("email sent")

                sleep(1)

                os.remove(pdf_path)
                    
        except IndexError:
            print("Error: ERROR SEARCHING OR SENDING EMAIL WITH voucher")



    today_date = date.today()
    
    today_date = today_date.strftime('%d/%m/%Y')

    open_filter_options(driver)

    # position of the filters
    filter_list = [3, 4, 5, 8, 19, 21, 22]    

    email_list = ['receiver_email@email.com']


    for i in filter_list:

        sleep(2)
        # logic to click on the first filter 

        driver.switch_to.default_content()

        iframe3 = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'paginaCentral')))

        # Switch focus to iframe
        driver.switch_to.frame(iframe3)
        
        sleep(2)
        print("INPUT")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="frm:_id120"]')))
        print("FILTER FOUND")
        input = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="frm:_id120"]')))
        print("input CLICKABLE")
        sleep(2)

        try:
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="frm:_id493"]')))
            print("FILTER FOUND")
            close_overlay = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="frm:_id493"]')))
            close_overlay.click()
        except Exception as e:

            pass
            

        sleep(2)

        input.click()

        print("SEARCHING FOR FILTER")

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, f'//*[@id="cmbOperacoes"]/option[{i}]')))

        print("FILTER FOUND")

        filter = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, f'//*[@id="cmbOperacoes"]/option[{i}]')))
        print("FILTER CLICKABLE")

        filter.click()

        print("FILTER CLICKED")

        sleep(2)

        filter_name = filter.text

        select_account_and_day(driver)

        sleep(2)

        text_filter_check = select_and_save(driver, filter_name = filter_name, today_date = today_date)

        if text_filter_check:
            continue

        sleep(2)
        print("Sending email ....")
        send_voucher(emails = email_list, filter_name = filter_name)


    sleep(10)
    print("#"*70)
    print("RPA FINISHED, ALL voucherS DOWNLOADED")
    driver.quit()
    sleep(15)



def main(start_date, end_date):

    driver = open_site()

    login(driver=driver, login = 'XXXXXXXX', senha = '*******')

    change_account(driver=driver)

    navigate_to_voucher(driver = driver)

    search_date = str(start_date) + str(end_date)
    search_date = search_date.replace('/','')

    select_filters(driver = driver, search_date=search_date)
   

sg.theme('BlueMono')
output = sg.Text()
layout = [
    [sg.Text('LAYOUT EXAMPLE: ', size=(30, 3))],
    [sg.Text('FROM:   01/05/2024', size=(20, 1))],
    [sg.Text('TO:   05/06/2024', size=(20, 2))],
    [sg.Text('FROM:  ', size=(10, 1)), sg.InputText(key='-FROM-')],
    [sg.Text('TO:  ', size=(10, 1)), sg.InputText(key='-TO-')],
    [output],
    [sg.Button('Ok'), sg.Button('Cancel')]
]

window = sg.Window('RPA PROJECT TO DOWNLOAD VOUCHERS', layout, size=(590, 250))

while True:
    event, values = window.read()   
    if event == sg.WINDOW_CLOSED or event == 'Cancel':
        break   
    if event == 'Ok':
        start_date = values['-DE-']
        end_date = values['-ATE-']
        if start_date.strip() == '' or end_date.strip() == '' or len(str(end_date.strip())) < 10:
            output.update(value='Please provide both dates correctly, follow the format (01/05/2024).')
        else:
            name = main(start_date=start_date, end_date=end_date)
            output.update(value=str(name))

window.close()



