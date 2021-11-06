#importing the needed packages
import imaplib
import email
import pyautogui
from email.header import decode_header
import webbrowser
import os
import urllib.request
from urllib.parse import urlparse
import webbrowser
import unittest
from selenium import webdriver
from selenium.webdriver.common import keys
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
import time

#setting the gmail username and password as global variables
username = "sumergealerts@gmail.com"
password = "P@ssw0rd@2020"


#declaring the function which is used to login to the gmail and extract the needed parameters to be filled in the form
def get_Mail():

    #counting the instances of Unseen emails
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    imap.login(username, password)
    status, messages = imap.select("HR")
    messages = int(messages[0])
    response, messagess = imap.search(None, 'UnSeen')
    messagess = messagess[0].split()
    count_messeges =len(messagess)
    N = count_messeges 
    print (N)
    if N>0:
        login()
    else:
        driver.quit()
    for i in range(1, N+1):
    # fetch the email message by ID 
        print ("fetch the email message by ID>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>") 
        print(i)           
        res, msg = imap.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])
                # decode the email subject
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    # if it's a bytes, decode to str
                    subject = subject.decode(encoding)
                # decode email sender
                From, encoding = decode_header(msg.get("From"))[0]
                if isinstance(From, bytes):
                    From = From.decode(encoding)

                #testing if the subject is extracted successfully
                print("Subject:", subject)
                print(">>>>>>>>>>>>>>>>>>>>>>>>sayed>>>>>>>>>>>>>>>>>>>>>>>>>")
                #since there are 2 templates, this if condition is inserted to check which of the 2 templates
                #is the one we are currently extracting the data from
                if (subject.find('Application for') != -1):
                    Position = subject[subject.find('Application for') + 16:subject.find(' from')]
                    print(Position)
                elif (subject.find('New application:') != -1):
                    Position = subject[subject.find('New application:') + 17:subject.find(' from')]
                    print(Position)

                #for logging purposes
                print(">>>>>>>>>>>>>>>>>>>>>>>>sayed>>>>>>>>>>>>>>>>>>>>>>>>>")
                print(">>>>>>>>>>>>>>>>>>>>>>>>sayed>>>>>>>>>>>>>>>>>>>>>>>>>")

                #extracting the applicant's name
                if (subject.find('from') != -1):
                    Applicant_name = subject[subject.find('from') + 5:]
                    print(Applicant_name)
                print(">>>>>>>>>>>>>>>>>>>>>>>>sayed>>>>>>>>>>>>>>>>>>>>>>>>>")

                print("From:", From)
                Fromq = From
                # if the email message is multipart
                if msg.is_multipart():
                    # iterate over email parts
                    for part in msg.walk():
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            # get the email body
                            body = part.get_payload(decode=True).decode()
                        except:
                            pass
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            # print text/plain emails and skip attachments
                            # print the extracted data
                            print(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>")
                            print(body.find('Contact Information'))
                            if (body.find('Contact Information') != -1):
                                mail = body[body.find('Contact Information') + 19:body.find('<’mailto:')]
                                print(mail)
                            else:
                                mail = "sumerge@sumerge.com"
                            if (body.find('Phone:') != -1):
                                phone = body[body.find('Phone:') + 7:body.find('Resume:')]
                                print(phone)
                            else:
                                phone = "00000000000"

                            print(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>")
                            print(body.find('Contact Information'))
                            if (body.find('Education') != -1):
                                university = body[body.find('Education') + 9:body.find('University')]
                                print(university)

                        elif "attachment" in content_disposition:
                            # download attachment
                            filename = part.get_filename()
                            if filename:
                                folder_name = clean(subject)
                                if not os.path.isdir(folder_name):
                                    # make a folder for this email (named after the subject)
                                    os.mkdir(folder_name)
                                filepath = os.path.join(folder_name, filename)
                                # download attachment and save it
                                open(filepath, "wb").write(part.get_payload(decode=True))
                else:
                    # extract content type of email
                    content_type = msg.get_content_type()
                    # get the email body
                    body = msg.get_payload(decode=True).decode()
                    if content_type == "text/plain":
                        # print only text email parts
                        # print(body)
                        print(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>")
                        if (body.find('Contact Information') != -1):
                            mail = body[body.find('Contact Information') + 19:body.find('<’mailto:')]
                            print(mail)
                        else:
                            mail = "sumerge@sumerge.com"
                        if (body.find('Phone:') != -1):
                            phone = body[body.find('Phone:') + 7:body.find('Resume:')]
                            print(phone)
                        else:
                            phone = "00000000000"
                        
                        print(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>")
                        print(body.find('Contact Information'))
                        if (body.find('Education') != -1):
                            university = body[body.find('Education') + 9:body.find('University')]
                            print(university)

                if content_type == "text/html":
                    # if it's HTML, create a new HTML file and open it in browser
                    folder_name = clean(subject)
                    if not os.path.isdir(folder_name):
                        # make a folder for this email (named after the subject)
                        os.mkdir(folder_name)
                    filename = "index.html"
                    filepath = os.path.join(folder_name, filename)
                    # write the file
                    open(filepath, "w").write(body)
                    # open in the default browser
                    # webbrowser.open(filepath)
                print("=" * 100)
        #passing the extracted parameters to the make_application function
        make_application(Applicant_name, mail, phone, Position, university)
    # close the connection and logout
    imap.close()
    imap.logout()


def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)

#function for downloading the attachement in the email
def download_file(download_url, filename):
    #pdf_path = "C:\\Users\\selsayed\\Documents\\HR_Application"
    #pdf_path = "C:\\Users\\selsayed\\Documents\\HR_Application"
    try:
        response = urllib.request.urlopen(download_url)
        file = open(filename + ".pdf", 'wb')
        file.close()
    except:
        print(" We're sorry, we had to expire this link for security reasons. You can still view and download the applicant's resume: just re-open the applicant email and click on the button that says View full application.")

#function to login to the ATS portal
def login():
    #using the chrome driver to access the portal and login using the provided credentials
    driver.get("https://ats.sumerge.com/web?#view_type=form&model=hr.applicant&action=128")
    Login = driver.find_element_by_id("login")
    Login.send_keys("selsayed@sumerge.com")
    pwd = driver.find_element_by_id("password")
    pwd.send_keys("\dqC32LLS^Qdv-%g")
    Login_bttn = driver.find_element_by_xpath("//*[@id='wrapwrap']/main/div/form/div[3]/button")
    Login_bttn.click()
    # driver.get("https://www.google.com")
    # driver.execute_script("window.open('https://ats.sumerge.com/web?#view_type=form&model=hr.applicant&action=128');")
    # driver.quit()

#function used to fill in the application for each applicant
def make_application(Applicant_name, mail, phone, Position, university):
    driver.get('https://ats.sumerge.com/web?#view_type=form&model=hr.applicant&action=128')
    driver.get('https://ats.sumerge.com/web?#view_type=form&model=hr.applicant&action=128')
    driver.get('https://ats.sumerge.com/web?#view_type=form&model=hr.applicant&action=128')
    time.sleep(4)
    #filling in the job title
    ####################################
    job_name = driver.find_element_by_xpath("//*[@id='o_field_input_41']")
    job_name.send_keys(Position)
    time.sleep(2)
    list = driver.find_elements_by_tag_name("li")
    for e in list:
        if e.text == Position:
            e.click()
            break

    #filling in the applicant's name
    ####################################
    app_name = driver.find_element_by_id("o_field_input_4")
    app_name.send_keys(Applicant_name)

    #filling in the applicant's email
    ########
    Emaill = driver.find_element_by_id("o_field_input_15")
    Emaill.send_keys(mail)

    #filling in the CV type
    ########
    CVTyp = Select(driver.find_element_by_xpath("//*[@id='o_field_input_11']"))
    CVTyp.select_by_visible_text("Inbound")

    #filling in the time the form was submitted
    ########
    date = driver.find_element_by_xpath(
        "/html/body/div[1]/div[2]/div[2]/div/div/div[1]/div/div[3]/table[3]/tbody/tr[4]/td[2]/div/input")
    date.click()
    #driver.find_element_by_xpath("//body").click()

    #filling in the applicant's phone number
    ########
    phonee = driver.find_element_by_id("o_field_input_17")
    phonee.send_keys(phone)

    #filling in the how did you hear about us field
    ########
    know_name = driver.find_element_by_xpath("//*[@id='o_field_input_47']")
    know_name.send_keys('Linkedin')
    time.sleep(1)

    #filling in the gender field
    ########
    Gender = Select(driver.find_element_by_xpath("//*[@id='o_field_input_23']"))
    Gender.select_by_visible_text("Male")

    #filling in the responsible person's name
    ########
    resp_name = driver.find_element_by_xpath("//*[@id='o_field_input_28']")
    resp_name.send_keys("elsayed")
    list = driver.find_elements_by_tag_name("li")
    for e in list:
        if e.text == "elsayed":
            e.click()
            break
    #filling in the university's name
    ########
    time.sleep(1)
    instiute_name = driver.find_element_by_xpath("//*[@id='o_field_input_79']")
    instiute_name.send_keys(university + "University")

    #filling in the grad year
    ########
    GRAD_YEAR = driver.find_element_by_id("o_field_input_81")
    GRAD_YEAR.send_keys("0000")
    ###########
    #Submit the application
    submit = driver.find_element_by_xpath("/html/body/div[1]/div[2]/div[1]/div[2]/div[1]/div/div[2]/button[1]")
    submit.click()
    time.sleep(3)
    ##########
    #uploading the document by clicking on the document button after all the compulsory data has been filled
    #then using the pyautogui to autofill the path and choosing the file saved earlier
    add_cv = driver.find_element_by_xpath("/html/body/div[1]/div[2]/div[1]/div[2]/div[2]/div/div[2]/button")
    add_cv.click()
    time.sleep(1)
    
    list = driver.find_elements_by_tag_name("li")
    for e in list:
        if e.text == "Add...":
            e.click()
            break
    time.sleep(1)
    cv_path = "C:\\Users\\selsayed\\Documents\\HR_Application\\"
    Applicant_name = "ahmed"
    pyautogui.write(cv_path + Applicant_name + ".pdf", interval=0.02)
    pyautogui.press('return')
    #since the page takes quite some time for the file to be saved, halt the for loop for 40 seconds to avoid
    #the termination of the loop
    time.sleep(2)



if __name__ == "__main__":
    # get_Mail()
    # Url = urlparse('https://www.dundeecity.gov.uk/sites/default/files/publications/civic_renewal_forms.zip')
    # download_file(Url.geturl(),"ahmed")
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(chrome_options=options, executable_path=r'C:/Users/selsayed/Documents/HR_Application/HR_APP/chromedriver.exe')
    get_Mail()
    print("Done.")

