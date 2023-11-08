from io import TextIOWrapper
import os
import win32com.client as win32
import json
import PyPDF2
import re
import csv
from datetime import date

def main():
    #program start

    #VARS
    config:configuration = configuration()
    cache:cache_manager = cache_manager()
    pdf:pdf_reader=pdf_reader(config)
    email:email_manager=email_manager()

    #METHOD CALLS
    cache.init_cache(config)

    for pdf_path in pdf._pdf_path_list: #loop through each pdf
        if cache.find_pdf(config,pdf_path) == False: #if the pdf email was not cached(no email sent)
            contents:dict = pdf.parse_pdf(pdf_path)
            email.send_email(contents,config)
            cache.cache_email(contents,pdf_path,config)

    input("Press ENTER to close this window...")

class configuration:
    _config_file_path:str="./config.json"

    def __init__(self) -> None:
        pass
    
    def get_mailing_list(self)->str:
        file = open(self._config_file_path)
        data = json.load(file)
        file.close()
        return data["settings"][0]["mailing_list"] 

    def get_driver_emails_path(self)->str:
        #acquire the relative path
        file = open(self._config_file_path)
        data = json.load(file)
        file.close()
        email_path = data["settings"][0]["email_list"] 
        return [os.path.join(email_path, file) for file in os.listdir(email_path)]
    
    def get_pdf_files(self)->str:
        #acquire the relative path
        file = open(self._config_file_path)
        data = json.load(file)
        file.close()
        pdf_dir = data["settings"][0]["pdf_dir"]
        return [os.path.join(pdf_dir, file) for file in os.listdir(pdf_dir)]
    
    def get_cache_files(self)->str:
        #acquire the relative path
        file = open(self._config_file_path)
        data = json.load(file)
        file.close()
        cache_dir = data["settings"][0]["cache_dir"]
        return [os.path.join(cache_dir, file) for file in os.listdir(cache_dir)]

    def get_cache_dir(self)->str:
        file = open(self._config_file_path)
        data = json.load(file)
        file.close()
        cache_dir = data["settings"][0]["cache_dir"]
        return cache_dir
    
class email_manager:
    #sends emails
    def __init__(self) -> None:
        pass

    def send_email(self,contents,config:configuration):
        # construct Outlook application instance
        try:
            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetNameSpace('MAPI')

            # construct the email item object
            mailItem = olApp.CreateItem(0)
            mailItem.Subject = 'ROUTE: '+contents["Route ID"]
            
            names  =contents["Driver 1 Name"]
            if contents["Driver 2 Name"] != "":
                names += ", " + contents["Driver 2 Name"]
            mailItem.HTMLBody = """
            <p>Hello, this is an automated email to notify you of your route tomorrow!</p>
            <p>ROUTE: """ +contents["Route ID"] +"""</p>
            <p>DRIVERS: """+ names +"""</p>
            <p>-Company, Inc.</p>"""

            cc_mail_list = self.get_mail_list(config)
            to_mail_list = self.parse_driver_emails(contents)

            mailItem.CC = cc_mail_list#attach extra recipients
            mailItem.To = to_mail_list
            
            mailItem.Attachments.Add(os.path.join(os.getcwd(),contents["PDF FILE"]))

            if to_mail_list == "":
                print("ERROR: NO EMAILS FOUND FOR [ "+contents["Driver 1 Name"] + " AND " + contents["Driver 2 Name"] + " ]")
                return

            mailItem.Display()
            mailItem.Save()
            mailItem.Send()
            print("SENT EMAIL TO: "+to_mail_list)
        except:
            print("FAILED TO SEND EMAIL TO: "+to_mail_list)
            print("FILE: "+ contents["PDF FILE"])

    def parse_driver_emails(self,contents:dict)->str:
        file = open("./driverEmailList.csv",'r')
        csvFile = csv.reader(file)
        email_list:str=""
        for lines in csvFile:
            if contents["Driver 1 ID"] == lines[2] or contents["Driver 2 ID"] == lines[2]:#ID
               email_list += lines[3]+";"#Email
        file.close()
        return email_list

    def get_mail_list(self,config:configuration):
        list:str=""
        list+=config.get_mailing_list()
        return list

class pdf_reader:
    #parse the pdfs in the pdf folder as laid out in config.ini
    _pdf_path_list:list[str] = []#list of pdfs

    def __init__(self,config:configuration) -> None:
        self._pdf_path_list = config.get_pdf_files()

    def parse_pdf(self,pdf_path)->dict:
        contents:dict={}
        file = open(pdf_path,'rb')#open PDF

        pdf_parser = PyPDF2.PdfReader(file)
        pdf_page = pdf_parser.pages[0]#cache the first page

        pdf_text = pdf_page.extract_text()

        pdf_text_list = pdf_text.split("\n")

        for i in pdf_text_list:#loop through each line of text in pdf
            if i.find("Driver 1:")!=-1:#find driver 1
                #ID
                try:
                    driver_id = re.findall(r"--\s(\d+)",i)[0]
                    contents["Driver 1 ID"] = driver_id
                except:
                    contents["Driver 1 ID"] = ""
                #Name
                try:
                    driver_name = re.findall(r"[A-Za-z]+,\s[A-Za-z]+",i)[0]
                    contents["Driver 1 Name"] = driver_name
                    pass
                except:
                    contents["Driver 1 Name"] = ""
                    pass

            if i.find("Driver 2:")!=-1:#find driver 2
                #ID
                try:
                    driver_id = re.findall(r"--\s(\d+)",i)[0]
                    contents["Driver 2 ID"] = driver_id
                except:
                    contents["Driver 2 ID"] = ""
                #Name
                try:
                        driver_name = re.findall(r"[A-Za-z]+,\s[A-Za-z]+",i)[0]
                        contents["Driver 2 Name"] = driver_name
                        pass
                except:
                        contents["Driver 2 Name"] = ""
                        pass

            if i.find("Route ID:")!=-1:#find route ID
                route_id = i.split("Route ID:")[1]
                contents["Route ID"] = str.strip( route_id )
        file.close()#close PDF

        contents["PDF FILE"] = pdf_path
        return contents

class cache_manager:
    #responsible for caching information for already emailed drivers to prevent duplicate emails
    def __init__(self) -> None:
        pass

    def init_cache(self,config:configuration):
        #clears the cache of files with old time stamps
        for file_path in config.get_cache_files():
            file = open(file_path,'r')
            line = file.readline()
            file.close()
            if date.today().strftime("%m/%d/%Y") != line:
                os.remove(file_path)

    def find_pdf(self,config:configuration,pdf_path:str)->bool:
        for file in config.get_cache_files():#loop through each file in the cache dir
            pdf_file_name = re.findall(r"\\([^/]+)\.pdf$",pdf_path)[0]
            file_name = re.findall(r"\\([^/]+)\.txt$",file)[0]
            if pdf_file_name == file_name:
                return True
        return False

    def cache_email(self,contents:dict,pdf_path:str,config:configuration):
        pdf_file_name:str = re.findall(r"\\([^/]+)\.pdf$",pdf_path)[0]
        path:str=config.get_cache_dir()+"/"+pdf_file_name+".txt"
        file = open(path,"x")
        file.write(date.today().strftime("%m/%d/%Y"))
        file.close()
        pass
#call main
main()