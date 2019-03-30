from bs4 import BeautifulSoup
import pandas as pd
import traceback
import argparse
import requests
import json
import sys

from requests.packages.urllib3.exceptions import InsecureRequestWarning
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)


class DegreeCalc:
    
    def __init__(self, kerberos_username, kerberos_password):
        self.url = "https://academics1.iitd.ac.in/Academics/index.php"
        self.username = kerberos_username
        self.password = kerberos_password
        self.headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36'}

    def get_response(self, url, headers=None, verify=False, type='GET', data="json"):
        if(not headers):
            headers = self.headers
        if(type=='GET'):
            response = requests.get(url, headers=headers, verify=verify)
            if(data=="json"):
                data = response.json()
            else:
                data = response.text
            return data, response.status_code
    
    def log_in(self):
        payload = {"username":self.username,"password":self.password, "submit-button":"Log+in", "page":"tryLogin"}
        
        #log in
        self.session = requests.Session()
        response = self.session.post(self.url, data=payload, verify=False)
        if(response.status_code==200):
            return str(response.text)
        return False
    
    def run(self):
        
        response = self.log_in()
        if(response==False):
            raise Exception("Login failed with the following error: \n%s" %(response))
        
        
    def test_func(self):
        self.response = """
       

        """
    
if __name__=='__main__':

    parser = argparse.ArgumentParser()
    parser.add_argument("config_file", help="path of JSON file having configuration data")
    
    args = parser.parse_args()
    
    try:
        fin = open(args.config_file, 'r')
        config_data = json.loads(fin.read().strip())
        fin.close()
    except:
        print("Error reading configuration data:\n")
        traceback.print_exc()
        sys.exit(1)
    
    try:
        kerberos_username = config_data["kerberos_username"]
        kerberos_password = config_data["kerberos_password"]
        
    except KeyError:
        print("Missing configuration data:\n")
        traceback.print_exc()
        sys.exit(1)
    
    calc = DegreeCalc(kerberos_username, kerberos_password)
    calc.run()
    