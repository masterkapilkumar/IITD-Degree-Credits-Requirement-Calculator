from bs4 import BeautifulSoup
import pandas as pd
import traceback
import argparse
import requests
import json
import xlwt
from xlwt import Workbook
import sys

from requests.packages.urllib3.exceptions import InsecureRequestWarning
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)


class DegreeCalc:
    
    def __init__(self, kerberos_username, kerberos_password, gradesheet_filename="gradesheet.xlsx"):
        self.url = "https://academics1.iitd.ac.in/Academics/index.php"
        self.username = kerberos_username
        self.password = kerberos_password
        self.sheet_name = gradesheet_filename
        self.headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36'}
        self.departments = ["BB1", "BB5", "CS5"]
        self.categories = {"BS":"Basic Sciences", "EA":"Engineering Arts and Science", "PL":"Programme-linked", "HU":"Humanities and Social Sciences", "OC":"Open Category", "DC":"Departmental Core", "DE":"Departmental Electives", "PC":"Programme Core", "PE":"Programme Elective", "NR":"Non Graded"}
        self.grade_points = {"A":10,"A-":9,"B":8,"B-":7,"C":6,"C-":5,"D":4}

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
    
    def generate_gradesheet(self, data):
        
        writer = pd.ExcelWriter(self.sheet_name)
        row=1
        col=1
        df = pd.DataFrame(columns=data.columns)
        columns = ['Course Code', 'Course Description', 'Course Category', 'Course Credits', 'Grade', 'Gradepoints']
        df.to_excel(writer,'gradesheet',startcol=col,startrow=0, columns=columns, index=False)
        
        btech_creds = 0
        btech_gpoints = 0
        mtech_creds = 0
        mtech_gpoints = 0
        
        for category in self.categories:
            category_name = self.categories[category]
            df = pd.DataFrame(columns=[category_name])
            df.to_excel(writer,'gradesheet',startcol=0,startrow=row, columns=[category_name], index=False)
            row += 1
            df = data.loc[data['Course Category'] == category]
            df.to_excel(writer,'gradesheet',startcol=col,startrow=row, columns=columns, index=False, header=False)
            row += df.shape[0]+2
            
            if(category in ['PC','PE']):
                mtech_creds += df[(df["Grade"].isin(self.grade_points))]["Course Credits"].sum()
                mtech_gpoints += df[(df["Grade"].isin(self.grade_points))]["Gradepoints"].sum()
            elif(category in ["BS", "EA", "HU", "PL", "DC", "DE", "OC"]):
                btech_creds += df[(df["Grade"].isin(self.grade_points))]["Course Credits"].sum()
                btech_gpoints += df[(df["Grade"].isin(self.grade_points))]["Gradepoints"].sum()
        
        
        cgpa = (btech_gpoints+mtech_gpoints)/(btech_creds+mtech_creds)
        btech_cgpa = btech_gpoints/btech_creds
        mtech_cgpa = mtech_gpoints/mtech_creds
        pd.DataFrame(columns=["B.Tech. Credits", btech_creds]).to_excel(writer,'gradesheet',startcol=4,startrow=row, index=False)
        row+=1
        pd.DataFrame(columns=["B.Tech. Gradepoints", btech_gpoints]).to_excel(writer,'gradesheet',startcol=4,startrow=row, index=False)
        row+=1
        pd.DataFrame(columns=["B.Tech. CGPA", btech_cgpa]).to_excel(writer,'gradesheet',startcol=4,startrow=row, index=False, float_format="%0.3f")
        row+=2
        
        pd.DataFrame(columns=["M.Tech. Credits", mtech_creds]).to_excel(writer,'gradesheet',startcol=4,startrow=row, index=False)
        row+=1
        pd.DataFrame(columns=["M.Tech. Gradepoints", mtech_gpoints]).to_excel(writer,'gradesheet',startcol=4,startrow=row, index=False)
        row+=1
        pd.DataFrame(columns=["M.Tech. CGPA", mtech_cgpa]).to_excel(writer,'gradesheet',startcol=4,startrow=row, index=False, float_format="%0.3f")
        row+=2
        pd.DataFrame(columns=["Overall CGPA", cgpa]).to_excel(writer,'gradesheet',startcol=4,startrow=row, index=False, float_format="%0.3f")
        
        writer.save()
        
        print("B.Tech. CGPA: %0.3f" % btech_cgpa)
        print("M.Tech. CGPA: %0.3f" % mtech_cgpa)
        print("Overall CGPA: %0.3f" % cgpa)
    
    def get_table_from_html(self, data):
        soup = BeautifulSoup(data, 'html.parser')
        tables = soup.findAll('table', attrs={'width':'900'})
        semester_data = []
        for semester in tables:
            courses = semester.find_all('tr')
            courses_data = []
            for course in courses[1:]:
                cols = course.find_all('td')
                course_details = [col.text for col in cols]
                courses_data.append(course_details)
            courses_data = pd.DataFrame(courses_data, columns=["Serial No.", "Course Code", "Course Description", "Course Category", "Course Credits", "Grade"])
            semester_data.append(courses_data)
        sem_data_merged = pd.concat(semester_data)
        sem_data_merged = sem_data_merged.sort_values(by=['Course Category'])
        sem_data_merged['Course Credits'] = pd.to_numeric(sem_data_merged['Course Credits'])
        return sem_data_merged
    
    def get_grades_url(self, data):
        soup = BeautifulSoup(data, 'html.parser')
        for link in soup.findAll('a'):
            if(link.text.strip()=="View Consolidated Grade Sheet"):
                return self.url+"?"+link.get('href').split("?")[1]
    
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
        
        self.grades_url = self.get_grades_url(response)
        (grades, _) = self.get_response(self.grades_url, data='text')
        
        # self.test_func()
        # grades = self.response
        table = self.get_table_from_html(grades)
        table['Gradepoints'] = table.apply(lambda row: (float(row['Course Credits']) * self.grade_points.get(row['Grade'], 0)), axis=1)
        self.generate_gradesheet(table)
        
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
    