from bs4 import BeautifulSoup
import pandas as pd
import traceback
import argparse
import requests
import json
import sys
import re

from requests.packages.urllib3.exceptions import InsecureRequestWarning
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)


class DegreeCalc:
    
    def __init__(self, kerberos_username, kerberos_password, gradesheet_filename="gradesheet.xlsx"):
        self.url = "https://academics1.iitd.ac.in/Academics/index.php"
        self.username = kerberos_username
        self.password = kerberos_password
        self.sheet_name = gradesheet_filename
        self.headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36'}
        self.categories = {"BS":"Basic Sciences", "EA":"Engineering Arts and Science", "PL":"Programme-linked", "HU":"Humanities and Social Sciences", "OC":"Open Category", "DC":"Departmental Core", "DE":"Departmental Electives", "PC":"Programme Core", "PE":"Programme Elective", "NR":"Non Graded", "OE": "Open-Elective"}
        self.departments = ["BB1", "BB5", "CH1", "CH7", "CE1", "CS1", "CS5", "EE1", "EE3", "ME1", "ME2", "MT1", "MT6", "PH1", "TT1"]
        self.requirements = {
        "BB1": {"PL":11, "DC":69, "DE":10, "OC":10, "PC":0, "PE":0, "NR":15},
        "BB5": {"PL":11, "DC":63, "DE":6, "OC":4, "NR":15, "PC":32, "PE":16},
        "CH1": {"PL":7, "DC":67, "DE":12, "OC":10, "PC":0, "PE":0, "NR":15}, 
        "CH7": {"PL":7, "DC":63, "DE":9, "OC":3, "NR":15, "PC":33, "PE":12},
        "CE1": {"PL":10, "DC":66, "DE":14, "OC":10, "PC":0, "PE":0, "NR":15},
        "CS1": {"PL":14, "DC":55, "DE":11, "OC":10, "PC":0, "PE":0, "NR":15}, 
        "CS5": {"PL":14, "DC":49, "DE":11, "OC":10, "NR":15, "PC":32, "PE":14},
        "EE1": {"PL":15, "DC":60, "DE":10, "OC":10, "PC":0, "PE":0, "NR":15}, 
        "EE3": {"PL":14, "DC":60, "DE":10, "OC":10, "NR":15, "PC":0, "PE":0},
        "ME1": {"PL":11, "DC":64, "DE":12, "OC":10, "NR":15, "PC":0, "PE":0},
        "ME2": {"PL":11, "DC":66, "DE":12, "OC":10, "NR":15, "PC":0, "PE":0},
        "MT1": {"PL":12.5, "DC":63.5, "DE":12, "OC":10, "PC":0, "PE":0, "NR":15}, 
        "MT6": {"PL":12.5, "DC":59.5, "DE":6, "OC":12, "NR":15, "PC":24, "PE":18}, 
        "PH1": {"PL":14.5, "DC":58, "DE":12, "OC":10, "PC":0, "PE":0, "NR":15},
        "TT1": {"PL":12, "DC":52, "DE":16, "OC":10, "PC":0, "PE":0, "NR":15}}
        for req in self.requirements:
            self.requirements[req]["BS"]=22
            self.requirements[req]["EA"]=18
            self.requirements[req]["HU"]=15
            self.requirements[req]["OE"]=0
        self.requirements["CH7"]["OE"]=3
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
    
    def generate_gradesheet_report(self, data):
        
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
        
        reqrmnt = {"BS":0, "EA":0, "HU":0, "PL":0, "DC":0, "DE":0, "OC":0, "OE":0, "PC":0, "PE":0, "NR":0}
        extra = {"BS":0, "EA":0, "HU":0, "PL":0, "DC":0, "DE":0, "OC":0, "OE":0, "PC":0, "PE":0, "NR":0}
        
        for category in self.categories:
            category_name = self.categories[category]
            df = pd.DataFrame(columns=[category_name])
            df.to_excel(writer,'gradesheet',startcol=0,startrow=row, columns=[category_name], index=False)
            row += 1
            df = data.loc[data['Course Category'] == category]
            df.to_excel(writer,'gradesheet',startcol=col,startrow=row, columns=columns, index=False, header=False)
            row += df.shape[0]+2
            
            if(category in ['PC','PE']):
                creds_earned = df[(df["Grade"].isin(self.grade_points))]["Course Credits"].sum()
                mtech_creds += creds_earned
                mtech_gpoints += df[(df["Grade"].isin(self.grade_points))]["Gradepoints"].sum()
            elif(category in ["BS", "EA", "HU", "PL", "DC", "DE", "OC", "OE"]):
                creds_earned = df[(df["Grade"].isin(self.grade_points))]["Course Credits"].sum()
                btech_creds += creds_earned
                btech_gpoints += df[(df["Grade"].isin(self.grade_points))]["Gradepoints"].sum()
            
            creds_earned = df[df["Grade"].isin(list(self.grade_points)+["NP", "S"])]["Course Credits"].sum()
            reqrmnt[category] = max(0, self.requirements[self.department][category] - creds_earned)
            extra[category] = max(0, creds_earned - self.requirements[self.department][category])
        
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
        
        if(sum(list(reqrmnt.values()))>0):
            print("Pending degrees requirements: ")
            for category in reqrmnt:
                if(reqrmnt[category]>0):
                    print("%s: %0.1f credits" % (self.categories[category], reqrmnt[category]))
        
        print()
        if(sum(list(extra.values()))>0):
            print("Extra credits completed: ")
            for category in extra:
                if(extra[category]>0):
                    print("%s: %0.1f credits" % (self.categories[category], extra[category]))
            print()
        
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
    
    def extract_department(self):
        try:
            self.department = re.search(r'([A-Z][A-Z][A-Z]?[0-9]).*', self.username.upper()).group(1)
        except AttributeError:
            print("Invalid kerberos id")
            sys.exit(1)
    
    def log_in(self):
        payload = {"username":self.username,"password":self.password, "submit-button":"Log+in", "page":"tryLogin"}
        
        #log in
        self.extract_department()
        self.session = requests.Session()
        response = self.session.post(self.url, data=payload, verify=False)
        if(response.status_code==200):
            return str(response.text)
        return False
    
    def run(self):
        # self.extract_department()
        response = self.log_in()
        if(response==False):
            raise Exception("Login failed with the following error: \n%s" %(response))
        
        self.grades_url = self.get_grades_url(response)
        (grades, _) = self.get_response(self.grades_url, data='text')
        
        # self.test_func()
        # grades = self.response
        table = self.get_table_from_html(grades)
        table['Gradepoints'] = table.apply(lambda row: (float(row['Course Credits']) * self.grade_points.get(row['Grade'], 0)), axis=1)
        self.generate_gradesheet_report(table)
        
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
    