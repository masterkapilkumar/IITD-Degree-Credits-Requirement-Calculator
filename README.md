# IITD-Degree-Credits-Requirement-Calculator

Generates an excel sheet for the gradesheet by grouping the courses into different degree requirement categories. It also calculates DGPA for B.Tech. and M.Tech. degrees.


## How to use
	usage: credits_calculator.py [-h] config_file

	positional arguments:
	  config_file  path of JSON file having configuration data

	optional arguments:
	  -h, --help   show this help message and exit
    
The JSON configuration file should have the following data:
 - **username**: Kerberos User id.
 - **password**: Kerberos password.

### Dependencies
- Python 3.x
- Python libraries
    * requests
    * BeautifulSoup
    * pandas
