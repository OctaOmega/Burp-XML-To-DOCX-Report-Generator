from docxtpl import DocxTemplate
from socket import getservbyname as getportnumber
from html.parser import HTMLParser
import base64
from bs4 import BeautifulSoup

doc = DocxTemplate("template.docx")

print(" ")
print("::::APPSEC Report Template Tool::::")
print(" ")
print("WARNING: File Name is Case Sensitive")
print(" ")

while True:
    try:
        xmlInFile = input("Please Enter your Burp XML File Name [Example: Test.xml]: ")
        # parse an xml file by name
        soup = BeautifulSoup(open(xmlInFile, 'r'), 'html.parser')
        print('Using XML Input File {}'.format(xmlInFile))
        break
    except(FileNotFoundError):
        print(f'File Not Found !!!\n' 
            f'Please check if XML file is on the same directory\n'
            f'Please check the file name entered\n')

print(" ")
print("---------------------------------------------------")
print("Please Provide Possible Values for the Below Fields")
print("---------------------------------------------------")
print(" ")



# pull all issue tags from XML
issues = soup.findAll('issue')


class MLStripper(HTMLParser):

    def __init__(self):
        super().__init__()
        self.reset()
        self.fed = []

    def handle_data(self, d):
        self.fed.append(d)

    def get_data(self):
        return ''.join(self.fed)

def strip_tags(html):
    s = MLStripper()
    s.feed(html)
    return s.get_data()

# define globals
global issueList
global vulnList
global skippedVulnList
global objTblRows

issueList = []
vulnList = []
skippedVulnList = []
document_obj = []
objTblRows = []

while True:
    try:
        malcode = input("Application MAL code: ")
        scan_type = input("SCAN type [PEN/DAST]: ")
        from_date = input("SCAN Start Date: ")
        to_date = input("SCAN End Date: ")
        type_application = input("Application Type [WEB/API]: ")
        name_application = input("Application Name: ")
        line_of_business = input("Application LOB: ")
        btrm_name = input("Application BTRM: ")
        appsec_no = input("APPSEC TICKET Number [JIRA]: ")
        requester_name = input("SCAN Requested by: ")
        assessor_name = input("Tester: ")
        assement_comp_date = input("Assesment Completion Date: ")
        tech_contact = input("TECH Contact: ")
        clarity_code = input("Clarity Code: ")
        break
    except (ValueError, RuntimeError, TypeError, NameError):
        print("Oops!  Invalid Input.  Try again...")

#Log file creation        
log_file=open(str(f'Log_{malcode}.txt'), 'w')

#Skip file creation 
skipped_file=open(str(f'Skipped_Issues_{malcode}.txt'), 'w')

for i in issues:
        name = i.find('name').text
        host = i.find('host')
        ip = host['ip']
        host = host.text
        path = i.find('path').text
        location = i.find('location').text
        severity = i.find('severity').text
        confidence = i.find('confidence').text
        try:
            issueBackground = i.find('issuebackground').text
            issueBackground = str(issueBackground).replace('&','/')
            issueBackground = strip_tags(issueBackground).split()
            issueBackground = " ".join(issueBackground)

        except:
            issueBackground = ' '

        try:
            remediationBackground = i.find('remediationbackground').text
            remediationBackground = str(remediationBackground).replace('&','/')
            remediationBackground = strip_tags(remediationBackground).split()
            remediationBackground = " ".join(remediationBackground)

        except:
            remediationBackground = ' '

        try:
            request = i.find('requestresponse').text
            request = base64.b64decode(request).split()[0]
            request = str(request.decode('utf-8'))
        except:
            request = ' '

        try:

            issueDetail = i.find('issuedetail').text
            issueDetail = str(issueDetail).replace('&','/')
            issueDetail = strip_tags(issueDetail).split()
            issueDetail = " ".join(issueDetail)

        except:
            issueDetail = ' '

        issueLine = 'Processed Issue: [{}]'.format(str('{} ({})'.format(name, '{} Risk'.format(severity))))
        # Check here to see if issueLine.
        protocol_in_use, domain_path = host.split('://')
        port = getportnumber(protocol_in_use)
        if issueLine not in str(vulnList):
            # DATA for word document
            document_obj.append({
                severity: {
                                        "severity": severity,
                                        "name": name, 
                                        "host": host, 
                                        "port":port, 
                                        "path": host+path, 
                                        "param": request, 
                                        "issueDetail": issueDetail, 
                                        "issueBackground": issueBackground,
                                        "remediationBackground": remediationBackground
                                    }
            })
            vulnList.append(issueLine)
            print(issueLine)
        # If issue/vuln has already been reported on.
        if issueLine in str(vulnList):
            print("------------------------------------------------------------------------------------------------- ")
            print('{} ({}) Risk: Has already been reported on! Skipping!!'.format(name, severity))
            print('Please check skipped_issues.txt if you would like to include the skipped path to reported issues')
            print("------------------------------------------------------------------------------------------------- ")
            print(" ")
            sendSkipped = ({'name':name, 'severity':severity, 'host':host, 'path':host+path})
            skippedVulnList.append(sendSkipped)

        # Log creation
        result = ({'name':name, 'severity':severity, 'host':host, 'ip':ip, 'path':host+path, 'location':location, 'confidence':confidence, 'issueBackground':issueBackground, 
                        'issueDetail':issueDetail, 'remediationBackground':remediationBackground})
        issueList.append(result)
        
#Log data
for logs in issueList:
    for log in  logs.items():
       log_file.writelines(str(log))
       log_file.writelines("\n")

#Skipfile data
for skips in skippedVulnList:
    for skip in skips.items():
        skipped_file.writelines(str(skip))
        skipped_file.writelines("\n")
        skipped_file.writelines(" ")

skipped_file.close()

# DOCXTPL DATA
context = {  
"malcode" : malcode, 
"scan" : scan_type,
"from_date" : from_date,  
"to_date" : to_date,
"typeofappl" : type_application,
"nameofappl" : name_application,
"line_of_business" : line_of_business,
"btrm_name" : btrm_name,
"appsec_no" : appsec_no,
"requester_name" : requester_name,
"assessor_name" : assessor_name,
"assement_comp_date" : assement_comp_date,
"tech_contact" : tech_contact,
"clarity_code" : clarity_code,
"host": host,
"ip": ip,
"document_obj": document_obj
}

#Document rendering
doc.render(context)
doc.save(f'Final_Report_For_{scan_type}_{malcode}.docx')
print("---------------------------------------------------")
print(':::::::Successfully Generated Your report:::::::::')
print("---------------------------------------------------")
log_file.writelines("Report Generated")
log_file.close()