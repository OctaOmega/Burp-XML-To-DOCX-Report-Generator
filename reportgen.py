'''''
Burp XML To Docx Report Generator
Version 2.0
Developer: Rajesh kannan 
Report Bugs @ rkannanm@outlook.com
'''''

from docxtpl import DocxTemplate
from socket import getservbyname as getportnumber
from html.parser import HTMLParser
import base64
from bs4 import BeautifulSoup
import PySimpleGUI as sg
from os import _exit


# Variables

xmlInFile = None #XML file 
Dest_folder = None # Destination Folder path
malcode = "NA"
scan_type = "NA"
from_date = "NA"
to_date = "NA"
type_application = "NA"
name_application = "NA"
line_of_business = "NA"
btrm_name = "NA"
appsec_no = "NA"
requester_name = "NA"
assessor_name = "NA"
assement_comp_date = "NA"
tech_contact = "NA"
clarity_code = "NA"
image_logo = ("burptdoc.png") #APP Logo
win_icon = ("app.ico") #APP ICON
docx_template_file = ("template.docx") #Report Template

# Layouts

logo_layout = [[sg.Image(image_logo,  size=(475,40))]]

file_list_column =[
    [
        sg.Text("Select XML File    "),
        sg.In(size=(45,1), enable_events=True, key="-XMLFILE-"),
        sg.FileBrowse(file_types=(("XML Files", "*.xml"),)),
        sg.Push(),
        sg.Button("Generate", size=(17,1), enable_events=True, key="-GENERATE-"),
    ],
    [
        sg.Text("Report Destination"),
        sg.In(size=(45,1), enable_events=True, key="-FOLDER-"),
        sg.FolderBrowse(),
        sg.Push(),
        sg.Button("Close", size=(17,1)),
    ],
]

input_Data1 = [
    [sg.Text("Malcode"), sg.Push(),sg.Input(size=(25,1), enable_events=True, key='-MALCODE-')],
    [sg.Text("Scan Type"), sg.Push(),sg.Input(size=(25,1), enable_events=True, key='-SCANTYPE-')],
    [sg.Text("From Date"), sg.Push(),sg.Input(size=(25,1), enable_events=True, key='-FRMDATE-')],
    [sg.Text("To Date"), sg.Push(),sg.Input(size=(25,1), enable_events=True, key='-TODATE-')],
    [sg.Text("Application Type"), sg.Push(),sg.Input(size=(25,1), enable_events=True, key='-APPTYPE-')],
    [sg.Text("Application Name"), sg.Push(),sg.Input(size=(25,1), enable_events=True, key='-APPNAME-')],
    [sg.Text("Line Of Business"), sg.Push(),sg.Input(size=(25,1), enable_events=True, key='-LOB-')],
]

input_Data2 = [
    [sg.Text("Application Owner"), sg.Push(),sg.Input(size=(25,1), enable_events=True, key='-APPOWNER-')],
    [sg.Text("SOW Reference"), sg.Push(),sg.Input(size=(25,1), enable_events=True, key='-REFNO-')],
    [sg.Text("Scan Requested By"), sg.Push(),sg.Input(size=(25,1), enable_events=True, key='-REQNAME-')],
    [sg.Text("Tester Name"), sg.Push(),sg.Input(size=(25,1), enable_events=True, key='-TESTNAME-')],
    [sg.Text("Scan Compilation Date"), sg.Push(),sg.Input(size=(25,1), enable_events=True, key='-COMPDATE-')],
    [sg.Text("Technical Contact"), sg.Push(),sg.Input(size=(25,1), enable_events=True, key='-TECHCONT-')],
    [sg.Text("Cost Center"), sg.Push(),sg.Input(size=(25,1), enable_events=True, key='-FINCODE-')],
]

logging_layout = [[sg.Text("Report Generation Progress:")],
                  [sg.ProgressBar(100, orientation='h', size=(75, 20), key='-PROGRESS BAR-')],
                  [sg.Multiline(size=(95,25), font='Courier 8', expand_x=True, expand_y=True, write_only=True,
                               reroute_stdout=True, reroute_stderr=True, echo_stdout_stderr=True, autoscroll=True, auto_refresh=True)]
                    # [sg.Output(size=(60,15), font='Courier 8', expand_x=True, expand_y=True)]
                  ]

layout = [
    [
        [sg.Column(logo_layout)], 
        [sg.Frame('Select Files:',file_list_column)], 
        [sg.Frame('Scan Descriptions:',[[sg.Column(input_Data1),sg.Column(input_Data2)]])],
        sg.Column(logging_layout),
    ]
]

window = sg.Window("Burp XML To Docx Report Generator", layout, icon=win_icon)

# Reading Events from Window

while True:
    event, values = window.read()
    progress_bar = window['-PROGRESS BAR-']
    if event == "-XMLFILE-":
        xmlInFile = values["-XMLFILE-"]
    if event == "-FOLDER-":
        Dest_folder = values["-FOLDER-"]
    if event == "-MALCODE-":
        malcode =  values["-MALCODE-"]
    if event == "-SCANTYPE-":
        scan_type =  values["-SCANTYPE-"]
    if event == "-FRMDATE-":
        from_date =  values["-FRMDATE-"]
    if event == "-TODATE-":
        to_date =  values["-TODATE-"]
    if event == "-APPTYPE-":
        type_application =  values["-APPTYPE-"]
    if event == "-APPNAME-":
        name_application =  values["-APPNAME-"]
    if event == "-LOB-":
        line_of_business =  values["-LOB-"]
    if event == "-APPOWNER-":
        btrm_name =  values["-APPOWNER-"]
    if event == "-REFNO-":
        appsec_no =  values["-REFNO-"]
    if event == "-REQNAME-":
        requester_name =  values["-REQNAME-"]
    if event == "-TESTNAME-":
        assessor_name =  values["-TESTNAME-"]
    if event == "-COMPDATE-":
        assement_comp_date =  values["-COMPDATE-"]
    if event == "-TECHCONT-":
        tech_contact =  values["-TECHCONT-"]
    if event == "-FINCODE-":
        clarity_code =  values["-FINCODE-"]
    if event == "-GENERATE-":
        current_count = 0
        try:
            print(" ")
            print("::::Burp To Docx Report Generator::::")
            print(" ")
            if xmlInFile:
                try:
                    doc = DocxTemplate(docx_template_file)
                    soup = BeautifulSoup(open(xmlInFile, 'r'), 'html.parser')
                    print('Using XML Input File {}'.format(xmlInFile))
                except(FileNotFoundError, NameError):
                    print('File Not Found - Select XML file to proceed!!!')
            else:
                print('File Not Found - Select XML file to proceed!!!')

            if Dest_folder:
                print('Report Destination Folder: '+Dest_folder)
            else:
                print('Destination Path Not Found - Select PATH to proceed!!!')

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

            try:
                #Log file creation        
                log_file=open(str(f'{Dest_folder}\Log_{malcode}.txt'), 'w')
                #Skip file creation 
                skipped_file=open(str(f'{Dest_folder}\Skipped_Issues_{malcode}.txt'), 'w')
            except (ValueError, RuntimeError, TypeError, NameError):
                print("Oops! Unable to write log files to Destination Path")
            for i in issues:
                    current_count+=5
                    progress_bar.update(current_count)
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
                        print(" ")
                        print('{} ({}) Risk: Has already been reported on - Skipping!! More info on skipped_issues.txt'.format(name, severity))
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
            doc.save(f'{Dest_folder}\Final_Report_For_{scan_type}_{malcode}.docx')
            print("---------------------------------------------------")
            print(':::::::Successfully Generated Your report:::::::::')
            print("---------------------------------------------------")
            log_file.writelines("Report Generated")
            log_file.close()
            progress_bar.update(current_count=100)
        except (ValueError, RuntimeError, TypeError, NameError):
            print("Oops!  Invalid Input.  Try again...")

     #Window Close event
    if event == "Close" or event =="Exit" or event == sg.WIN_CLOSED:
        window.close()
        _exit(0) 