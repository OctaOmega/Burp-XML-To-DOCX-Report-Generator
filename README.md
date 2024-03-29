# Burp XML 2 DOCX 2.0
## Burp XML to Docx Report Generator (Converter)

#### Create MS Word reports from Burp XML file using Document (DOCX) Template.

## Installation

#### Clone git repository:

```
git clone https://github.com/OctaOmega/Burp-XML-To-DOCX-Report-Generator.git
```
#### Download as Zip:

https://github.com/OctaOmega/Burp-XML-To-DOCX-Report-Generator/archive/refs/heads/main.zip

## Windows 
```
run reportgen.exe from the Directory
```
#### 	**Please _do not remove or rename or replace_ any of the image files or template.docx** - This will break the application.
#### Refer jinja2 for managing tags or docxtpl documentation here - https://docxtpl.readthedocs.io/en/latest/ for customizing template.docx
#### Final Report.docx, Log.txt & Skipped.txt will be generated in the specified _Report Destination_ path.

### Samples Report and XML file are in Sample Directory 

### reportgen.exe for Windows OS  - GUI

![Screen1](https://user-images.githubusercontent.com/85091462/194997449-430cefe9-931d-4c69-9089-ab89607c97c1.jpg)

### reportgen.py source - Linux / Windows (Running as python Script)

#### Make sure to install the dependencies.

```
pip install docxtpl
pip install socket
pip install HTMLParser
pip install bs4
pip install PySimpleGUI
```
#### Change file Permission:
```
chmod +x reportgen.py
```
## Output

### Finalreport_{malcode}.docx file based on template.docx

![docx_report](https://user-images.githubusercontent.com/85091462/194998102-ff271772-a296-4cb2-9104-e3fff3177eee.jpg)

### log.txt (Log file)

![Log-txt](https://user-images.githubusercontent.com/85091462/194998128-d3422d07-6d94-4307-9a3f-d58f4b874044.jpg)

### Skipped.txt
#### Duplicate issues that are skipped from the report

![Duplicate issues that are skipped from the report](https://user-images.githubusercontent.com/85091462/194998161-d3b0d924-ade9-4282-b399-86068fe116d0.jpg)

