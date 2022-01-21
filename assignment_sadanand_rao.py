import glob                         
import pandas as pd                 
import xml.etree.ElementTree as ET  

##### Provide File Path #####
filepath = "C:/Users/Owner/Documents/nuula/assignment/data" ##Enter your folder path here
xml_filepath = filepath + "/*.xml"

####### Functions Area #######
### Extract Errors
# Read all the xml files from the folder
def extract_errors(filepath):
    extracted_data = pd.DataFrame()
    for xmlfile in glob.glob(filepath):
        extracted_data = extracted_data.append(parse_errors(xmlfile), ignore_index=True)
    return extracted_data
# Extract all the Errors
def parse_errors(xml_file):
    df = pd.DataFrame()
    tree = ET.parse(xml_file)
    root = tree.getroot()
    data =[]
    for elem in root.iter("Nuula"):
        for sub_elem in elem:
            if sub_elem.tag == 'Errors':
                for error_elem in sub_elem:
                    data.append(error_elem.attrib)
    df = data
    return df

### Extract Messages
# Read all the xml files from folder
def extract_messages(filepath):
    extracted_data = pd.DataFrame()
    for xmlfile in glob.glob(filepath):
        extracted_data = extracted_data.append(parse_messages(xmlfile), ignore_index=True)
    return extracted_data
# Extract all the Messages
def parse_messages(xml_file):
    df = pd.DataFrame()
    tree = ET.parse(xml_file)
    root = tree.getroot()
    data =[]
    for elem in root.iter("DataExtract900jer"):
        for sub_elem in elem:
            if sub_elem.tag == 'Messages':
                for msg_elem in sub_elem:
                    data.append(msg_elem.attrib)
    df = data
    return df

### Extract Rules
# Read all the files xml files from folder
def extract_rules(filepath):
    extracted_data = pd.DataFrame()
    for xmlfile in glob.glob(filepath):
        extracted_data = extracted_data.append(parse_rules(xmlfile), ignore_index=True)
    return extracted_data
# Extract Rules from the xml
def parse_rules(xml_file):
    df = pd.DataFrame()
    tree = ET.parse(xml_file)
    root = tree.getroot()
    data =[]
    for elem in root.iter("DataExtract900jer"):
        for sub_elem in elem:
            if sub_elem.tag == 'Rules':
                for rule_elem in sub_elem:
                    data.append(rule_elem.attrib)
    df = data
    return df

###### Program Area #######

# Create data frames for individual objects
df_error = pd.DataFrame()
df_message = pd.DataFrame()
df_rule = pd.DataFrame()
# Write the data into data frame
df_error = extract_errors(xml_filepath)
df_message = extract_messages(xml_filepath)
df_rule = extract_rules(xml_filepath)

# Write the data frame to excel 
excel_filepath = filepath + "/output_file.xlsx"
writer = pd.ExcelWriter(excel_filepath, engine='xlsxwriter')
df_error.to_excel(writer, sheet_name='errors', index = False)
df_message.to_excel(writer, sheet_name='messages', index = False)
df_rule.to_excel(writer, sheet_name='rules', index = False)
writer.save()

