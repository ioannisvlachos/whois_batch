import os
import time
import shutil
import time
import re
import glob
import string
import pandas as pd
import xlsxwriter
import requests


pattern_ip_v4 = '\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}'
pattern_ip_v6 = '\w{4}:\w{4}:\w{4}:\w{4}:\w{4}:\w{4}:\w{4}:\w{4}'


if os.path.exists('isps.xlsx'):
    os.remove('isps.xlsx')

if os.path.exists(".Temp/"):
    shutil.rmtree('.Temp/')
    
#create temporary dir
os.mkdir(os.getcwd() + "\\" + ".Temp")
   
#create xlsx output file
workbook = xlsxwriter.Workbook('isps.xlsx')
worksheet = workbook.add_worksheet()

xls_files = glob.glob('**/*.xls', recursive = True)
xlsx_files = glob.glob('**/*.xlsx', recursive = True)
data = ''
row = 1
col = 0

def document_to_string(path_to_file):
    file_name = os.path.basename(path_to_file) 
    if os.path.splitext(file_name)[1] == ".xlsx":
        xl_file = pd.read_excel(file_name)
        excel = pd.DataFrame(xl_file).to_string()
        return excel
    if os.path.splitext(file_name)[1] == ".xls":
        xl_file = pd.read_excel(file_name)
        excel = pd.DataFrame(xl_file).to_string()
        return excel 
        
def whois_search(ip):
    whois_data = requests.get('https://stat.ripe.net/data/historical-whois/data.json?resource=' + ip)
    string1 = ''   
    pattern = '(?<=netname)(.*)(?=country)'    
    dict_data = whois_data.json()        
    for x in dict_data['data']['objects']:
        for y in x['attributes']:
            for z in y.values():
                string1 = string1 + z
            
    result = re.search(pattern, string1)
    return result.group()
        

        
for x in xlsx_files:        
    data = document_to_string(x)
    result_ip_v4 = re.compile(pattern_ip_v4).findall(data)         
    result_ip_v6 = re.compile(pattern_ip_v6).findall(data) 
    
for ip in result_ip_v4:
    netname = whois_search(ip)
    print('[*] RESULT! IP {} is provided by {}'.format(ip, netname))
    worksheet.write(row, col, ip)
    worksheet.write(row, col + 1, netname)
    row += 1
    #time.sleep(1)
               
for ip in result_ip_v6:
    netname = whois_search(ip)
    print('[*] RESULT! IP {} is provided by {}'.format(ip, whois_search(ip)))
    worksheet.write(row, col, ip)
    worksheet.write(row, col + 1, netname)
    row += 1    
        
    
workbook.close()
shutil.rmtree('.Temp/')
input('Press \'ENTER\' key to exit...')
