import paramiko
import time
import requests
import win32com
import xlrd
from win32com.client import Dispatch
from Mypack.Code.Parse_csv import unm,psw

class Libraries:

    workbook = xlrd.open_workbook('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Input\\Test Cases.xlsx','rb')
    global sheet,row_idx,wb_copy
    sheet = workbook.sheet_by_index(0)
    row_idx = range(1,sheet.nrows)
    from xlrd import open_workbook
    from xlutils.copy import copy
    wb = open_workbook('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Input\\Test Cases.xlsx')
    wb_copy = copy(wb)

    def readxl(self):
        for row_idx in range(1, sheet.nrows):  # Iterate through rows
            print('-' * 40)
            print('Row: %s' % row_idx)  # Print row number
            for col_idx in range(0, sheet.ncols):  # Iterate through columns
                global cell_obj
                cellObj = sheet.cell(row_idx, col_idx)  # Get cell object by row, col
                if cellObj.value == "SKIP" or "Skip":
                    self.WriteOutputData(row_idx,12,"Row has been skipped")
                    break
                self.methodexe(cellObj, row_idx)

    def methodexe(self,cellObj,row_idx):
        workbook = xlrd.open_workbook('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Input\\Test Cases.xlsx','rb')
        #global sheet,wb_copy
        sheet = workbook.sheet_by_index(0)
        if cellObj.value == "POST":
            url = sheet.cell_value(row_idx, 4)# print(url)
            tn = int(sheet.cell_value(row_idx, 5))          # print(TN)
            payload = sheet.cell_value(row_idx, 6)        # print(payload)
            headers = {'Content-Type': 'application/json'}
            resp = requests.post(url,data=payload,auth=(unm,psw), verify=False, headers=headers)
            stscode = resp.status_code
            self.ExcelWriteResponse(row_idx, 8, resp.text)

            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            ssh.connect('mpls-ds-test.inet.qwest.net',port=22,username=unm,password='Y72rPgxG')
            print("SSH connection established")
            channel = ssh.invoke_shell()
            workbook = xlrd.open_workbook('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Input\\Test Cases.xlsx','r')
            sheet = workbook.sheet_by_index(0)

            while True:
                if channel.recv_ready():
                    channel_data = str(channel.recv(99999))
                    time.sleep(2)
                    self.cmdoutputwrite(channel_data)
                    if channel_data.__contains__(".tcl"):
                        a = channel_data.split("\\r\\r\\n")
                        b = a[0:14] #14]
                        b.pop(0)
                        b.pop(0)
                        b.pop(1)
                        b.pop(1)
                        b.pop(1)
                        b.pop(4)
                        b.pop(4)
                        #b.pop(5)#print(b)
                        self.WriteOutputData(row_idx, 10,'\n'.join(b))
                else:
                    continue

                if channel_data.__contains__('Last login'):
                    channel.send("cd /var/improv-qatest")
                    channel.send('\n')
                    time.sleep(2)
                elif channel_data.__contains__('/var/improv-qatest'):
                    channel.send(f'grep {tn} improv.log')
                    channel.send('\n')
                    time.sleep(2)
                    #break
                elif channel_data.__contains__(f'-w {tn} -f'):
                    time.sleep(2)
                    with open('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Supporting files\\Result.txt','r')as file:
                        data = (file.read())
                        substr = data[-35:]
                        str1 = substr.split('-Z')
                        str2 = str1[1].split('-S')
                        mainstr = str2[0]
                        channel.send(f'./improv-lookup.tcl {mainstr}')
                        channel.send('\n')
                        time.sleep(2)

                elif channel_data.__contains__('-u'):
                    time.sleep(2)
                    with open('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Supporting files\\Result.txt','r')as file:
                        data = (file.read())
                        substr = data[-53:]
                        str1 = substr.split('-u')
                        str2 = str1[1].split('-w')
                        mainstr = str2[0]
                        channel.send(f'./improv-lookup.tcl {mainstr}')
                        channel.send('\n')
                        time.sleep(2)
                else:
                    #break
                    set1 = sheet.cell_value(row_idx,7)
                    # print(set1)
                    if resp.text.__contains__('SUCCESS'):# and list[set1] == b:
                        self.WriteOutputData(row_idx,11,'PASS')
                    else:
                        self.WriteOutputData(row_idx,11,'FAIL')
                        #self.WriteOutputData(row_idx,10,'Response got failed')
                    break
                    #row_idx,10,'Actual and expected configs did not match')


        elif cellObj.value == 'PUT':
            url = sheet.cell_value(row_idx, 4)
            tn = int(sheet.cell_value(row_idx, 5))
            payload = sheet.cell_value(row_idx, 6)
            headers = {'Content-Type': 'application/json'}
            resp = requests.put(url, data=payload, auth=(unm, psw), verify=False, headers=headers)
            self.ExcelWriteResponse(row_idx,8,resp.text)
            # if resp.text.__contains__('ERROR'):
            #     self.WriteOutputData(row_idx,9,'FAIL')
            #     self.WriteOutputData(row_idx,10,'Response got failed')
            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            ssh.connect('mpls-ds-test.inet.qwest.net',port=22,username=unm,password='Y72rPgxG')
            #print("SSH connection established")
            channel = ssh.invoke_shell( )
            workbook = xlrd.open_workbook('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Input\\Test Cases.xlsx','r')
            sheet = workbook.sheet_by_index(0)

            while True:
                if channel.recv_ready( ):
                    channel_data = str(channel.recv(99999))
                    time.sleep(2)
                    #print(channel_data)
                    self.cmdoutputwrite(channel_data)
                    if channel_data.__contains__(".tcl"):
                        a = channel_data.split("\\r\\r\\n")
                        b = a[ 0:14 ]  # 14]
                        b.pop(0)
                        b.pop(0)
                        b.pop(1)
                        b.pop(1)
                        b.pop(1)
                        b.pop(4)
                        b.pop(4)
                        self.WriteOutputData(row_idx, 10,'\n'.join(b))
                else:
                    continue
                if channel_data.__contains__('Last login'):
                    channel.send("cd /var/improv-qatest")
                    channel.send('\n')
                    time.sleep(2)
                elif channel_data.__contains__('/var/improv-qatest'):
                    channel.send(f'grep {tn} improv.log')
                    channel.send('\n')
                    time.sleep(2)
                elif channel_data.__contains__(f'-w {tn} -f'):
                    time.sleep(2)
                    with open('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Supporting files\\Result.txt','r')as file:
                        data = (file.read())
                        substr = data[-35:]
                        str1 = substr.split('-Z')
                        str2 = str1[1].split('-S')
                        mainstr = str2[0]
                        channel.send(f'./improv-lookup.tcl {mainstr}')
                        channel.send('\n')
                        time.sleep(2)

                elif channel_data.__contains__('-u'):
                    time.sleep(2)
                    with open('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Supporting files\\Result.txt','r')as file:
                        data = (file.read())
                        substr = data[-53:]
                        str1 = substr.split('-u')
                        str2 = str1[1].split('-w')
                        mainstr = str2[0]
                        channel.send(f'./improv-lookup.tcl {mainstr}')
                        channel.send('\n')
                        time.sleep(2)
                else:
                    set1 = sheet.cell_value(row_idx,7)
                    if resp.text.__contains__('SUCCESS'):  # and list[set1] == b:
                        self.WriteOutputData(row_idx, 11,'PASS')
                    else:
                        self.WriteOutputData(row_idx, 11,'FAIL')
                    break


        elif cellObj.value == 'GETbyTN' :#or cellObj.value == 'GETbyUName':
            url = sheet.cell_value(row_idx, 4)
            resp = requests.get(url, auth=(unm, psw), verify=False)
            self.ExcelWriteResponse(row_idx, 8, resp.text)
            if resp.text.__contains__("wtn"):
                self.WriteOutputData(row_idx, 11,'PASS')
            else:
                self.WriteOutputData(row_idx, 11,'FAIL')

        elif cellObj.value == "GETbyUName":
            url = sheet.cell_value(row_idx, 4)
            resp = requests.get(url, auth=(unm, psw), verify=False)
            self.ExcelWriteResponse(row_idx,8,resp.text)

            if resp.text.__contains__("userName"):
                self.WriteOutputData(row_idx, 11,'PASS')
            else:
                self.WriteOutputData(row_idx, 11,'FAIL')

        elif cellObj.value == 'DELETE':
            url = sheet.cell_value(row_idx, 4)
            tn = int(sheet.cell_value(row_idx, 5))
            payload = sheet.cell_value(row_idx, 6)
            headers = {'Content-Type': 'application/json'}
            resp = requests.delete(url, data=payload, auth=(unm, psw), verify=False, headers=headers)
            self.ExcelWriteResponse(row_idx, 8, resp.text)
            # if resp.text.__contains__('ERROR'):
            #     self.WriteOutputData(row_idx, 11,'FAIL')
            #     self.WriteOutputData(row_idx, 12,'Response got failed')

            ssh = paramiko.SSHClient( )
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy( ))
            ssh.connect('mpls-ds-test.inet.qwest.net',port=22,username='ab17764',password='Y72rPgxG')
            #print("SSH connection established")
            channel = ssh.invoke_shell( )
            workbook = xlrd.open_workbook('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Input\\Test Cases.xlsx','r')
            # global sheet,row_idx,wb_copy
            sheet = workbook.sheet_by_index(0)
            # row_idx = range(1,sheet.nrows)
            while True:  # Working code
                if channel.recv_ready( ):
                    # global channel_data
                    channel_data = str(channel.recv(99999))
                    time.sleep(2)
                    #print(channel_data)
                    self.cmdoutputwrite(channel_data)
                    if channel_data.__contains__(".tcl"):
                        a = channel_data.split("\\r\\r\\n")
                        b = a[0:14]  # 14]
                        b.pop(0)                        # print(b)
                        b.pop(0)                        # print(b)
                        b.pop(2)                        # print(b)
                        b.pop(2)                        # print(b)
                        b.pop(5)
                        b.pop(5)    # print(b)
                        self.WriteOutputData(row_idx, 10,'\n'.join(b))

                else:
                    continue
                if channel_data.__contains__('Last login'):
                    channel.send("cd /var/improv-qatest")
                    channel.send('\n')
                    time.sleep(2)
                elif channel_data.__contains__('/var/improv-qatest'):
                    channel.send(f'grep {tn} improv.log')
                    channel.send('\n')
                    time.sleep(2)
                    # break
                elif channel_data.__contains__(f'-w {tn} -f'):
                    # print(tn)
                    time.sleep(2)
                    with open('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Supporting files\\Result.txt','r')as file:
                        data = (file.read())
                        substr = data[-35:]
                        str1 = substr.split('-Z')
                        str2 = str1[1].split('-S')
                        mainstr = str2[0]
                        channel.send(f'./improv-lookup.tcl {mainstr}')
                        channel.send('\n')
                        time.sleep(2)

                elif channel_data.__contains__('-u'):
                    time.sleep(2)
                    with open('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Supporting files\\Result.txt','r')as file:
                        data = (file.read())
                        substr = data[-53:]
                        str1 = substr.split('-u')
                        str2 = str1[1].split('-w')
                        mainstr = str2[0]
                        channel.send(f'./improv-lookup.tcl {mainstr}')
                        channel.send('\n')
                        time.sleep(2)
                else:
                    set1 = sheet.cell_value(row_idx, 10)
                    if resp.text.__contains__("SUCCESS") and set1.__contains__("1:*"):
                        self.WriteOutputData(row_idx, 11,'PASS')
                    else:
                        self.WriteOutputData(row_idx, 11,'FAIL')
                    break

    def ExcelWriteResponse(self, row, col, resp):
        excel = win32com.client.Dispatch("Excel.Application")
        wkbook = excel.Workbooks.Open('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Input\\Test Cases.xlsx')
        sheet = wkbook.Sheets("Sheet1")
        cell = sheet.Cells(1,1)
        if (cell == None):
            for row_idx in (1,sheet.nrows):
                wb_copy.get_sheet(0).write(row_idx, 8, resp)
                # row_idx += 1
                wb_copy.save('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Results\\Result.xls')
            print("empty")
        else:
            wb = xlrd.open_workbook('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Results\\Result.xls')
            wb_copy.get_sheet(0).write(row, col, resp)
            wb_copy.save('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Results\\Result.xls')

    def WriteOutputData(self, row, col, data):
        if not data:
            with open('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Supporting files\\Result.txt','r')as file:
                data = (file.read())
            wb = xlrd.open_workbook('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Results\\Result.xls')
            wb_copy.get_sheet(0).write(row, col, data)
            wb_copy.save('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Results\\Result.xls')
        else:
            wb = xlrd.open_workbook('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Results\\Result.xls')
            wb_copy.get_sheet(0).write(row,col,data)
            wb_copy.save('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Results\\Result.xls')

    def cmdoutputwrite(self,channel_data):
        #print(channel_data)
        file = open('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Supporting files\\Result.txt','w')
        file.write(channel_data)
        file.close()

    def formatcelldata(self,channel_data):
        excel = win32com.client.Dispatch("Excel.Application")
        wkbook = excel.Workbooks.Open('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Results\\Result.xls')
        sheet = wkbook.Sheets("Sheet1")
        wb_copy.save('C:\\Users\\AB17764\\Improv-Auto\\Mypack\\Results\\Result.xls')

#Libraries().readxl()
