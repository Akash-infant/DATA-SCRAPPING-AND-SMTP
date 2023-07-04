import bs4
import requests
import xlsxwriter
import json
import xlrd

class webscarp:
   final_data = []
   read_data=[]

   def data_extract(self,soup):
       data=json.loads(str(soup))
       data=data['data']
       #print(data)
       print("****************************************************")
       print("THE NUMBER OF DATA PRESENT IN THE PAGES IS",len(data))
       for dt in data:
           self.meta_data = [
           dt['symbol'],
           dt['open'],
           dt['dayHigh'],
           dt['dayLow'],
           dt['previousClose'],
           dt['lastPrice'],
           dt['change'],
           dt['pChange'],
           dt['totalTradedVolume'],
           dt['totalTradedValue'],
           dt['nearWKH'],
           dt['nearWKL']
           ]
           #print(self.meta_data)
           self.final_data.append(self.meta_data)
       #print(self.final_data)

   def excel_data(self):
       workbook=xlsxwriter.Workbook("NSE_data.xlsx")
       worksheet=workbook.add_worksheet()
       bold=workbook.add_format({'bold':True})
       worksheet.write('A1','SYMBOL',bold)
       worksheet.write('B1', 'OPEN', bold)
       worksheet.write('C1', 'HIGH', bold)
       worksheet.write('D1', 'LOW', bold)
       worksheet.write('E1', 'PREV_CLOSE', bold)
       worksheet.write('F1', 'LAST_PRICE', bold)
       worksheet.write('G1', 'CHANGE', bold)
       worksheet.write('H1','PERC_CHANGE', bold)
       worksheet.write('I1', 'VOLUME(shares)', bold)
       worksheet.write('J1','VALUE(₹lakhs)', bold)
       worksheet.write('K1','52 W_H',bold)
       worksheet.write('L1','52 W_L', bold)
       print("**************************************")
       print("EXCEL SHEET CREATED SUCESSFULLY......")
       row=1
       col=0
       for data in  self.final_data:
           worksheet.write(row,col,data[0])
           worksheet.write(row, col+1, data[1])
           worksheet.write(row, col+2, data[2])
           worksheet.write(row, col+3, data[3])
           worksheet.write(row, col+4, data[4])
           worksheet.write(row, col+5, data[5])
           worksheet.write(row, col+6, data[6])
           worksheet.write(row, col+7, data[7])
           worksheet.write(row, col+8, data[8])
           worksheet.write(row, col+9, data[9])
           worksheet.write(row, col+10, data[10])
           worksheet.write(row, col+11, data[11])

           row+=1
       print("*************************************")
       print("DATA WRITTEN ON EXCEL SUCESSFULLY...")

       # chart1=workbook.add_chart({'type':'line'})
       # chart1.add_series({'categories':'=Sheet1$B$2:$B$50','values':'Sheet1$A$2:$A$50'})
       # chart1.set_title({'name':'STOCK DATA'})
       # worksheet.insert_chart('T4',chart1)

       chart2 = workbook.add_chart({'type': 'column'})
       chart2.add_series({'categories': '=Sheet1!$A$3:$A$52', 'values': '=Sheet1!$B$2:$B$52'})
       chart2.add_series({'categories': '=Sheet1!$A$3:$A$52', 'values': '=Sheet1!$C$2:$C$52'})
       chart2.set_title({'name': 'STOCK DATA'})
       worksheet.insert_chart('T10',chart2)
       workbook.close()
       print("******************************************************")
       print(" GRAPH OF THE DATA SUCESSFULLY DRAWN ON EXCEL SHEET")

   def read_excel(self):
       wb=xlrd.open_workbook("NSE_data.xlsx")
       worksheet=wb.sheet_by_name("Sheet1")
       num_rows=worksheet.nrows
       num_cols=worksheet.ncols

       for cur_row in range(0,num_rows,1):
           row_review=[]
           for cur_col in range(0,num_cols,1):
               review=worksheet.cell_value(cur_row,cur_col)
               row_review.append(review)
           self.read_data.append(row_review)

   def mail(self):

       import smtplib
       from email.message import EmailMessage

       msg = EmailMessage()
       msg['Subject'] = 'STOCK DETAILS'
       msg['From'] = 'infantakash00@gmail.com'
       msg['To'] = input("ENTER THE MAIL ID THAT DATA SHOULD BE SENT")

       with open("NSE_data.xlsx", "rb") as f:
           file_data = f.read()
       #print("the file data in binary is", file_data)
       file_name = f.name
       print("THE FILE NAME IS  ", file_name)
       msg.add_attachment(file_data, maintype="application", subtype="xlsx", filename=file_name)

       with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
           server.login("124015034@sastra.ac.in", "infant2022")
           server.send_message(msg)
       print("******************************************")
       print("MAIL SENT SUCCESSFULLY")

try:

   print("*********************************")
   print("NSE- NATIONAL STOCK EXCHANGE")
   print("*********************************")
   urllink='https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%2050'
   header={'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36'
   }
   response=requests.get(url=urllink,headers=header)
   soup=bs4.BeautifulSoup(response.content,'html.parser')

   #print(soup)
   w = webscarp()
   w.data_extract(soup)
   print("***************************************")
   print("WEB PAGE EXTRACTED SUCESSFULLY......")
   w.excel_data()
   w.read_excel()
   w.mail()
   print("**********************************************")
   print("DATA SCRAPPED SUCESSFULLY STORED IN EXCEL \nTHANK YOU")
   print("***********************************************")
except Exception as e:
   print("Exception occurs",e,"PLEASE CHECK AND RUN AGAIN....")

