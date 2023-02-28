import pandas as pd
import re
from calendar import monthrange
import csv
import os
def total_days_of_month(input_year,input_month):
    num_days = monthrange(input_year, input_month)[1]
    return num_days
def date_preprocessing(input_date):
    input_date = re.sub('\s+',' ',input_date)
    input_date = re.sub('-+','-',input_date)
    input_date = re.sub('[.]+','.',input_date)
    input_date = re.sub('[,]+',',',input_date)
    input_date = re.sub("[']+",' ',input_date)
    input_date = re.sub("[/]+",'/',input_date)
    input_date = re.sub("[:]+",':',input_date)
    input_date = re.sub(r'([a-zA-Z]+)(\d+)', r'\1 \2', input_date)
    input_date = re.sub(r'(\d+)([a-zA-Z]+)', r'\1 \2', input_date)
    return input_date
def preprocess_formating(input_date):
    input_date = re.sub('\s+',' ',input_date)
    input_date = re.sub('-+',' ',input_date)
    input_date = re.sub('[.]+',' ',input_date)
    input_date = re.sub('[,]+',' ',input_date)
    input_date = re.sub("[']+",' ',input_date)
    input_date = re.sub("[/]+",' ',input_date)
    input_date = re.sub("[:]+",' ',input_date)
    input_date = re.sub('\s+',' ',input_date)
    return input_date

def num_there(s):
    return any(i.isdigit() for i in s)
    

# a = date_preprocessing("12.02.2022")
# b = preprocess_formating(a)
# print(num_there(b))
month_name_dict = {'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12}


def make_right_format(val):
    previous_date = val
    updated_date = val
    try: 
        if val:
            val = val.lower()
            if num_there(val):
                val = date_preprocessing(val)
                val = preprocess_formating(val)
                # print(val)
                val_list = val.split(' ')
                temp_date = ''
                if len(val_list)==2:
                    month = ''
                    day = ''
                    year = ''
                    if val_list[0].isdigit() and len(val_list[0])==4:
                        val_list[0],val_list[1] = val_list[1],val_list[0]
                    if val_list[0].isdigit() and int(val_list[0])<=12:
                        month = int(val_list[0])
                    if not val_list[0].isdigit():
                        for x in month_name_dict:
                            if x in val_list[0]:
                                month = month_name_dict[x]
                                break 

                    if val_list[1].isdigit():
                        year = int(val_list[1])
                        if year<100:
                            year = year+2000
                    if year!='' and month!='':
                        day = total_days_of_month(year,month)
                    if day!='' and month!='' and year!='':
                        temp_date = str(day)+'/'+str(month)+'/'+str(year)
                    else:
                        temp_date = 'Wrong Format'
                elif len(val_list)>=3:
                    month = ''
                    day = ''
                    year = ''
                    if val_list[0].isdigit() and len(val_list[0])==4:
                        val_list[0],val_list[2] = val_list[2],val_list[0]
                    if val_list[0].isdigit():
                        day = val_list[0]
                    if val_list[1].isdigit() and int(val_list[1])<=12:
                        month = int(val_list[1])
                    if not val_list[1].isdigit():
                        for x in month_name_dict:
                            if x in val_list[1]:
                                month = month_name_dict[x]
                                break 

                    if val_list[2].isdigit():
                        year = int(val_list[2])
                        if year<100:
                            year = year+2000
                    
                    if day!='' and month!='' and year!='':
                        temp_date = '('+str(day)+'/'+str(month)+'/'+str(year)+')'
                    else:
                        temp_date = 'Wrong Format'
                updated_date = temp_date
                print(updated_date)
    except:
        pass
    return updated_date


if __name__ == "__main__":
    for filename in os.listdir("./input_file"):
        read_file = pd.read_excel ("./input_file/"+filename)
        output_file_name = filename.split(".")[0]
        read_file.to_csv ("test.csv", 
                        index = None,
                        header=True)
        data = pd.read_csv("test.csv", keep_default_na=False)
        header = []
        with open('test.csv') as file_obj:
            reader_obj = csv.reader(file_obj)
            print(filename)
            for row in reader_obj:
                header = row
                break
        shelf_index = None
        date_index = None
        if 'SH./LIFE' in header:
            shelf_index = header.index('SH./LIFE')
        if 'Date' in header:
            date_index = header.index('Date')
        final_list = []
        cnt = 0
        with open('test.csv') as file_obj:
            reader_obj = csv.reader(file_obj)
            for row in reader_obj:
                if shelf_index is not None:
                    self_date = make_right_format(row[shelf_index])
                    row[shelf_index] = self_date
                if date_index is not None:
                    print(cnt)
                    date_val = make_right_format(row[date_index])
                    row[date_index] = date_val
                cnt = cnt+1
                final_list.append(row)
                
        with open('output_csv/'+output_file_name+'.csv','w',newline='') as f:
            writer=csv.writer(f)
            for val in final_list:
                writer.writerow(val)
        read_file = pd.read_csv ('output_csv/'+output_file_name+'.csv')
        read_file.to_excel ('output_xlsx/'+output_file_name+'.xlsx', index = None, header=True)

                
                
        
        
        # sn_no = data['S/N'].tolist()
        # part_number = data['PART NUMBER'].tolist()
        # name = data['NAME'].tolist()
        # shelf_life = data["SH./LIFE"].tolist()
        # date_f = data['Date'].tolist()
        # u_price = data['U/Price(USD)'].tolist()
        # opening_qty = data['OpeningQty'].tolist()
        # opening_val = data['OpeningValue(USD)'].tolist()
        # purchase_qty = data['PurchasedQty'].tolist()
        # purchase_val = data['PurchasedValue(USD)'].tolist()
        # issued_qty = data['IssuedQty'].to_list()
        # issued_val = data['IssuedValue(USD)'].tolist()
        # qty = data['Qty'].tolist()
        # closing_qty = data['ClosingQty'].tolist()
        # closing_val = data['ClosingValue(USD)'].tolist()
        # serial_number = data['Serial Number'].tolist()
        # grn = data['G.R.N'].tolist()
        # alt_pno = data['Alt P/No.'].tolist()
        # en = data['UNIT'].tolist()



        # with open('output_csv/'+output_file_name+'.csv','w',newline='') as f:
        #     writer=csv.writer(f)
        #     writer.writerow(['S/N','PART NUMBER','NAME','UNIT','SH./LIFE','Date','U/Price(USD)','OpeningQty',
        #     'OpeningValue(USD)','PurchasedQty','PurchasedValue(USD)','IssuedQty','IssuedValue(USD)','Qty',
        #     'ClosingQty','ClosingValue(USD)','Serial Number','G.R.N','Alt P/No.'])
        #     cnt = 0
        #     for i in range(0,len(sn_no)):
        #         shelf_date = make_right_format(shelf_life[i])
        #         date_val = make_right_format(date_f[i])
        #         writer.writerow([sn_no[i],part_number[i],name[i],en[i],shelf_date,date_val,u_price[i],opening_qty[i],opening_val[i]
        #         ,purchase_qty[i],purchase_val[i],issued_qty[i],issued_val[i],qty[i],closing_qty[i],closing_val[i],serial_number[i],
        #         grn[i],alt_pno[i]])
        # read_file = pd.read_csv ('output_csv/'+output_file_name+'.csv')
        # read_file.to_excel ('output_xlsx/'+output_file_name+'.xlsx', index = None, header=True)
    
                