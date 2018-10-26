import xlrd
#import excel_to_json
import json
import os

from collections import OrderedDict

appent_dir= '/home/mpigelati/PycharmProjects/Srinivas/my_test_folder'
dir_list=[]

def dir_info():
    for root,dir,file in os.walk(appent_dir,topdown=False):
        print("my_dir",dir)
    #print("dir_info",dir)
    return dir

def file_update(name,number,file,string):
    print('name->',name, 'number->',number,'file->', file,'string->',string)
    #print("hellow")
    #

    for dir_check in dir_list:
        #print('dir_check',type(dir_check))
        #print('number',type(number))
        if number ==  dir_check:
            #print("pass")
            print("join ",os.path.join(appent_dir,number,file))
            file = os.path.join(appent_dir, number, file)
            with open(file,"a") as fd:
                fd.write(string)
       #print(dir_check)


def data_print():
    with open("data.json", 'r')as fd:
        JS_Data = json.load(fd, )
        for cars in JS_Data:
            #print(cars['Name'])
            #print(cars['Number'])
            #print(cars['Access_ID'])
            #print(cars['Secreat_key'])
            file_update(cars['Name'],cars['Number'],cars['File'],cars['String'])
            print("\n\n")

def excel_2_json():
    wb=xlrd.open_workbook("/home/mpigelati/PycharmProjects/Srinivas/sample.xlsx")
    sh=wb.sheet_by_index(0)
    # List to hold dictionaries
    cars_list = []
    # Iterate through each row in worksheet and fetch values into dict
    for rownum in range(1, sh.nrows):
     cars = OrderedDict()
     row_values = sh.row_values(rownum)
     cars['Name'] = row_values[0]
     cars['Number'] = str(int(row_values[1]))
     cars['File'] = row_values[2]
     cars['String'] = row_values[3]
     #cars['string'] = row_values[4]
     cars_list.append(cars)

    # Serialize the list of dicts to JSON
    j = json.dumps(cars_list)
    # Write to file
    with open('data.json', 'w') as fd:
        fd.write(j)

if __name__ == '__main__':

    dir_list = dir_info()
    excel_2_json()
    data_print()
    #dir_list = dir_info()
    #print(dir_list)
    #for line in dir_list:
     #   print("temp",line)