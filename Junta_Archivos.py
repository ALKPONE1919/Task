import glob

import xlrd

import csv

 

wf = csv.writer(open('csvout_temp.csv', 'w', encoding="utf-8"))

wfl = csv.writer(open('csvout_log.csv', 'w', encoding="utf-8"))

 

with open('path.csv',newline='') as p:

    #xlfiles = csv.reader(p)

    #xlfiles = csv.reader(f)

    #print(xlfiles)

 

    for xlf in p:       

        xlf = xlf.replace('\r\n', '')       

        #print(xlf)

       

        xlf +=  '*.xlsx'

      

        unit = xlf[:xlf.find(',')]

        direc = xlf[xlf.find(',')+1:]

 

        print('xlf: ' + xlf)

        #print('unit:' + unit)

        #print('direc:' + direc)

       

        #wf = csv.writer(open('csvout_temp.csv', 'w', encoding="utf-8"))

 

        for files in glob.iglob(direc):

            #print(files)                   

            if "~$" not in files:               

                workbook = xlrd.open_workbook(files)

                sheet = workbook.sheet_by_index(0)

                try:

                    for row in range(4,sheet.nrows):

                        if sheet.cell(4, 1) != "":

                                wf.writerow([unit]+sheet.row_values(row))

                    wfl.writerow([files]+['Passed'])

                except:

                    wfl.writerow([files]+['Failed'])

                           

with open('csvout_temp.csv',encoding="utf-8",newline='') as fin:

    with open('csvout_final.csv','w',encoding="utf-8",newline='') as fout:

        r = csv.reader(fin)

        w = csv.writer(fout)

        for row in r:               

            if row != []:

                row = [col.replace('\r\n','') for col in row]

                w.writerow(row)

 
