import xlrd
import xlwt
import string
files=['1.xls','2.xls','3.xls','4.xlsx','5.xls','6.xls']
start_date=xlrd.xldate.xldate_from_date_tuple((2017,3,13),0)
end_date=xlrd.xldate.xldate_from_date_tuple((2017,5,20),0)
Tmin=0
Tmax=1
Valid=2
Count=1
result={}           #No->[[Tmin,Tmax,Valid],relative_date]
sumr={}             #No->['Name',count]
def check_valid(info):
    tmin=[string.atoi(s) for s in info[Tmin].split(':')]
    tmax=[string.atoi(s) for s in info[Tmax].split(':')]
    if ((tmax[0]-tmin[0])*3600+(tmax[1]-tmin[1])*60+tmax[2]-tmin[2] in range(1800,7200)):
        if (info[Valid]==0):
            sumr[person['No']][Count]=sumr[person['No']][Count]+1
        info[Valid]=1
        return 1
    
for file_name in files:
    data=xlrd.open_workbook(file_name)
    table=data.sheets()[0]
    for i in range(1,table.nrows):
        person={'Name':table.col_values(2)[i],'No':table.col_values(9)[i],'Date':table.col_values(12)[i],'Time':table.col_values(14)[i]} 
        if (person['Date'] in range(int(start_date),int(end_date))):
            if result.has_key(person['No']):
                if (result[person['No']][int(person['Date']-start_date)][Valid]==-1):
                    result[person['No']][int(person['Date']-start_date)][Tmin]=person['Time']
                    result[person['No']][int(person['Date']-start_date)][Tmax]=person['Time']
                    result[person['No']][int(person['Date']-start_date)][Valid]=0
                else:   
                    result[person['No']][int(person['Date']-start_date)][Tmin]=min(result[person['No']][int(person['Date']-start_date)][Tmin],person['Time'])
                    result[person['No']][int(person['Date']-start_date)][Tmax]=max(result[person['No']][int(person['Date']-start_date)][Tmax],person['Time'])
                    result[person['No']][int(person['Date']-start_date)][Valid]=check_valid(result[person['No']][int(person['Date']-start_date)])
            else:
                temp=[]
                for i in range(int(end_date-start_date)):
                    temp.append([0,0,-1])
                temp[int(person['Date']-start_date)]=[person['Time'],person['Time'],0]
                result.update({person['No']:temp})
                sumr.update({person['No']:[person['Name'],0]})
book=xlwt.Workbook()
table=book.add_sheet('YGSK')
table.write(0,0,'StuNo')
table.write(0,1,'Name')
table.write(0,2,'Valid_Count')
i=0
for x in sumr.keys():
    if (sumr[x][1]>0):
        i=i+1
        table.write(i,0,string.atoi(x))
        table.write(i,1,sumr[x][0])
        table.write(i,2,sumr[x][1])
book.save('result.xls')
print 'Done!'
