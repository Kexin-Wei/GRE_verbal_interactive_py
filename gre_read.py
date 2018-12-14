import xlrd
import operator
import xlwt
from xlutils.copy import copy
import os
import platform


xlsfile=r"6.xls"
xlsfile1=r"xdf.xls"
xlsfile2=r"再要你命3000核心词汇考法精析.xls"
xlsfile3=r"GRE短语乱序.xlsx"
file=[xlsfile,xlsfile1,xlsfile2,xlsfile3]

# choose file
index=int(input(" Choose which one to learn:\n \
1. 6 stufe\n \
2. xdf 6 stufe\n \
3. 3000\n \
4. phrase\n"))-1

#file pyth operation
pypath=os.path.dirname(__file__)
if pypath :
    if platform.system()=="Linux":
        path=os.path.dirname(__file__)+'/'+file[index]
    if platform.system()=="Windows":
        path=os.path.dirname(__file__)+'\\'+file[index]
else:
    path=file[index]

# get data from file
book=xlrd.open_workbook(path,"rb")
sheet0=book.sheet_by_index(0)
sheet1=book.sheet_by_index(1)
nrow=sheet0.nrows

# get explanation from different file
if index !=3:
    explanation=sheet0.col_values(1)
#print(explanation)
if index==1 or index==0:
    explanation2=sheet0.col_values(2)  #2,3 for 6.xls and xdf.xls
    explanation3=sheet0.col_values(3)

#set write excel
bookcp=copy(book)
sheet1w=bookcp.get_sheet(1)

# read last time record
word=sheet0.col_values(0)
num=int(sheet1.cell(0,0).value) # last time stop at round num

#set the list
num_list=list(range(nrow))

#choose continue or start a new turn
start=input("\n###########################\n \
Would you want continue?\n \
(type 'no' for a new turn)\n\
###########################\n")

if operator.eq(start,"no") :
    sheet1w.write(0,0,'0')
    num=0
    print(" Start a new turn now!!")
else:
    print("\n^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^\n \
Continue at : Round ",num+1,"\n\
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^\n")

#choose study mode
# 1 typing and recorrecting mode(default)
# 2 fast view mode
temp=input("\n****************************\n \
Choose the study mode :\n \
1. Typing correct mode(default)\n \
2. Fast view mode\n\
****************************\n")
mode=1
if temp==str(2):
    mode=2
if mode==1 and index==3:
    print(" Sorry for GRE短语乱序.slx there is only a fast view mode\n")
    input(" Type any key to continue :")
    mode=2

#start learning
for x in num_list:
    again=0
    if num and x<num :
        pass
    else:
        print("%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%")
        word[x]=word[x].rstrip()
        print(" Round ",x+1,'/',nrow,'\n')

        #6 stufe and xdf 6 stufe
        if index==0 or index==1:
            print(' ',word[x],"\n ===\n",explanation[x],'\n',explanation2[x],'\n',explanation3[x])
        #3000
        if index==2:
            print(' ',word[x],'\n ===')
            if explanation[x].find('; '):
                #print("Get this")
                for y in explanation[x].split(';'):
                    print(y.lstrip())
            else:
                for y in explanation[x].split('；'):
                    print(y)
        # phrase
        if index==3:
            print(' ',word[x])

        # mode 1 : reprint mode
        if mode==1:
            str=input("\n Please reprint :")

            while operator.ne(str,word[x]) :
                if operator.eq(str,"stop!") :
                    break
                again=1
                print("\n!! Wrong !!\n")
                print(' ',word[x],'\n')
                str=input(" Please reprint :")

            if again:
                str=input("\n Again to testify :")
                while operator.ne(str,word[x]) :
                    if operator.eq(str,"stop!") :
                        break
                    print("\n!! Wrong !!\n")
                    print(' ',word[x],'\n')
                    str=input(" Please reprint :")


            if operator.eq(str,"stop!") :
                break

        # mode 2: fasr view mode
        if mode==2:
            flag=input("\n Waiting.....")
            if operator.eq(flag,"stop!") :
                break
    a=x

if a+1==nrow:
    print(" ## Congradulation!!! List Finished!! ##")
    sheet1w.write(0,0,"0")
else:
    print("\n&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&\n Stop at : Round ",a+1)
    sheet1w.write(0,0,a)
bookcp.save(path)
