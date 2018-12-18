#-*- coding: utf-8 -*-
import xlrd
import operator
import xlwt
from xlutils.copy import copy
import os
import platform
import re
#********************************************************************************
#**************** replace spanish *****************
def inputjudge( outprint, x, index, wordlist,explanation):
    # input sepcial word progress
    in=input(outprint)
    if wordlist[x].find('\u00E9')>=0:
        #print(operator.eq(input,word),1,input.replace("e/",'\u00E9'))
        operator.eq(input.replace("e/",'\u00E9'),word)
    if wordlist[x].find('\u00E8')>=0:
        #print(operator.eq(input,word),2,input.replace("e\\",'\u00E8'))
        operator.eq(input.replace("e\\",'\u00E8'),word)
    if wordlist[x].find('\u00EF')>=0:
        operator.eq(input.replace("i..",'\u00EF'),word)

    if in.find(":f ")>=0:#close function
        sword=in.split(' ',1)[1]
        #print("here")
        findw(sword,index,wordlist,explanation)
        #need to input judge agian
        return inputjudge("\n Please reinput again: ",x,index,wordlist,explanation)
    elif input.find(":s ")>=0:
        return ":s"
    else:
    #print(operator.eq(input,word),3)
        return operator.eq(input,wordlist[x])

#****************** find the word *****************
def findw(sword,index,wordlist,explanation):
    for aim in wordlist:
        aim2=aim.rstrip()
        # special word progress
        if sword.find('*')>=0:
            search=r""+sword.replace('*',"\*")
        else:
            search=r""+sword
        if re.search(search,aim2):
            x=wordlist.index(aim)
            printwe(x,index,wordlist,explanation)
            flagg=1
    if flagg==0:
        print("No such a word")
    return

#******************* print ************************
def printwe(x,index,wordlist,explanation):
    #nexplana=len(explana)
    #****6 stufe and xdf 6 stufe******
    if index==0 or index==1:
        print(' ',wordlist[x])
        print(explanation[1][x],'\n',explanation[2][x],'\n',explanation[3][x])
    #************* 3000*************
    if index==2:
        input(' \u3010'+wordlist[x]+'\u3011')
        z=1
        if explanation[1][x].find('\u003b')>0:
            #print("Get this")
            for y in explanation[1][x].split(';'):
                print(' ',z,'.'+y.lstrip())
                z=z+1
        else:
            for y in explanation[1][x].split('\uff1b'):
                print(' ',z,'.'+y)
                z=z+1
    #*********** phrase**************
    if index==3:
        print(' ',wordlist[x])
    return
#**********************************************************************************
#**********************************************************************************

xlsfile=r"6.xls"
xlsfile1=r"xdf.xls"
xlsfile2=r"再要你命3000核心词汇考法精析.xls"
xlsfile3=r"GRE短语乱序.xlsx"
file=[xlsfile,xlsfile1,xlsfile2,xlsfile3]


#****************** choose file********************
index=int(input(" Choose which one to learn:\n \
1. 6 stufe\n \
2. xdf 6 stufe\n \
3. 3000\n \
4. phrase\n"))-1

#************* file pyth operation********************
pypath=os.path.dirname(__file__)
if pypath :
    if platform.system()=="Linux":
        path=os.path.dirname(__file__)+'/'+file[index]
    if platform.system()=="Windows":
        path=os.path.dirname(__file__)+'\\'+file[index]
else:
    path=file[index]

#******************** get data from file******************
book=xlrd.open_workbook(path,"rb")
sheet0=book.sheet_by_index(0)
sheet1=book.sheet_by_index(1)
nrow=sheet0.nrows

#****************** get explanation from different file *****
explanation=[[]]
if index !=3:
    #explanation1=sheet0.col_values(1) #len=2[0,1]
    explanation.append(sheet0.col_values(1))
#print(explanation)
if index==1 or index==0:
    #explanation2=sheet0.col_values(2)  #2,3 for 6.xls and xdf.xls
    #explanation3=sheet0.col_values(3)  #len=4[0,1,2,3]
    explanation.append(sheet0.col_values(2))
    explanation.append(sheet0.col_values(3))

#************* set write excel********************
bookcp=copy(book)
sheet1w=bookcp.get_sheet(1)

#************ read last time record********************
wordlist=sheet0.col_values(0)
num=int(sheet1.cell(0,0).value) # last time stop at round num

#*********** set the list************************
num_list=list(range(nrow))

#********************************************************************************
#******************* choose study mode*******************
# 1 typing and recorrecting mode(default)
# 2 fast view mode
temp=input("\n************************************\n \
Choose the study mode :\n \
1. Typing correct mode(default)\n \
2. Fast view mode\n\
************************************\n")
mode=1
if temp==str(2):
    mode=2
if mode==1 and index==3:
    print(" Sorry for GRE短语乱序.slx there is only a fast view mode\n")
    input(" Type any key to continue :")
    mode=2

#******************************************************************************
#************* choose continue or start a new turn*****************
start=input("\n###################################\n \
Would you want continue?\n \
(type 'no' for a new turn)\n\
###################################\n")

if operator.eq(start,"no") :
    sheet1w.write(0,0,'0')
    num=0
    print(" Start a new turn now!!")
else:
    print("\n%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%\n \
Continue at : Round ",num+1,"\n\
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%\n")

#for pause after 5 words
pause=0
review=""
#***************** start learning ****************
for x in num_list:
    again=0
    a=x-1
    if num and x<num :
        pass
    else:
#********* rest********************
        if pause > 4:
            sheet1w.write(0,0,a)
            bookcp.save(path)
            print("=======\n Maybe a rest and review?")
            for y in review.split('|'):
                stop=inputjudge("  "+y,x,index,wordlist,explanation)
                if operator.eq(stop,":s"):
                    break
            if operator.eq(stop,":s"):
                break
            review=""
            pause=0
        wordlist[x]=wordlist[x].rstrip()
        if pause==4 :
            review=review+wordlist[x]
        else:
            review= review+wordlist[x]+'|'
        print("\n Round ",x+1,'/',nrow)
        print("&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&")
#***************** file ****************************
        printwe(x,index,wordlist,explanation)
#***************** mode *********************************
        # mode 1 : reprint mode
        if mode==1:
            ToF=inputjudge(" Please reprint :", x, index, wordlist,explanation)
            if operator.eq(ToF,":s") :
                break
            while ToF:
                while ToF:
                    print(" !!! Wrong !!!")
                    print(' ',wordlist[x])
                    ToF=inputjudge(" Please reprint :", x, index, wordlist,explanation)
                if operator.eq(ToF,":s") :
                    break
                ToF=inputjudge(" Again to testify :", x, index, wordlist,explanation)
            if operator.eq(ToF,":s") :
                break

        # mode 2: fasr view mode
        if mode==2:
            flag=inputjudge(" Waiting.....", x, index, wordlist,explanation)
            if operator.eq(flag,":s") :
                break
        pause=pause+1


if a+1==nrow:
    print(" ## Congradulation!!! List Finished!! ##")
    sheet1w.write(0,0,"0")
else:
    print("\n&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&\n Stop at : Round ",a+1)
    sheet1w.write(0,0,a)
bookcp.save(path)
