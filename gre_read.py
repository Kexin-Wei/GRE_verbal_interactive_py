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
def comp( input, word, wordlist):
    if word.find('\u00E9')>=0:
        #print(operator.eq(input,word),1,input.replace("e/",'\u00E9'))
        return operator.eq(input.replace("e/",'\u00E9'),word)
    if word.find('\u00E8')>=0:
        #print(operator.eq(input,word),2,input.replace("e\\",'\u00E8'))
        return operator.eq(input.replace("e\\",'\u00E8'),word)
    if word.find('\u00EF')>=0:
        return operator.eq(input.replace("i..",'\u00EF'),word)
    if input.find(":find ")>=0 and wordlist!=0:#close function
        search=input.split(' ',1)[1]
        #print("here")
        x=findw(search,wordlist)
    #print(operator.eq(input,word),3)
    return operator.eq(input,word)

#******************find the word *****************
def findw(sword,words):
    for aim in words:
        aim2=aim.rstrip()
        if sword.find('*')>=0:
            search=r""+sword.replace('*',"\*")
        else:
            search=r""+sword
        if re.search(search,aim2):
            x=words.index(aim)
            print(" "+aim)
            flagg=1
    if flagg==0:
        print("No such a word")
    return
#**********************************************************************************
#**********************************************************************************

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
    explanation1=sheet0.col_values(1)
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

#********************* choose study mode******************
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
#start learning
for x in num_list:
    again=0
    if num and x<num :
        pass
    else:
#********* rest********************
        # rest
        if pause > 4:
            sheet1w.write(0,0,a)
            bookcp.save(path)
            print("=======\n Maybe a rest and review?")
            for y in review.split('|'):
                stop=input("  "+y)
                if comp(stop,"stop!",0):
                    break
            if comp(stop,"stop!",0):
                break
            review=""
            pause=0
        word[x]=word[x].rstrip()
        if pause==4 :
            review=review+word[x]
        else:
            review= review+word[x]+'|'
        print("\n Round ",x+1,'/',nrow)
        print("&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&")
#***************** file ****************************
        #****6 stufe and xdf 6 stufe******
        if index==0 or index==1:
            print(' ',word[x],"\n",explanation1[x],'\n',explanation2[x],'\n',explanation3[x])
        #************* 3000*************
        if index==2:
            input(' \u3010'+word[x]+'\u3011')
            z=1
            if explanation1[x].find('\u003b')>0:
                #print("Get this")
                for y in explanation1[x].split(';'):
                    print(' ',z,'.'+y.lstrip())
                    z=z+1
            else:
                for y in explanation1[x].split('\uff1b'):
                    print(' ',z,'.'+y)
                    z=z+1
        #*********** phrase**************
        if index==3:
            print(' ',word[x])
#***************** mode *********************************
        # mode 1 : reprint mode
        if mode==1:
            str=input(" Please reprint :")

            #print(comp(str,word[x]),0)

            while comp(str,word[x],word)==0:
                while comp(str,word[x],0)==0:
                    if comp(str,"stop!",0) :
                        break
                    print(" !!! Wrong !!!")
                    print(' ',word[x])
                    str=input(" Please reprint :")
                if comp(str,"stop!",0) :
                    break
                str=input(" Again to testify :")
            if comp(str,"stop!",0) :
                break

        # mode 2: fasr view mode
        if mode==2:
            flag=input("\n Waiting.....")
            if comp(flag,"stop!",word) :
                break
        pause=pause+1
    a=x


if a+1==nrow:
    print(" ## Congradulation!!! List Finished!! ##")
    sheet1w.write(0,0,"0")
else:
    print("\n&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&\n Stop at : Round ",a+1)
    sheet1w.write(0,0,a)
bookcp.save(path)
