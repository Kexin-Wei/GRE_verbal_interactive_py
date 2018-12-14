import xlrd
import random
import operator
import xlwt
from xlutils.copy import copy
import os

xlsfile=r"verbal/6.xls"
xlsfile1=r"verbal/xdf.xls"
xlsfile2=r"verbal/再要你命3000核心词汇考法精析.xls"
xlsfile3=r"verbal/GRE短语乱序.xlsx"
file=[xlsfile,xlsfile1,xlsfile2,xlsfile3]

index=int(input("Choose which one to learn:\n1. 6 stufe\n2. xdf 6 stufe\n3. 3000\n4. phrase\n"))-1
pypath=os.path.dirname(__file__)
if pypath :
    path=os.path.dirname(__file__)+'/'+file[index]
else:
    path=file[index]

book=xlrd.open_workbook(path,"rb")
sheet0=book.sheet_by_index(0)
sheet1=book.sheet_by_index(1)
nrow=sheet0.nrows
#print(nrow)

bookcp=copy(book)
sheet1w=bookcp.get_sheet(1)



word=sheet0.col_values(0)
if index !=3:
    explanation=sheet0.col_values(1)
#print(explanation)
if index==1 or index==0:
    explanation2=sheet0.col_values(2)  #2,3 for xdf.xls
    explanation3=sheet0.col_values(3)

num_list=list(range(nrow))
#random.shuffle(num_list) #uncomment except 6.xlsx

num=int(sheet1.cell(0,0).value)


if num!=0 :
    print("\n^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^\n  Continue at : Round ",num+1,"\n^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^\n")
    if num in num_list:
        a=num
        #print(a)

for x in num_list:
    if num and x<num :
        pass
    else:
        #************** 6 stufe*************************
        print("%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%")
        word[x]=word[x].rstrip()
        print(" Round ",x+1,'/',nrow,'\n')
        if index==0:
            #print modify mode word,exp->word
            print(' ',word[x],"\n ===\n",explanation[x],'\n',explanation2[x],'\n',explanation3[x])
            str=input("\n Please reprint :")

            # uncomment for fast mode
            if operator.eq(str,"stop!") :
                break
            # end uncomment

            while operator.ne(str,word[x]) :
                print("\n!! Wrong !!\n")
                print(' ',word[x],'\n')
                str=input(" Please reprint :")
                if operator.eq(str,"stop!") :
                    break

            """#comment for fast mode
            flag=input("Continue ? (y/n)")
            if operator.eq(flag,"n") :
                break
            """

        else:#menmorized mode
            #*************** xdf ************************
            if index==1:
                print(' ',word[x],"\n ===\n",explanation[x],'\n',explanation2[x],'\n',explanation3[x])

            #***************** 3000 *************************
            if index==2:
                print(' ',word[x],'\n ===')
                for y in explanation[x].split('；'):
                    print(y)

            #************** phrase ****************
            if index==3:
                print(' ',word[x])

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
