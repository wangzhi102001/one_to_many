import pandas as pd
import os #用到的一个新库
import re


'''
io : string, path object ; excel 路径。 
sheetname : string, int, mixed list of strings/ints, or None, default 0    返回多表使用sheetname=[0,1],若sheetname=None是返回全表    注意：int/string 返回的是dataframe，而none和list返回的是dict of dataframe 
header : int, list of ints, default 0 指定列名行，默认0，即取第一行，数据为列名行以下的数据    若数据不含列名，则设定 header = None 
skiprows : list-like,Rows to skip at the beginning，省略指定行数的数据 
skip_footer : int,default 0, 省略从尾部数的int行数据 
index_col : int, list of ints, default None指定列为索引列，也可以使用u”strings”
names : array-like, default None, 指定列的名字。


'''

 
url=''

 




def urlfmt(url):
    url=re.sub("\"",'',url)
    return url
    

def oneEtoMe(url="./data.xlsx",group='村居',n=0):
    '''url：excel的路径；n：excel文件中列名的行号，第一行：0，第二行：1，默认为第一行；group：依据分组的列名，默认为村居,
    请将要拆分的表放在首页！'''
    #url=urlfmt(url)
    data=pd.read_excel(url,header = n)
    #data['日期']=pd.to_datetime(data['时间']).dt.date
    data_excel=[]
    sheetname=[]
    for x in data.groupby(group):
        data_excel.append(x[1])
        sheetname.append(x[0])
    for i in range(len(sheetname)): #区别在于循环创建多个路径，路径中加入变量工作表名称
        data_excel[i].iloc[:,0:9].to_excel(url[:url.find('\\')+1] + str(sheetname[i]) + ".xlsx")
        #桌面新建了一个data文件夹，将拆分的工作簿输出到这里



def Metoone(url):
    '''op:文件夹路径，拖入窗体即可'''
    #op=urlfmt(op)
    url=url+'\\' 
    name_list=os.listdir(op) #用os库获取该文件夹下的文件名称
    data=[] 
    for x in range(len(name_list)):
        df=pd.read_excel(url+name_list[x])#循环导入多个excel文件
        data.append(df)#将每个excel写入到data变量中
    data=pd.concat(data)#合并data变量，转化成Dataframe
    data.to_excel(url+'sum.xlsx',index=False)#输出合并后的excel
    


def Ms_to_ones(url):
    #url=urlfmt(url)
    zx=pd.ExcelFile(url)#获取工作簿里面的属性
    data=zx.parse(zx.sheet_names) #调用属性中的所有sheet名称并将数传入变量data
    data=pd.concat(data)#合并变量中的所有表组成新的DataFrame
    data.to_excel(url[:url.find('\\')+1]+'new.xlsx',index=False)#输出excel文件到桌面，不展示索引


def oneE_to_Ms(url,group='村居',n=0):
    #url=urlfmt(url)
    data=pd.read_excel(url,header = n)#导入数据
    
    data_excel=[] #建一个用于存储多个sheet的空集
    sheetname=[] #建议一个用于存储多个sheet名称的空集
    for x in data.groupby(group): #根据日期字段进行分组
        data_excel.append(x[1]) #将拆分的sheet存储到data_excle里面
        sheetname.append(x[0]) #将拆分的sheet名称存储到sheetname里面
    writer=pd.ExcelWriter(url[:url.find('\\')+1]+'new.xlsx')#定义一个最终文件存储的对象，防止覆盖
    for i in range(len(sheetname)):#创建一个循环将多个sheet输出
        data_excel[i].iloc[:,0:9].to_excel(writer,sheet_name=str(sheetname[i]),index=False)
        #循环将多个sheet表中的数据及对应的sheet表名称输出至桌面，并且不展示索引

b=True

while b:
    innum = input('''
    1.将1个excel拆分成多个excel；
    2.将一个文件夹中的多个excel合成一个excel；
    3.将1个excel下的多个工作表合并为一个工作表，并准存到一个新的excel文件；
    4.将1个excel中的一个工作表，拆分成多个工作表，并转存到一个新的excel文件。
    5.输入q退出。
    请输入（1/2/3/4/q）：
    ''')
    if innum == "1":
        geturl()
        group=input("请输入分组依据（列名）")
        n = int(input("请输入列名的行号，在第一行为0，第二行为1，以此类推"))
        oneEtoMe(url,group,n)
        print("转换完成")
    elif innum == "2":
        geturl-()    
        Metoone(url,group,n)
        print("转换完成")
    elif innum == "3":
        geturl()      
        Ms_to_ones(url,group,n)
        print("转换完成")
    elif innum == '4':
        geturl()
        group=input("请输入分组依据（列名）")
        n = int(input("请输入列名的行号，在第一行为0，第二行为1，以此类推"))
        oneE_to_Ms(url,group,n)
        print("转换完成")
    elif innum == 'q':
        break
    else :
        print("输入错误，请重新输入")










