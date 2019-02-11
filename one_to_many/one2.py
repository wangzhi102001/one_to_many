import numpy
import pandas as pd
import re
import os

numpy.set_printoptions(suppress=True, threshold=np.nan)
#suppress=True 取消科学记数法
#threshold=np.nan 完整输出（没有省略号）

    #将一个文件拆分成多个XLSX文件；
   # url：excel的路径；n：excel文件中列名的行号，第一行：0，第二行：1，默认为第一行；group：依据分组的列名，默认为村居,
    #请将要拆分的表放在首页！'''
    #url=urlfmt(url)
n=0
group = "村居"
path = "./xlsx/"
if not os.path.exists(path):
        # 如果不存在则创建目录
        # 创建目录操作函数
    os.makedirs(path)  
    print (path+' 创建成功')    
else:
    # 如果目录存在则不创建，并提示目录已存在
    print (path+' 目录已存在')


data=pd.read_excel("./data.xlsx"or "./data.xls",header = n)
#data['日期']=pd.to_datetime(data['时间']).dt.date
data_excel=[]
sheetname=[]
for x in data.groupby(group):
    data_excel.append(x[1])
    sheetname.append(x[0])
for i in range(len(sheetname)): #区别在于循环创建多个路径，路径中加入变量工作表名称
    data_excel[i].iloc[:,0:30].to_excel("./xlsx/" + str(sheetname[i]) + ".xlsx")
    #桌面新建了一个data文件夹，将拆分的工作簿输出到这里
