# -*- coding: utf-8 -*-
"""
Created on Mon Sep  5 21:41:41 2022

爬取当日可转债数据_完整功能版第3.8版
@author: xhd

新增功能：增加当日最高价与开盘价之差，如果大于0.05，则置为1；增加收盘价与开盘价之差
        增加将flag为1的写入另一个文件夹，方便导入
        增加今日数据，最小数据，最大数据的文件夹
        3.7版  形成日期子文件夹，并在子文件夹中写入数据
        3.8版  增加预估剩余时间

注意：如果有新增数据列要写入，千万不要写入原txt，应另起
"""



##########################################################


# from selenium import webdriver   
# from selenium.webdriver.edge.options import Options   
import re   #导入正则工具包
#不用看 import json   
import datetime   #导入日期工具包
import os   #导入操作系统包，用来打开文件夹等操作
import xlrd   #导入excel操作包
from msedge.selenium_tools import EdgeOptions   #导入设置msedge浏览器工具包
from msedge.selenium_tools import Edge   #导入操作msedge浏览器工具包

import time   #导入时间工具包




def read_excel(excel_path, sheet_num=0):#从excel中获得可转债代码和可转债名称
    """
    :param excel_path:  xls/xlsx 路径
    :param sheet_num:   sheet下标，默认为0,即第一个sheet页
    :return:
        
    传入要获取的可转债信息表格的绝对路径，
    通过return的形式返回names、numbers
    """

    # 判断文件是否存在
    if os.path.exists(excel_path):
        # 打开excel文件，获得句柄
        excel_handle = xlrd.open_workbook(excel_path)
        
        # 获取第一个工作表(就是excel底部的sheet)
        sheet = excel_handle.sheet_by_index(sheet_num)
        
        # nrows 返回该工作表有效行数
        names = sheet.col_values(2)[1:]   #可转债名称
        numbers = sheet.col_values(1)[1:]   #可转债代码,（2）[1]是打开excel看了所需要的数据在哪的，可根据实际情况调整

    else:
        raise FileNotFoundError("文件不存在")
        
    return names,numbers
 
def get_today():

    time = list(str(datetime.date.today()))   #可以试一下去掉list、str等命令的形式，出来的格式不是所需格式
    for i in range(1, len(time)):
        time[0] = time[0] + time[i]
    # time[0] = time[0][:8] + '16'    #修改为特定日期
    date = time[0]
    return date  #返回当天日期
    
    
        
def make_dirs():            #创建子文件夹，日期形式
    path1 = 'E:\\kezhuanzhai\\data_today\\' + date
    path2 = 'E:\\kezhuanzhai\\data_max\\' + date
    path3 = 'E:\\kezhuanzhai\\data_min\\' + date

    os.makedirs(path1)
    os.makedirs(path2)
    os.makedirs(path3)

    #得到以当天日期命名的文件夹，方便数据存储
    




        
if __name__ == '__main__':  #程序自动执行下面的语句
    names,numbers = read_excel(r"C:\Users\dell\.spyder-py3\原始数据7.10 - 副本.xlsx")
    date = get_today()
    # make_dirs()
    

        

i = 200; time_all = 0    
while i < len(names):

    


    
    
    #爬取股票信息，这一部分感觉可以放在循环外，我也不知道为什么当时没有放出去
    def get_stock(url):   #URL为对应网址

        
        edge_options = EdgeOptions()
        edge_options.use_chromium = True
        
        # 设置无界面模式，也可以添加其它设置
        edge_options.add_argument('headless')
        driver = Edge(options=edge_options)  #driver以options的参数进行
        
        #不用看 msedgedriver.exe下载后放到浏览器的目录下
        #不用看 driver = webdriver.Edge("C:\Program Files (x86)\Microsoft\Edge\Application\msedgedriver.exe")   #告诉python启动edge的路径

        driver.get(url)   # 你要进入的页面
        # print(driver.page_source)
        time.sleep(1)  #暂定1s，等待页面加载完毕
    
        pattern1 = re.compile('<label>最新</label>.*?>(.*?)</span>', re.S)#获取该股票当日收盘价的正则表达式
        C_price = re.findall(pattern1, driver.page_source)#得到该股票当日收盘价，pattern1为匹配规则，driver.page_source为页面资源
        #print("收盘价:", C_price[0])
        
        pattern2 = re.compile('<label>成交量</label>.*?>(.*?)</span>', re.S)#获取该股票当日成交量的正则表达式
        turnover = re.findall(pattern2, driver.page_source)#得到该股票当日成交量
        #print("成交量:", turnover[0])
    
        pattern3 = re.compile('<label>今开</label>.*?>(.*?)</span>', re.S)#获取该股票当日开盘价的正则表达式
        O_price = re.findall(pattern3, driver.page_source)#得到该股票当日开盘价
        #print("开盘价", O_price[0])
        
        pattern4 = re.compile('<label>成交额</label>.*?>(.*?)</span>', re.S)#获取该股票当日成交额的正则表达式
        business = re.findall(pattern4, driver.page_source)#得到该股票当日成交额
        # print("成交额:", business[0])
    
        pattern5 = re.compile('<label>最高</label>.*?>(.*?)</span>', re.S)#获取该股票当日最高价的正则表达式
        Max_price = re.findall(pattern5, driver.page_source)#得到该股票当日最高价
        # print("最高价:", Max_price[0])
    
        pattern6 = re.compile('<label>最低</label>.*?>(.*?)</span>', re.S)#获取该股票当日最低价的正则表达式
        Min_price = re.findall(pattern6, driver.page_source)#得到该股票当日最低价
        # print("最低价:", Min_price[0])
    
        pattern7 = re.compile('<label>量比</label>.*?>(.*?)</span>', re.S)#获取该股票当日量比的正则表达式
        Volume_ratio = re.findall(pattern7, driver.page_source)#得到该股票当日量比
        # print("量比:", Volume_ratio[0])
    
    
    
        driver.close()  # 关闭浏览器
        a = []
        a.append(date)
        date_list = a
                
        Max_sub_O = [str(round((float(Max_price[0]) - float(O_price[0]))/float(O_price[0]), 4))]  #计算最高价与开盘价之差
       
        C_sub_O = [str(round((float(C_price[0]) - float(O_price[0]))/float(O_price[0]), 4))]  #计算收盘价与开盘价之差
        
        C_sub_Min = round((float(C_price[0]) - float(Min_price[0]))/float(O_price[0]), 4)#计算收盘价与最低价之差
        
        if (float(Max_price[0]) - float(O_price[0]))/float(O_price[0]) >= 0.05:  #满足该条件，置flag为1
            flag = ['1']
        else : flag = ['0']
        

        data = date_list + O_price + C_price + turnover + business + Max_price + Min_price + Volume_ratio + Max_sub_O + C_sub_O + flag
        # data = ["日期", "开盘价", "收盘价", "成交量", "成交额", "最高价", "最低价", "量比", "最高价以开盘价之差", "收盘价与开盘价之差"]
        content = data  
        
        
        #不用看 print(content)
        #不用看 print(date, O_price)
        #不用看 print(Max_sub_O, type(Max_sub_O))
        #不用看 print(['\'' + str(Max_sub_O) + '\''], type(str(Max_sub_O)))

        print("爬取数据成功")
        
        return content, C_sub_Min  #返回数据

    #不用看 def write_to_file(content, names):#将当日数据写入txt文件中    
    #不用看     name = 'C:\\Users\\dell\\.spyder-py3\\data\\' + str(names[i]) + '.txt'
    #不用看     with open(name, 'a', encoding='utf-8') as f:
    #不用看         f.write(json.dumps(content, ensure_ascii=False)+'\n')
    #不用看     print(type(content))
    #不用看     print("写入数据成功")
        
    def write_to_file(content, names):#将当日数据写入原txt文件中data_all
        
        name = 'E:\\kezhuanzhai\\data_all\\' + str(names[i])[:4] + '.txt'   #设置的命名规则
        with open(name, 'a', encoding='utf-8') as f:
            f.write('   '.join(content)+'\n')
            f.close()
        
        print("写入数据成功")

    def write_max_to_file(content, names):#将当日flag为1的数据写入txt文件中
        
        name = 'E:\\kezhuanzhai\\data_max\\' + date + '\\' + str(names[i])[:4] + '.txt'
        with open(name, 'a', encoding='utf-8') as f:
            f.write('   '.join(content)+'\n')
            f.close()
        
        print("波动写入成功")
        
    def write_min_to_file(content, names):#将当日数据写入txt文件中
        
        name = 'E:\\kezhuanzhai\\data_min\\' + date + '\\' + str(names[i])[:4] + '.txt'
        with open(name, 'a', encoding='utf-8') as f:
            f.write('   '.join(content)+'\n')
            f.close()

     
    def write_today_to_file(content, names):#将当日数据写入txt文件中
    
        name = 'E:\\kezhuanzhai\\data_today\\' + date + '\\' + str(names[i])[:4] + '.txt'
        with open(name, 'a', encoding='utf-8') as f:
            f.write('   '.join(content)+'\n')
            f.close()
        

        
    
    
    
    def sh_or_sz(numbers_2):#判断是上交所还是深交所
        if numbers_2 == '1':
            str0 = 'sh'
        else:
            str0 = 'sz'
        return str0
            
    def fix_url(url, str0):#将链接修改成对应转债的链接
        url = url[:36] + str0 + str(numbers[i])[:6] + url[44:]
        return url
        
        
    
    def main():#主函数
        url = "https://quote.eastmoney.com/concept/sz123097.html?from=classic"
        numbers_2 = str(numbers[i])[1]
        
        str0 = sh_or_sz(numbers_2)
        
        
        url = fix_url(url, str0)  #针对不同可转债，生成对应链接

        [content, C_sub_Min] = get_stock(url)

        write_to_file(content, names)
        
        write_today_to_file(content, names)
        
        if content[-1] == '1':
            write_max_to_file(content, names)
            
        if C_sub_Min >0.03:
            write_min_to_file(content, names)

        print("OVER")
        
    time_start = time.time()    #程序开始时间
    main()
    time_end = time.time()      #程序结束时间
    i = i+1
    time_sub = time_end - time_start    
    time_all = time_all + time_sub
    time_left = time_all*len(names)/i - time_all
    print('已完成' + str(i) + '/' + str(len(names)), '进度:' + str((i)*100/len(names)) + '%')
    print('预估剩余时长：', int(time_left), 's')
    print('\n\n\n')

    


