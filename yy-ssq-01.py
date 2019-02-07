#!/usr/local/bin/python3

#从非凡彩票网站抓取双色球历史开奖数据，写入excel文件 
#2019-02-04，yangying

import requests  
from bs4 import BeautifulSoup  
import xlwt  
import time
from datetime import datetime 
 

#获取第一页的内容  
def get_one_page(url):  
    headers = {  
        'User-Agent':'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.140 Safari/537.36'  
    }  
    response = requests.get(url,headers=headers)  
    if response.status_code == 200:  
        return response.text  
    return None  
 
#解析第一页内容，数据结构化  
def parse_one_page(html):  
 
    soup = BeautifulSoup(html,'lxml')  
    i = 0  
    for item in soup.select('tr')[6:-10]:  
 
        yield{  
            #'time':item.select('td')[i].text,  
            'qihao':item.select('td')[i].text,  
            'hongqiu':item.select('td')[i+1].text,  
            #'qianqu_2':item.select('td em')[1].text,  
            #'qianqu_3':item.select('td em')[2].text, 
            #'qianqu_4':item.select('td em')[3].text, 
            #'qianqu_5':item.select('td em')[4].text, 
            'lanqiu':item.select('td')[i+2].text,  
            #'houqu2':item.select('td em')[1].text 
            #'group_selection_6':item.select('td')[i+5].text,  
            #'sales':item.select('td')[i+6].text,  
            #'return_rates':item.select('td')[i+7].text  
        }  
 
 
#将数据写入Excel表格中  
def write_to_excel():  
    f = xlwt.Workbook()      
    sheet1 = f.add_sheet('results',cell_overwrite_ok=True)
    row0 = ["期号","红球","蓝球"] 
    #写入第一行 
    for j in range(0,len(row0)): 
        sheet1.write(0,j,row0[j]) 
 
    #依次爬取每一页内容的每一期信息，并将其依次写入Excel 
    i=0

    t_year = datetime.now().year

    for k in range(2005,t_year+1):  
        #url = 'http://kaijiang.zhcw.com/zhcw/html/3d/list_%s.html' %(str(k))  
        #if k<2019:      
        url = 'http://www.ffcp.cn/zs/SSQ/%s.html' %('SSQZH') if k == t_year else  'http://www.ffcp.cn/zs/SSQ/%s.html' %('SSQZHY'+str(k))
         #   pass
        #else:
         #   url = 'http://www.ffcp.cn/zs/DLT/DLT.html' 
         #   pass
          
            
        
        html = get_one_page(url)  
        print('正在保存%d年的开奖结果。'%k) 
        #写入每一期的信息  
        for item in parse_one_page(html):  
            sheet1.write(i+1,0,item['qihao'])  
            sheet1.write(i+1,1,item['hongqiu'])  
            sheet1.write(i+1,2,item['lanqiu'])  
            #sheet1.write(i+1,3,item['qianqu_3'])  
            #sheet1.write(i+1,4,item['qianqu_4'])  
            #sheet1.write(i+1,5,item['qianqu_5'])  
            #sheet1.write(i+1,6,item['houqu1'])  
            #sheet1.write(i+1,7,item['houqu2'])  
            #sheet1.write(i+1,8,item['sales'])  
            #sheet1.write(i+1,9,item['return_rates'])  
            i+=1  
 
 
    f.save('ssq-01.xls')  
 
def main():  
    write_to_excel()  
 
if __name__ == '__main__':  
    main()  