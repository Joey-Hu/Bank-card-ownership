import urllib, sys
from urllib import request
import ssl
import pandas as pd
import json
import xlwt
import time
import random


# 读取excel
data = pd.read_excel(r'D:/software/Desktop/银行卡/not_match_7.xls', usecols=[0])
# data = pd.read_csv(r'test1.csv')
data_li = data.values.tolist()
data = []
for s_li in data_li:
    data.append(s_li[0])
print(data)
print(len(data))

# 创建excel工作表
workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('sheet1')


row = 0
execute_times = 1

for card_num in data:
    print("执行次数"+ str(execute_times))
    # 暂停10秒
    time.sleep(2)
    querys = 'bankcard=' + str(card_num)
    print(querys)
    # 查询
    host = 'https://jisuyhkgsd.market.alicloudapi.com'
    path = '/bankcard/query'
    method = 'ANY'
    appcode = '8923c88261eb4bc688a3fb60522b0d48'
    bodys = {}
    url = host + path + '?' + querys
    # 添加headers
    headers = {}
    user_agent_list = [
        "Opera/8.0 (Windows NT 5.1; U; en)",
        "Mozilla/5.0 (Windows NT 5.1; U; en; rv:1.8.1) Gecko/20061208 Firefox/2.0.0 Opera 9.50",
        "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; en) Opera 9.50",
        "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:34.0) Gecko/20100101 Firefox/34.0",
        ]
    headers['User-Agent'] = random.choice(user_agent_list)
    # 添加异常处理
    try:
        req = urllib.request.Request(url, headers=headers)
        req.add_header('Authorization', 'APPCODE ' + appcode)
        # //根据API的要求，定义相对应的Content-Type
        req.add_header('Content-Type', 'application/json; charset=UTF-8')
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        response = request.urlopen(req, context=ctx)
        content = response.read().decode()
        # 获取到json
        # 解析json
        if (content):
            content1 = json.loads(content)
            card_info = content1['result']
            print(card_info)
            for key, value in card_info.items():
                if key == "bankcard":
                    worksheet.write(row, 0, value)
                elif key == "name":
                    worksheet.write(row, 1, value)
                elif key == "province":
                    worksheet.write(row, 2, value)
                elif key == "city":
                    worksheet.write(row, 3, value)
                elif key == "type":
                    worksheet.write(row, 4, value)
                elif key == "len":
                    worksheet.write(row, 5, value)
                elif key == "bank":
                    worksheet.write(row, 6, value)
                elif key == "bankno":
                    worksheet.write(row, 7, value)
                elif key == "logo":
                    worksheet.write(row, 8, value)
                elif key == "tel":
                    worksheet.write(row, 9, value)
                elif key == "website":
                    worksheet.write(row, 10, value)
                elif key == "iscorrect":
                    worksheet.write(row, 11, value)
        row += 1
    except Exception as err:
        print(err)
    finally:
        execute_times += 1
workbook.save('D:/software/Desktop/银行卡/not_match_search_result_7.xls')




