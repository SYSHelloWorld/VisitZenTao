# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
from bs4 import BeautifulSoup
import requests
import xlwt
import re

def CreateExcel(listdic):
    # 创建新的workbook
    workbook = xlwt.Workbook(encoding= 'ascii')

    # 创建新的sheet表
    worksheet = workbook.add_sheet("禅道问题")
    a = 0
    # 往表格写入内容
    for i in listdic:
        worksheet.write(a,0,i['id'])
        worksheet.write(a,1,i['ProblemCreateTime'])
        worksheet.write(a,2,i['ProblemCreator'])
        worksheet.write(a,3,i['ProblemTitle'])
        worksheet.write(a,4,i['ProblemContent'])
        worksheet.write(a,6,i['ProbblemServity'])
        worksheet.write(a,8,i['ImpactEnvironment'])
        worksheet.write(a,9,i['ProblemType'])
        if i['ProblemType']=="04数据问题":
            worksheet.write(a,10,"是")
        else:
            worksheet.write(a,10,"否")
        if i['ProblemType']=="03程序问题" and (i['ProblemStatus']=="已关闭" or i['ProblemStatus']=="已解决"):
            worksheet.write(a,11,"是")
        elif i['ProblemType']=="03程序问题" and i['ProblemStatus']=="处理中":
            worksheet.write(a,11,i['ProblemStatus'])
        else:
            worksheet.write(a,11,"否")
        a=a+1
    #这里填入想要的excel名字
    workbook.save("上海测试.xls")



# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    global dic
    url = "http://sjzc.syxcx.top/zentao/bug-browse-19--all.html"
    payload = {}
    #填入自己的cookie等数值
    headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'en,zh-CN;q=0.9,zh;q=0.8',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Cookie': 'zentaosid=76cc04f65ae2f071392788dd4c6d023a; device=desktop; theme=default; lang=zh-cn; preProductID=19; bugModule=0; qaBugOrder=id_desc; downloading=1; preBranch=0; tab=qa; windowWidth=494; windowHeight=748; _dd_s=logs=1&id=90d7c360-bfda-49cc-ab5a-5554417fa037&created=1655381144144&expire=1655382508086',
        'Host': 'sjzc.syxcx.top',
        'Referer': 'http://sjzc.syxcx.top/zentao/qa/',
        'Upgrade-Insecure-Requests': '',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.0.0 Safari/537.36'
    }
    response = requests.request("GET", url, headers=headers, data=payload)
    soup = BeautifulSoup(response.text, 'html.parser')  # 开始解析
    links_th = soup.find_all('tr')  # 获取所有的tr标签形成列表
    listdic = []
    for i in links_th[3:]:
        text = (i.find_all("td"))  # 获取tr标签下所有td标签形成列表
        ProblemLink= "http://sjzc.syxcx.top" + text[2].find("a").get('href')
        responseContent = requests.request("GET", ProblemLink, headers=headers, data=payload)
        soupContent = BeautifulSoup(responseContent.text, 'html.parser')
        linksContent = soupContent.find('div', attrs={'class':'detail-content article-content'}).contents
        ProblemContent = ''
        for b in linksContent[0:]:
            ProblemContent=ProblemContent+b.get_text()
        #正则表达式多行匹配，然后删去不想要结果
        Match1= re.compile(r'\[问题提交人\].*\[步骤\]',re.DOTALL)
        nihao1=re.sub(Match1,'',ProblemContent)
        Match2= re.compile(r'\[结果.*\[账号\]',re.DOTALL)
        nihao2 = re.sub(Match2,'',nihao1)
        #形成json文件
        dic = {
                "id": int(text[0].get_text()),  # 解析获取text信息
                "ProbblemServity": text[1].get_text(),
                "ProblemTitle": str(text[2].find("a").get_text()),
                "ProblemType": text[3].get_text(),
                "ProblemCreator": text[5].get_text(),
                "ImpactEnvironment": "电票平台",
                "ProblemStatus": text[4].get_text(),
                "ProblemCreateTime": "2022-" + text[6].get_text(),
                "ProblemContent": nihao2
        }
        listdic.append(dic)
        #print(dic)
    CreateExcel(listdic)







