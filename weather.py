import requests
import xlwt
from bs4 import BeautifulSoup
import datetime

def datelist(start, end):
    start_date = datetime.date(*start)
    end_date = datetime.date(*end)
    result = []
    curr_date = start_date
    while curr_date != end_date:
        result.append("%04d-%02d-%02d" % (curr_date.year, curr_date.month, curr_date.day))
        curr_date += datetime.timedelta(1)
    result.append("%04d-%02d-%02d" % (curr_date.year, curr_date.month, curr_date.day))
    return result

def get_html():
    global h
    s = requests.session()
    url = 'http://www.whepb.gov.cn/airSubair_water_lake_infoView/v_listhistroy.jspx?type=0'
    headers = {
        'User-Agent':'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:44.0) Gecko/20100101 Firefox/44.0',
        'Referer':'http://www.whepb.gov.cn/airSubair_water_lake_infoView/v_listhistroy.jspx?type=0',
        'Host':'www.whepb.gov.cn',
        'Cookie':'	JSESSIONID=C7DC84CE4FF191ADB582B6C8D39F0749.tomcat; JSESSIONID=C7DC84CE4FF191ADB582B6C8D39F0749.tomcat; clientlanguage=zh_CN',
        'Connection':'keep-alive',
        'Accept-Language':'zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3',
        'Accept-Encoding':'gzip, deflate',
        'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
    }

    r = s.post(url, data=dt, headers=headers)
    h = r.content.decode('utf-8')

def get_excel():
    book = xlwt.Workbook(encoding = 'utf-8',style_compression=0)
    sheet = book.add_sheet('data',cell_overwrite_ok = True)

    global dt
    j = 0
    for each in datelist((2014, 1, 1), (2014, 1, 31)):
        dt = {'cdateEnd':each,'pageNo1':'1','pageNo2':''}
        get_html()
        soup = BeautifulSoup(h)
        #j = 0
        for tabb in soup.find_all('tr'):
            i=0;
            for tdd in tabb.find_all('td'):
                #print (tdd.get_text()+",",)
                sheet.write(j,i,tdd.get_text())
                i = i+1
            j=j+1
    book.save(r'.\\re'+each+'.xls')
get_excel()
