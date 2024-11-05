import json
from urllib.request import urlopen
import xlwt  # 进行excel操作
import requests
import hashlib
import time
from urllib.parse import quote
import urllib.parse
def makeexcle(nextpage,wts,w_rid):

    url = f"https://api.bilibili.com/x/v2/reply/wbi/main?oid=59846708&type=1&mode=3&pagination_str={nextpage}&plat=1&web_location=1315875&w_rid={w_rid}&wts={wts}"
    print(url)
    head = {  # 模拟浏览器头部信息
        "cookies":"buvid3=8D4A212F-F4BF-84F3-BFA4-36496DE2B25C50712infoc; b_nut=1726194550; _uuid=79392CAF-A85D-FE44-123D-6928566FEDC148924infoc; buvid4=6D388718-B6B0-C74E-53AB-B3939395FDFB52531-024091302-yuS89Dg2%2F0we8pBIChsahQ%3D%3D; header_theme_version=CLOSE; rpdid=|(J~JJlJ~|YJ0J'u~kYkuJm~|; DedeUserID=352397956; DedeUserID__ckMd5=5be7fc519ea0add2; enable_web_push=ENABLE; iflogin_when_web_push=1; buvid_fp_plain=undefined; LIVE_BUVID=AUTO8617269298504603; hit-dyn-v2=1; CURRENT_BLACKGAP=0; fingerprint=22ac2b80e24dab1b14131b46824e84d2; CURRENT_QUALITY=116; CURRENT_FNVAL=4048; buvid_fp=22ac2b80e24dab1b14131b46824e84d2; bili_ticket=eyJhbGciOiJIUzI1NiIsImtpZCI6InMwMyIsInR5cCI6IkpXVCJ9.eyJleHAiOjE3MzA4MjgzNDMsImlhdCI6MTczMDU2OTA4MywicGx0IjotMX0.0fkbkQhEHBIy32ELhR0v2CKHi2PGNULK41lgtb4Df0w; bili_ticket_expires=1730828283; SESSDATA=1c3f85a4%2C1746166760%2Cc3ca0%2Ab2CjCdl7rAQHOf6aLCGfFH37eA6cywypR4cutrwr1fic94bjhlRc3ilywTYPOQLJPcxTgSVjZRRkNUc3YtUmpsdi1aQktVSlN1TFpYUWt4M2MwYkphZTl5X2tDTkNleXRPcmg5cmhwTEhaRnlyMlFyZEJhTnFSX3duUlFoeHpaSDB5cWtoSTdnbkt3IIEC; bili_jct=41904f88d2073cce0c7899c16ffd80e6; PVID=23; bp_t_offset_352397956=996043658002169856; sid=6akwx5fl; home_feed_column=5; browser_resolution=1698-820; b_lsid=7CAA4FF2_192FB23A87D",
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.110 Safari/537.36 Edg/96.0.1054.62'
    }
    resp = requests.get(url=url,headers=head)
    resp.encoding = 'utf-8'
    replys = resp.json()
    print(replys)
    replye = replys['data']['replies']
    Nextpage = nextpage
    print(Nextpage)
    nowtime = int(time.time())
    hasher(NextPage=Nextpage,wts=nowtime)
    print(nextpage)
    # https://bangumi.bilibili.com/sponsor/web_api/v2/rank/week?season_id=1293&season_type=1&page=1&pagesize=100 充电的人
    book = xlwt.Workbook(encoding="utf-8",style_compression=0) #创建workbook对象
    sheet = book.add_sheet('幸运星评论', cell_overwrite_ok=True) #创建工作表
    col = ("昵称","评论","性别","评论时间")

    for i in range(0, 4):
        sheet.write(0, i, col[i])  # 列名
    row_index = 1
    for index in replye:

        dit = {
            '昵称':index['member']['uname'],
            '评论':index['content']['message'],
            '性别': index['member']['sex'],
            '评论时间': index['reply_control']['time_desc'],
        }
        print(dit)
        data = dit
        for j,key in enumerate(col):
            sheet.write(row_index, j, data[key])  # 数据
        row_index += 1
    book.save('幸运星评论.xls')
    return Nextpage
def hasher(NextPage,wts):
    if NextPage == '%7B%22offset%22%3A%22%22%7D':
        en = [
            "mode=3",
            "oid=59846708",
            "pagination_str=%7B%22offset%22%3A%22%22%7D",
            "plat=1",
            "type=1",
            "web_location=1315875",
            "wts=1730805574",
        ]
        Jt = '&'.join(en)
        string = Jt + 'ea1db124af3c7062474693fa704f4ff8'
        MD5 = hashlib.md5()
        MD5.update(string.encode('utf-8'))
        w_rid = MD5.hexdigest()
        print(111)
        print(w_rid)
        return w_rid
    else:
        pagination_str = quote(NextPage)
        NextPage_string = pagination_str.replace('%22', '%5C%22')
        NextPage_string1 = NextPage_string.replace('%7B%5C%22offset%5C%22%3A%5C%22%7B%', '%7B%22offset%22%3A%22%7B%')
        NextPage_string2 = NextPage_string1.replace('%5C%22%7D', '%22%7D')
        print(NextPage_string2)
        en =[
            "mode=3",
            "oid=59846708",
            f"pagination_str={NextPage_string2}",
            "plat=1",
            "type=1",
            "web_location=1315875",
            f"wts={wts}"
        ]
        Jt = '&'.join(en)
        string = Jt + 'ea1db124af3c7062474693fa704f4ff8'
        MD5 = hashlib.md5()
        MD5.update(string.encode('utf-8'))
        w_rid = MD5.hexdigest()
        print(w_rid)
        return w_rid

if __name__ == "__main__":  # 当程序执行时
     for i in range (1,51):
         print(f"正在爬取第{i}页")
         xk = '{"offset":"{\"type\":1,\"direction\":1,\"session_id\":\"1772325578757618\",\"data\":{}}"}'
         wts = int(time.time())
         print(wts)
         pagination_str = quote(xk)
         NextPage_string = pagination_str.replace('%22', '%5C%22')
         NextPage_string1 = NextPage_string.replace('%7B%5C%22offset%5C%22%3A%5C%22%7B%', '%7B%22offset%22%3A%22%7B%')
         NextPage_string2 = NextPage_string1.replace('%5C%22%7D', '%22%7D')
         w_rid = hasher(NextPage=xk,wts=wts)
         makeexcle(nextpage=NextPage_string2,wts=wts,w_rid=w_rid)
    # init_db("movietest.db")
     print("爬取完毕！")
