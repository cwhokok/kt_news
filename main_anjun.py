import requests
import datetime
import pandas as pd
import jellyfish
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os


# 초기 Data
df_email = pd.read_csv(os.path.dirname(os.path.realpath(__file__))+'/email_anjun.csv') ## 담당자-연락처(email)
dic_email = df_email.set_index('name').T.to_dict()
df = pd.read_csv(os.path.dirname(os.path.realpath(__file__))+'/db_anjun.csv') ## 담당자-키워드
dic = df.groupby('name')['keyword'].apply(list).to_dict() ## 데이터 >> {이름:[키워드...]}

## Naver Api 초기설정
url = 'https://openapi.naver.com/v1/search/news.json'
headers_list = [{'X-Naver-Client-Id':'saJIehvVEaCgdJFuqgKI', 'X-Naver-Client-Secret':'THAztdVvGa'},{'X-Naver-Client-Id':'hBNIfPqLrB1fvHX9YxYl', 'X-Naver-Client-Secret':'6j_tM2cxbx'},{'X-Naver-Client-Id':'cFTQk9WP6rmMclGsq_RF', 'X-Naver-Client-Secret':'pqmgbuh7ff'},{'X-Naver-Client-Id':'KNsdR7f7MhQmKzw3P62K', 'X-Naver-Client-Secret':'qs5qe9yEHc'},{'X-Naver-Client-Id':'hOabRvueIRecs0g7YhsT', 'X-Naver-Client-Secret':'AsvffAZg2K'}]
headers = headers_list[datetime.datetime.today().day%5]

## html 초기세팅
html_msg_start = """
<!DOCTYPE html>
<html lang="en">
<head>
 <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
 <meta charset="utf-8">
 <meta name="author" content="jcat">
 <title>Safety News</title>

<style>
table{table-layout:; border-collapse:collapse; border-bottom:2px solid lightgrey;}
th{
 border-bottom: 2px solid #c72526;
 font-familly:'Gulim';
 font-size:15pt;
 font-weight:bold;
 color:#c72526;
}
th.cc{background-repeat:no-repeat;}
td{
 padding-left:15px; padding-right:15px;
 font-familly:'Gulim';
 font-size:15pt;
 font-weight:normal;
 color:#333;
 text-align:left;
}
a:link { color:black }
tr.rc{border-bottom: 2px solid lightgrey;}
td.rh{background-color:#f2f2f2; border-bottom: 2px solid lightgrey;text-align:center;font-weight:bold;}
td.cc{border-bottom: 1px solid lightgrey;border-left: 1px solid lightgrey;padding-top:4px;padding-bottom:4px;font-size:13pt;}
td.cce{border-bottom: 2px solid lightgrey;border-left: 1px solid lightgrey;padding-top:4px;padding-bottom:4px;font-size:13pt;}
td.rhe{background-color:#ffffff;  border-bottom:2px solid lightgrey;}
td.re{border-bottom:2px solid lightgrey;border-left: 1px solid lightgrey;}
td.ree{border-left:1px solid white; border-bottom:2px solid lightgrey;}

</style>
</head>
<body>
    <table cellspacing=0>
    <tr>
      <th class='rh' width="120px"> </th>
      <th class='cc' width="">뉴스제목</th>
    </tr>
"""
html_msg_end = """
    </table>
    <br/>
    <div style='width:360px;padding-top:10px;margin-left:5px;text-align:left;'>
        <div style='font-size:10pt;font-weight:bold;color:#333;'>마음으로 안전의식 행동으로 안전실천</div>
        <div style='font-size:28pt;font-weight:bold;margin-top:-10px;font-familly:'Arial';'><span style='color:red;'>S</span>afety <span style='color:red'>N</span>ews</div>
    </div>
    <div id='pubDesc' style='padding-top:12px;font-size:9pt;color:#777;margin-left:5px'>
    문의/개선 사항은<br>
    <b>강남/서부네트워크본부 NIT기술팀 전상영 과장<br>
    강남/서부네트워크본부 NIT기술팀 김동현 사원</b><br>
    에게 문의바랍니다.
    </div>
</body>
</html>
"""
html_msg_all = ""

# 각 ITC별 키워드 검색후 메일
for name,keywords in dic.items():
    print(name)
    if name not in dic_email:  ## 메일 미등록시 pass
        continue

    if 'all' in keywords:
        html_msg = html_msg_all
    else:
        # msg_array = []
        html_msg = html_msg_start
        for keyword in keywords:  ## 키워드별 묶음
            print(keyword + "------------------------------------------")
            try:
                params = {'query': keyword, 'display': 100, 'start': 1, 'sort': 'sim'}
                r = requests.get(url, params=params, headers=headers)
                j = r.json()
                newslist = j['items']
            except:
                print("naver api error")
                continue

            try:
                keyword_msg_array = []
                for item in newslist:  ## 뉴스리스트 - 제목,링크
                    # 제목 정비
                    replace_title = str(item['title']).replace("<b>", "").replace("</b>", "").replace("&quot;", "")

                    # 키워드와 제목 연관성 확인
                    if (jellyfish.jaro_similarity(keyword, replace_title) < 0.5):
                        print(keyword + " !!!제목연관성!!! " + replace_title)
                        continue
                    ## 중복뉴스 제거
                    similarit = 0
                    for msg_i in keyword_msg_array:
                        if (jellyfish.jaro_similarity(msg_i[1], replace_title) > 0.6):
                            print(msg_i[1] + " !!!중복!!! " + replace_title)
                            similarit = 1
                            break
                    if similarit == 1:  # 중복시 저장안하고 다음 뉴스 검색
                        continue

                    ## 일주일 이내만 추출
                    pubDate = datetime.datetime.strptime(item['pubDate'], '%a, %d %b %Y %H:%M:%S +0900')
                    _day = pubDate - datetime.datetime.now()
                    if _day.days < -6:  ## -1
                        break

                    keyword_msg_array.append([keyword, replace_title, item['link']])
                    # msg_array.append([keyword, replace_title, item['link']])

                ## 뉴스내용 html
                if len(keyword_msg_array) > 0:
                    html_table = "<tr class='rc'><td class='rh' rowspan="+str(len(keyword_msg_array))+">"+keyword+"</td>"
                    for idx in range(len(keyword_msg_array)):
                        if html_table.find("</tr>") > 0:
                            html_table += "<tr class='rc'>"
                        if idx == len(keyword_msg_array)-1:
                            html_table = html_table+"<td class='cce'><a href='"+keyword_msg_array[idx][2]+"'>"+keyword_msg_array[idx][1]+"</a></td></tr>"
                        else:
                            html_table = html_table+"<td class='cc'><a href='"+keyword_msg_array[idx][2]+"'>"+keyword_msg_array[idx][1]+"</a></td></tr>"
                    html_msg+=html_table
            except:
                print("why error")
                continue
        html_msg+=html_msg_end

    try:
        ## Mail 초기설정
        msgRoot = MIMEMultipart('related')  # 그대로 작성
        msgRoot['Subject'] = '[Safety News] ' + str(datetime.datetime.today().month) + '월 ' + str(datetime.datetime.today().day) + '일 안전뉴스레터'
        msgRoot['From'] = 'GnNitTeam@gmail.com'
        msgRoot['To'] = dic_email[name]['email']
        msgAlternative = MIMEMultipart('alternative')
        msgRoot.attach(msgAlternative)

        ## 메일 내용
        if name == '가가가': ## 전체수신내용
            html_msg_all = html_msg
        msgText = MIMEText(html_msg, 'html')
        # msgText = MIMEText(pd.DataFrame(msg_array, columns=['기업명', '제목', 'Link']).to_html(index=False, justify='center', col_space=100), 'html')
        msgAlternative.attach(msgText)

        ## 메일 전송
        s = smtplib.SMTP('smtp.gmail.com', 587)  # 세션 생성
        s.starttls()  # TLS 보안 시작
        s.login('gnnitteam', 'cywyatspmibwfxdz')  # 로그인 인증
        s.sendmail(msgRoot['From'],msgRoot['To'], msgRoot.as_string())  # 메일 보내기
        s.quit()
    except:
        print("mail send error")
        pass
