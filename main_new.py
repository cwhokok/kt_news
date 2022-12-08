import requests
import datetime
import pandas as pd
import jellyfish
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from bs4 import BeautifulSoup

# 초기 Data
db_jijache = pd.read_csv(os.path.dirname(os.path.realpath(__file__))+'/db_jijache.csv', encoding='cp949') ## 담당자-키워드
## html 초기세팅
html_msg_start = """
<!DOCTYPE html>
<html lang="en">
<head>
 <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
 <meta charset="utf-8">
 <meta name="author" content="jcat">
 <title>press release</title>

<style>
table{table-layout:; border-collapse:collapse; border-bottom:2px solid lightgrey;}
th{
 border-bottom: 2px solid #c72526;
 font-familly:'Gulim';
 font-size:15pt;
 font-weight:bold;
 color:#c72526;
 width:120px;
}
th.cc{background-repeat:no-repeat;width:500px;}
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
    <h2 sytle='width:100%;font-size:25pt;text-align:center;'>["""+datetime.datetime.today().strftime("%Y년%m월%d일")+""" - 보도자료]</h2>"""
html_msg_end = """
    </table>
    <br/>
    <div id='pubDesc' style='padding-top:12px;font-size:9pt;color:#777;margin-left:5px'>
    문의/개선 사항은<br>
    <b>강남/서부법인고객본부 컨설팅기획팀 이강욱 과장<br>
    강남/서부네트워크본부 NIT기술팀 전상영 과장</b><br>
    에게 문의바랍니다.
    </div>
</body>
</html>
"""
html_msg_all = ""


html_msg = html_msg_start
new_line=True
table_title= " "

for keyword in db_jijache['jijache']:   ## 지자체 리스트 순회
    print(keyword + "------------------------------------------")
    if keyword.startswith("pass"):  ## pass 발생시 테이블 분리 / 테이블 제목 추출
        new_line = True
        table_title = keyword.split()[1]
        continue

    sub_title= " "
    similar_keyword= " "
    if '/' in keyword:  ## 서울시청/서울특별시청 경우 >> 검색은 서울시청, 표시는 서울특별시
        sub_title=keyword.split('/')[1]
        keyword=keyword.split('/')[0]
    else:
        sub_title=keyword

    if ' ' in keyword:  ## 부산 강서구청 경우 유사도 검색은 강서구청으로만
        similar_keyword = keyword.split()[1]
    else:
        similar_keyword = keyword

    start = 1   ## 검색페이지 index
    pre_df = pd.DataFrame()
    ## Data 크롤링
    while True:
        try:
            url=' '
            if datetime.datetime.today().weekday()!=0:  ## 월요일 아닌경우
                url = 'https://search.naver.com/search.naver?where=news&query={}&sm=tab_opt&sort=0&photo=4&field=0&pd=4&ds=&de=&docid=&related=0&mynews=0&office_type=0&office_section_code=0&news_office_checked=&nso=so%3Ar%2Cp%3A1d&is_sug_officeid=0&start={}&refresh_start=0'.format(
                    keyword, start)
            else:
                url = 'https://search.naver.com/search.naver?where=news&query={}&sm=tab_opt&sort=0&photo=4&field=0&pd=4&ds={}&de={}&docid=&related=0&mynews=0&office_type=0&office_section_code=0&news_office_checked=&nso=so%3Ar%2Cp%3A1d&is_sug_officeid=0&start={}&refresh_start=0'.format(
                    keyword,datetime.datetime.today().strftime("%Y.%m.%d") if datetime.datetime.today().weekday() != 0 else (datetime.datetime.today() - datetime.timedelta(days=3)).strftime("%Y.%m.%d"),
                    datetime.datetime.today().strftime("%Y.%m.%d"), start)
            headers = {
                'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.150 Safari/537.36',
                'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9'}
            response = requests.get(url, headers=headers)
            soup = BeautifulSoup(response.text, 'lxml')
            news_title = [title['title'] for title in soup.find_all('a', attrs={'class': 'news_tit'})]  # 기사 제목
            news_url = [url['href'] for url in soup.find_all('a', attrs={'class': 'news_tit'})]  # 기사 url

            df = pd.DataFrame({'검색어': keyword, 'title': news_title, 'link': news_url})
            # print(df)
            pre_df = pd.concat([pre_df, df], ignore_index=True)
            if len(df) <= 1:  ##마지막 페이지 (기사 하나)
                break
            start += 10

        except:  # 오류발생시 몇 페이지까지 크롤링했는지 page를 확인하기
            print(start)
            break

    try:
        ## Data 처리
        keyword_msg_array = []
        for index, item in pre_df.iterrows():
            # 제목 정비
            replace_title = str(item['title']).replace("<b>", "").replace("</b>", "").replace("&quot;", "")

            # 키워드와 제목 연관성 확인
            if (jellyfish.jaro_similarity(similar_keyword, replace_title) < 0.6):
                continue
            ## 중복뉴스 제거
            similarit = 0
            for msg_i in keyword_msg_array:
                if (jellyfish.jaro_similarity(msg_i[1], replace_title) > 0.7):
                    print(msg_i[1] + " !!!중복!!! " + replace_title)
                    similarit = 1
                    break
            if similarit == 1:  # 중복시 저장안하고 다음 뉴스 검색
                continue

            keyword_msg_array.append([keyword, replace_title, item['link']])

        ## 뉴스내용 html 처리
        if len(keyword_msg_array) > 0: ## 내용있으면 테이블 추가
            ## 새로운 테이블
            if new_line:
                if table_title != "서울특별시":
                    html_msg = html_msg + """</table>"""
                html_msg = html_msg + """
                <br><br><b sytle='font-size:15pt;'>[""" + table_title + """]</b><br>
                <table cellspacing=0>
                <tr>
                  <th class='rh' width="120px">고객명</th>
                  <th class='cc' width="">뉴스제목</th>
                </tr>
                """
                new_line=False

            html_table = "<tr class='rc'><td class='rh' rowspan=" + str(len(keyword_msg_array)) + ">" + sub_title + "</td>"
            for idx in range(len(keyword_msg_array)):
                if html_table.find("</tr>") > 0:
                    html_table += "<tr class='rc'>"
                if idx == len(keyword_msg_array) - 1:
                    html_table = html_table + "<td class='cce'><a href='" + keyword_msg_array[idx][2] + "'>" + \
                                 keyword_msg_array[idx][1] + "</a></td></tr>"
                else:
                    html_table = html_table + "<td class='cc'><a href='" + keyword_msg_array[idx][2] + "'>" + \
                                 keyword_msg_array[idx][1] + "</a></td></tr>"
            html_msg += html_table
    except:
        print("why error")
        continue

html_msg = html_msg + html_msg_end


def send_email(mail_to):
    try:
        ## Mail 초기설정
        msgRoot = MIMEMultipart('related')  # 그대로 작성
        msgRoot['Subject'] = '[press release] ' + str(datetime.datetime.today().month) + '월 ' + str(
            datetime.datetime.today().day) + '일'
        msgRoot['From'] = 'GnNitTeam@gmail.com'
        msgRoot['To'] = mail_to
        msgAlternative = MIMEMultipart('alternative')
        msgRoot.attach(msgAlternative)

        ## 메일 내용
        msgText = MIMEText(html_msg, 'html')
        # msgText = MIMEText(pd.DataFrame(msg_array, columns=['기업명', '제목', 'Link']).to_html(index=False, justify='center', col_space=100), 'html')
        msgAlternative.attach(msgText)

        ## 메일 전송
        s = smtplib.SMTP('smtp.gmail.com', 587)  # 세션 생성
        s.starttls()  # TLS 보안 시작
        s.login('gnnitteam', 'cywyatspmibwfxdz')  # 로그인 인증
        s.sendmail(msgRoot['From'], msgRoot['To'], msgRoot.as_string())  # 메일 보내기
        s.quit()
    except:
        print("mail send error")
        pass


##mail send
send_email('sangyoung.jun@kt.com')
# send_email('sunhee.ryu@kt.com')
# send_email('kangwook.lee@kt.com')
# send_email('young-in.kim@kt.com')
# send_email('daehan.park@kt.com')
