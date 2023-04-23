#seleniumをcolaboratoryで使えるようにする#########################################################
%%shell
# Ubuntu no longer distributes chromium-browser outside of snap
#
# Proposed solution: https://askubuntu.com/questions/1204571/how-to-install-chromium-without-snap

# Add debian buster
cat > /etc/apt/sources.list.d/debian.list <<'EOF'
deb [arch=amd64 signed-by=/usr/share/keyrings/debian-buster.gpg] http://deb.debian.org/debian buster main
deb [arch=amd64 signed-by=/usr/share/keyrings/debian-buster-updates.gpg] http://deb.debian.org/debian buster-updates main
deb [arch=amd64 signed-by=/usr/share/keyrings/debian-security-buster.gpg] http://deb.debian.org/debian-security buster/updates main
EOF

# Add keys
apt-key adv --keyserver keyserver.ubuntu.com --recv-keys DCC9EFBF77E11517
apt-key adv --keyserver keyserver.ubuntu.com --recv-keys 648ACFD622F3D138
apt-key adv --keyserver keyserver.ubuntu.com --recv-keys 112695A0E562B32A

apt-key export 77E11517 | gpg --dearmour -o /usr/share/keyrings/debian-buster.gpg
apt-key export 22F3D138 | gpg --dearmour -o /usr/share/keyrings/debian-buster-updates.gpg
apt-key export E562B32A | gpg --dearmour -o /usr/share/keyrings/debian-security-buster.gpg

# Prefer debian repo for chromium* packages only
# Note the double-blank lines between entries
cat > /etc/apt/preferences.d/chromium.pref << 'EOF'
Package: *
Pin: release a=eoan
Pin-Priority: 500


Package: *
Pin: origin "deb.debian.org"
Pin-Priority: 300


Package: chromium*
Pin: origin "deb.debian.org"
Pin-Priority: 700
EOF

# Install chromium and chromium-driver
apt-get update
apt-get install chromium chromium-driver

# Install selenium
pip install selenium


!cp /usr/lib/chromium-browser/chromedriver/usr/bin


#モジュールのインストール#########################################################
from bs4 import BeautifulSoup
import requests
import pandas as pd
import datetime
from datetime import datetime as dt
import re
from dateutil import relativedelta
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
import sys
sys.path.insert(0,"/usr/lib/chromium-browser/chromedriver")

#前準備#########################################################
chrome_options=webdriver.ChromeOptions()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
wd=webdriver.Chrome("chromedriver",chrome_options=chrome_options)


#記入用の表を作成
cols=["タイトル","主催","開催日","開催時間","会場","内容","定員","参加費","申込締切日","URL"]

#格納用の変数を定義
title=""
syusai="マクロミル"
holdday=""
holdtime=""
venue=""
info=""
capacity=""
fee=""
deadline=""
URL=""

df=pd.DataFrame(columns=cols)
now=dt.today().strftime("%Y{}%m{}%d{}").format(*"年月日")

not_info="詳しくはURLにアクセス"


#セミナー情報取得#########################################################
#セミナーサイトにアクセスし、HTML構文を取得する
res=requests.get("https://www.macromill.com/seminar/")
soup=BeautifulSoup(res.content,"html.parser")

#セミナー一覧を取得する
seminarHTML=soup.select_one("#secPlan_area > div.un_pressList >dl")
seminarListdt=seminarHTML.select("dt")  #開催日、開催時間を取得
seminarListdd=seminarHTML.select("dd")  #タイトルとURLを取得

for i in range(len(seminarListdt)):

  dfAppendList=[]

  seminarHTML=soup.select_one("#secPlan_area > div.un_pressList >dl")
  seminarListdt=seminarHTML.select("dt")  #開催日、開催時間を取得
  seminarListdd=seminarHTML.select("dd")  #タイトルとURLを取得

  #開催時間、開催日時、会場
  applyStatus=seminarListdt[i].select("span")[1].get_text()
  venue=seminarListdt[i].select("span")[2].get_text()
  holdinfo=seminarListdt[i].select_one("span").get_text()

  try:
    holdtime=re.findall("\d\d:\d\d～\d\d:\d\d",holdinfo)[0] #正規表現で時間のみ取得

  except IndexError:
    holdtime="不明"

  holdday=holdinfo.split("(")[0] #日付のみ取得

  #開催日時kが今日でない、かつ募集中のままなら取得する
  if holdday!=now and applyStatus=="募集中":

    #参加費
    fee=seminarListdd[i].select_one("span").get_text()

    #タイトル
    title=seminarListdd[i].select_one("a").get_text() #タイトル
    title=title.replace("\t","")
    title=title.replace("\n","")

    #URL
    URL=seminarListdd[i].select_one("a[href]").get("href")

    #詳細サイトにアクセス
    wd.get(URL)

    #とれなかったら取得できない情報に定数を入れる
    try:
      #内容
      information=wd.find_element(By.CSS_SELECTOR,"#block3>div").text
      #切り出し位置を条件分岐
      split_str_osusume=information.find("このような方へおすすめです")
      split_str_gaiyou=information.find("開催概要")

      if split_str_osusume>0:
        info=information.split("このような方へおすすめです")[0]

      elif split_str_gaiyou>0:
        info=information.split("開催概要")[0]

      else:
        #分割せず処理
        info=not_info

      #定員、締切日
      blockinfo=wd.find_element(By.CSS_SELECTOR,"#block3>div>div.el_tabkeBlock.el_tabkeType_list.hp_mgnTopM").text
      capainfo=blockinfo.split("定員")[1]
      capainfo=capainfo.split("\n")
      capacity=capainfo[1] #定員
      deadline=capainfo[2].split("（")[0] #締切日

    except:
      info=not_info
      capacity=not_info
      deadline=not_info

    #データ追加
    dfAppendList.append(title)
    dfAppendList.append(syusai)
    dfAppendList.append(holdday)
    dfAppendList.append(holdtime)
    dfAppendList.append(venue)
    dfAppendList.append(info)
    dfAppendList.append(capacity)
    dfAppendList.append(fee)
    dfAppendList.append(deadline)
    dfAppendList.append(URL)

    df.loc[i]=dfAppendList

    #元のページに戻る
    res=requests.get("https://www.macromill.com/seminar/")
    soup=BeautifulSoup(res.content,"html.parser")


#スプレッドシートに転記#########################################################
!pip install --upgrade gspread #スプレッドシートを操作するためのツール
!pip install --upgrade oauth2client #ドライブの情報にアクセスするためのツール
!pip install --upgrade gspread_dataframe

#アクセスの認証手続き
from google.colab import auth
auth.authenticate_user()

import gspread
from google.auth import default

creds, _=default()
gc=gspread.authorize(creds)

SprSheetName="セミナー転記用"   #ブック名を定義
workbook=gc.open(SprSheetName)  #ブックを指定
worksheet=workbook.worksheet(syusai)  #シートを指定

from gspread_dataframe import set_with_dataframe
set_with_dataframe(worksheet,df,row=1,col=1,include_index=False,resize=False)