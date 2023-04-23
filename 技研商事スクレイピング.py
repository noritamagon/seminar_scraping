#モジュールのインストール#########################################################
from bs4 import BeautifulSoup
import requests
import pandas as pd
import datetime
from datetime import datetime as dt
import re
from dateutil import relativedelta


#前準備#########################################################
#記入用の表を作成
cols=["タイトル","主催","開催日","開催時間","会場","内容","定員","参加費","申込締切日","URL"]

#格納用の変数を定義
title=""
syusai="技研商事インターナショナル"
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


#セミナー情報取得#########################################################
#セミナーサイトにアクセスし、HTML構文を取得する
res=requests.get("https://www.giken.co.jp/seminar-event/")
soup=BeautifulSoup(res.content,"html.parser")

#セミナー一覧を「info-list-column active」から取得
seminarHTML=soup.select_one("#re-contents > section > div > div.info-list-column.active")
seminarList=seminarHTML.select(".column-box")


#セミナーの数だけループする
for i in range(len(seminarList)):

  dfAppendList=[]

  #開催日時。開催時間
  holdinfo=seminarList[i].select_one(".day-category").get_text()

  try:
    holdtime=re.findall("\d\d:\d\d～\d\d:\d\d",holdinfo)[0] #正規表現で時間のみ取得
  
  except IndexError:
    holdtime="不明"
  
  holdday=holdinfo.split("(")[0] #日付のみ取得

  #開催日時が今日でなければ取得する
  if holdday!=now:

    #タイトル
    title=seminarList[i].select_one("h3").get_text()

    #会場
    holdrequireList=seminarList[i].select("dl dd")
    venue=holdrequireList[0].get_text() #会場
    fee=holdrequireList[1].get_text() #料金
    capacity=holdrequireList[2].get_text() #参加人数

    #deadline（期限の記入がないため、仮で開催日の前日を締切日に設定する）
    tdatetime=dt.strptime(holdday,"%Y年%m月%d日")
    deadline=tdatetime-datetime.timedelta(days=1)
    deadline=deadline.strftime("%Y{}%m{}%d{}").format(*"年月日")

    #URL
    URL=seminarList[i].select_one("a[href]").get("href")

    #URLにアクセスし内容を取得する
    res=requests.get(URL)
    soup=BeautifulSoup(res.content,"html.parser")
    info=soup.select_one("#seminar >section:nth-of-type(1) >div >p").get_text()

    #dataframe処理
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
    res=requests.get("https://www.giken.co.jp/seminar-event/")
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