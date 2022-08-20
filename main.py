# Googleスプレッドシートとの連携に必要なライブラリ
import gspread
from oauth2client.service_account import ServiceAccountCredentials

import yfinance as yf
import pandas as pd
import time
import datetime
import os
import tweepy
from dotenv import find_dotenv, load_dotenv
import matplotlib.pyplot as plt
from PIL import Image, ImageFont, ImageDraw
import numpy as np

# 設定
file_name = 'python-investment'

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

# 環境変数の読み込み
load_dotenv(find_dotenv())

# 辞書オブジェクト。認証に必要な情報をHerokuの環境変数から呼び出している
credential = {
"type": "service_account",
"project_id": os.environ['SHEET_PROJECT_ID'],
"private_key_id": os.environ['SHEET_PRIVATE_KEY_ID'],
"private_key": os.environ['SHEET_PRIVATE_KEY'],
"client_email": os.environ['SHEET_CLIENT_EMAIL'],
"client_id": os.environ['SHEET_CLIENT_ID'],
"auth_uri": "https://accounts.google.com/o/oauth2/auth",
"token_uri": "https://oauth2.googleapis.com/token",
"auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
"client_x509_cert_url":  os.environ['SHEET_CLIENT_X509_CERT_URL']
}
# スプレッドシートにアクセス
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credential, scope)
gc = gspread.authorize(credentials)
sh = gc.open(file_name)

CONSUMER_KEY = os.environ['CONSUMER_KEY_2018']
CONSUMER_SECRET = os.environ['CONSUMER_SECRET_2018']
ACCESS_TOKEN = os.environ['ACCESS_TOKEN_2018']
ACCESS_TOKEN_SECRET = os.environ['ACCESS_TOKEN_SECRET_2018']
BEARER_TOKEN = os.environ['BEARER_TOKEN_2018']

# 変数は適宜さっきの奴で指定
client = tweepy.Client(BEARER_TOKEN, 
                       CONSUMER_KEY,
                       CONSUMER_SECRET, 
                       ACCESS_TOKEN, 
                       ACCESS_TOKEN_SECRET)

auth = tweepy.OAuthHandler(CONSUMER_KEY, CONSUMER_SECRET)
auth.set_access_token(ACCESS_TOKEN, ACCESS_TOKEN_SECRET)
API = tweepy.API(auth)

# シートから全部から読み込み
def get_records(wks):
    record = pd.DataFrame(wks.get_all_records())
    return record

def get_data(start_date,end_date,df_IPO_list,reverse):

    df = pd.DataFrame(index=[], columns=['Open', 'High', 'Low', 'Close', 'Adj Close', 'Volume', 'Ticker', 'Rate'])
    results = []
    for ticker, name in zip(df_IPO_list['Ticker'],df_IPO_list['Name']):
        data = yf.download(str(ticker) , start = start_date, end = end_date)
        data['Ticker'] = ticker
        #出来高0のデータがあるので削除
        data = data[data['Volume'] != 0]
        #一番古い日の株価を基準にする
        stn = data['Adj Close'][0] 
        #基準からの増加率を%で計算
        data['Rate'] = round(data['Adj Close']*100/stn,2) 
        # https://note.nkmk.me/python-pandas-at-iat-loc-iloc/ pandasで任意の位置の値を取得・変更するat, iat, loc, iloc
        result = {'Ticker':ticker, 'Name':name ,'Rate':data.at[data.index[-1],'Rate']}
        results.append(result)
        df = pd.concat([df, data],sort=False).drop_duplicates()

    results = sorted(results, key=lambda x:x['Rate'],reverse=reverse)

    return results ,df

def graph(tweet_list,results,df):
    figsize_px = np.array([1200, 675])
    dpi = 100
    figsize_inch = figsize_px / dpi
    fig, ax = plt.subplots(figsize=figsize_inch, dpi=dpi) #ベースを作る
    ax.set_ylabel('%') #y軸のラベルを設定する
    for i in results[0:5]:
        df1 = df[df['Ticker'] == i['Ticker']]
        shortName = yf.Ticker(str(i['Ticker'])).info['shortName']
        ax.plot(df1['Rate'],label=shortName) #グラフを重ねて描写していく
        tweet_list = tweet_list + '\n' + i['Name'] + ' ' + str(i['Rate']) + '%'

    ax.legend()
    ax.grid()
    fig.savefig('img1.png', bbox_inches='tight')

    return tweet_list

def tweet(tweet_list,URL):
    media_ids = []
    img  = API.media_upload('./img1.png')
    media_ids.append(img.media_id)
    tweet = tweet_list + '\n' +URL
    client.create_tweet(text=tweet, media_ids=media_ids)

wks_config = sh.worksheet('config')
df_config = get_records(wks_config)

def main(wks_IPO_list,start_date,end_date,tweet_list,reverse,URL):
    df_IPO_list = get_records(wks_IPO_list)
    results , df = get_data(start_date,end_date,df_IPO_list,reverse)
    tweet_list = graph(tweet_list,results,df)
    URL = URL
    tweet(tweet_list,URL)

if __name__ == "__main__":
    
    for i in range(0,len(df_config)):
        wks_IPO_list = sh.worksheet(df_config['wks_IPO_list'][i])
        start_date = datetime.datetime.strptime(df_config['start_date'][i], '%Y/%m/%d').date()
        end_date = datetime.datetime.strptime(df_config['end_date'][i], '%Y/%m/%d').date()
        tweet_list = df_config['tweet_list'][i]
        # 昇順か降順か選択
        reverse = bool(df_config['reverse'][i])
        URL = df_config['URL'][i]
      
        main(wks_IPO_list,start_date,end_date,tweet_list,reverse,URL)
        time.sleep(60)

# git init
# git add .
# git commit -m "1st commit"
# git remote add origin git@github.com:sheep-0411/ipo.git
# git push -u origin main