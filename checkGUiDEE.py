from __future__ import print_function
# coding: utf-8
import gspread
import json
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from flask import *

#ServiceAccountCredentials：Googleの各サービスへアクセスできるservice変数を生成します。
from oauth2client.service_account import ServiceAccountCredentials

#2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

#認証情報設定
#ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
credentials = ServiceAccountCredentials.from_json_keyfile_name('python-google-sheet.json', scope)

#OAuth2の資格情報を使用してGoogle APIにログインします。
gc = gspread.authorize(credentials)

#共有設定したスプレッドシートキーを変数[SPREADSHEET_KEY]に格納する。
SPREADSHEET_KEY2 = "1YHbhRnTtiON1nQMyo5O8kk8fDwiXcOJQihYxcO0-cVc"
worksheet2 = gc.open_by_key(SPREADSHEET_KEY2).worksheet("GUiDEE利用状況レポート")
worksheet3 = gc.open_by_key(SPREADSHEET_KEY2).worksheet("週次")
mentor_names = set(worksheet3.col_values(3))#「週次」の2列目からメンターの名前をsetで拾う
num_address_worksheet = gc.open_by_key(SPREADSHEET_KEY2).worksheet("メアド一覧")#「離脱者週次報告」の「メアド一覧」を開く→これをCAOのファイルに変えたい

name_id_dic = {k:v for k,v in zip(num_address_worksheet.col_values(2),num_address_worksheet.col_values(4))}#「離脱者週次報告」の「メアド一覧」を開いて,2行目と4行目から名前、アドレスを取得
id_name_dic = {v:k for k,v in name_id_dic.items()}#上記のname_id_dicの{name:address}を{address:name}に変える

pair_id_list = []
exception_pair_id_list = []

#メンターのカレンダーのみ見る
mentor_id_list = []
for name in mentor_names:#「離脱者週次報告」の「週次」から取得したmentorの名前のsetについて繰り返し処理
    if name != "メンター" and name!= "":#空欄と「メンター」の文字を除く

        mentor_id_list.append(name_id_dic[name])#mentorの名前をmentorのアドレスに変更(離脱者週次報告の「メアド一覧」)し、mentor_id_listに格納

#ペアのアドレスを取得
test_pairs= zip(worksheet3.col_values(4),worksheet3.col_values(7))#離脱者週次報告「週次」から、ペアのアドレスをzipで取得
test_pair_id_list =[]
for test_pair_id in test_pairs:
    test_pair_id_list.append(list(test_pair_id))#検証対象ペアを（メンター、メンティー）の順でリストを取得

#上記と同じように、順番逆のペアリストを作る
inverse_test_pairs= zip(worksheet3.col_values(7),worksheet3.col_values(4))
inverse_test_pair_id_list = []

for inverse_test_pair_id in inverse_test_pairs:
    inverse_test_pair_id_list.append(list(inverse_test_pair_id))#検証対象ペアを（メンティー、メンター）の順で取得


now = datetime.datetime.utcnow().isoformat() + 'Z'

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
def tra_Z_JST_isoformat(date):
    date = (datetime.datetime.strptime(date, "%Y-%m-%d %H:%M")-datetime.timedelta(hours=9)).isoformat()+"Z"
    return date

def tra_Z_JST_datetime(date2):
    date2 = datetime.datetime.strptime(date2, "%Y-%m-%d　%H:%M")-datetime.timedelta(hours=9)
    return date2


app = Flask(__name__)
@app.route("/", methods=["GET", "POST"])
def checkGUiDEE():
    """引数はdatetime型"""
    #GUiDEEを使ったペアの、[メンター、メンティー,ステータス,startの時間(datetime)]を格納したリストを作成する
    GUiDEE_list =[]
                  #GUIDEEリストに格納するリストを作成するためのリスト
    GUiDEE_list_factor = []
                  # 該当期間にGUiDEEのカードを発行し、かつステータスが「実施待ち」「準備中」ではないペアの[メンター.メンティー,ステータス,start_time]を格納するリスト
    done_GUiDEE_list = []
                  #該当期間にGUiDEEのカードを発行したペアの[メンター、メンティー]を格納するリスト
    done_GUiDEE_pair = []
                  
    done_GUiDEE_name_list = []
    
    if request.method == "GET":
        return """
        <p>いつ以降の使用者を把握したいですか？</p>
       <p>記入例）2019-11-01 10:00</p>
        <form action="/" method="POST">
        <input name="GUiDEEbefore" value = "2019-11-01 10:00"></input>
        
        <p>いつまでの使用者を把握したいですか？</p>
        <p>記入例）2019-11-01 19:00</p>
        <input name="GUiDEEafter" value = "2019-11-01 19:00"></input>
        <input type="submit" value="check">
        </form>

        """
    else:
        try:
            GUiDEE_timeMin = tra_Z_JST_datetime(request.form["GUiDEEbefore"])
            GUiDEE_timeMax = tra_Z_JST_datetime(request.form["GUiDEEafter"])
            
            mentor_mentee_status_start=zip(worksheet2.col_values(4),worksheet2.col_values(7),worksheet2.col_values(8),worksheet2.col_values(9))
                

            for mentor, mentee, status,start in mentor_mentee_status_start :#1行目を除く、startをdatetimeオブジェクトに変換する処理
                if (start != "start_at") and (start != " ") and (start != "　"):
                        
                    start_datetime = datetime.datetime.strptime(start, "%Y-%m-%d %H:%M") - datetime.timedelta(hours = 9)
                        
                    GUiDEE_list_factor = [mentor,mentee,status,start_datetime]
                    GUiDEE_list.append(GUiDEE_list_factor)

                #該当期間に1on1をしたペアの[メンター,メンティー,ステータス,スタート時間]をリストで格納
                #monitored_time = datetime.datetime.utcnow()  - datetime.timedelta(days = 10)
            for GUiDEE_list_factor in GUiDEE_list:
                if (GUiDEE_timeMin <=GUiDEE_list_factor[3] <=GUiDEE_timeMax) and (GUiDEE_list_factor[2] != "実施待ち") and(GUiDEE_list_factor[2] != "準備中"):
                    done_GUiDEE_list.append(GUiDEE_list_factor)

            #上記,monitored_GUiDEE_listから、「GUiDEEを使用したペアの名前」のみを取得
            for mentor_mentee_status_start in done_GUiDEE_list:
                    done_GUiDEE_pair.append(mentor_mentee_status_start[:2])

            for pair_ids in  done_GUiDEE_list:
                     if '#N/A' not in pair_ids:
                        done_GUiDEE_name_list.append([id_name_dic[pair_ids[0]],id_name_dic[pair_ids[1]]])
                    
                
            if done_GUiDEE_name_list:
                return "使用者はこの人達です！{}".format(done_GUiDEE_name_list)
            else:
                return """該当期間に使用者がいませんでした...
                            <p>離脱者週次報告の「GUiDEE利用状況レポート」は最新のものか確認してみてください！</p>"""
            
            
        except :
            return
            """値が不正でした！もう一度お願いします！
        
            いつ以降の使用者を把握したいですか？
            <form action="/" method="POST">
            <input name="GUiDEEbefore"></input>
        
            いつまでの使用者を把握したいですか？
            <form action="/" method="POST">
            <input name="GUiDEEafter"></input>
            </form>"""
    
if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=8888, threaded=True)
