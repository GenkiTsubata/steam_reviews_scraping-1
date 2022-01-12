import json
import requests
import pandas as pd
from io import BytesIO
from flask import Flask, render_template, request, redirect, send_file, Response
from flask_bootstrap import Bootstrap
from werkzeug.datastructures import ImmutableOrderedMultiDict
import openpyxl


app = Flask(__name__)
bootstrap = Bootstrap(app)

"""
【構造】
1. webページよりユーザーがゲームIDを入力
2. ダウンロードボタンが押されると、入力されたゲームIDがフォームによりPOSTとして"/download"へ送信される
3. 2.により取得したゲームIDを元にSteamWorksAPIを叩き、レビュー内容とプレイ時間を受け取る
4. 受け取ったレビューデータをExcel形式に変換しユーザーにダウンロードさせる
"""


@app.route('/')
def index():
    # トップページへアクセスした際に表示するテンプレートを指定している
    # render_templateを使用すると、templatesフォルダ内のファイルを探しに行く
    # 以下で参照しているのは、templates/index.html
    return render_template('index.html')


@app.route('/download', methods=['GET', 'POST'])
def excel_file_download():
    try:
        if request.method == 'POST':
            # requestで受け取ったデータの順番を維持 今回のケースだと必要ないかも
            request.parameter_storage_class = ImmutableOrderedMultiDict

            # request内に存在するフォームから送信されたデータを取り出している
            values = request.form
            game_id = values.get('gameid')

            url = f'https://store.steampowered.com/appreviews/{game_id}?json=1'

            # APIを叩く際のオプション
            params = {
                'filter': 'recent',
                'language': 'japanese',
                'num_per_page': '100',
            }

            response = requests.get(url, params)

            # responseで受け取った文字列データがjson形式だと宣言する
            res = json.loads(response.text)

            # 取得したレビューの総件数を取得
            total = res['query_summary']['total_reviews']

            # 1ページ100件と考え、何ページ分読み込む必要があるか計算
            pages = (total // 100) + 1

            # data_listという空のリストを用意し、スクレイピングしたデータを[レビュー本文, プレイ時間(h)]という形で追加していく
            data_list = []
            for i in range(pages):
                # 100件取得（APIの仕様で100件までしか取得できない）
                response = requests.get(url, params)
                res = json.loads(response.text)

                # データの取得と整形
                for j in range(100):
                    try:
                        data = [res['reviews'][j]['review'], round(res['reviews'][j]['author']['playtime_forever'] / 60, 1)]
                    except Exception as e:
                        break
                    # データ追加
                    data_list.append(data)

                # APIから取得したcursorという値を次回のリクエストに加えることで、前回の続きからレビューを取得
                params['cursor'] = res['cursor']

            # 追加したデータをDataFrameという形式に変換し、出力しやすくする
            df = pd.DataFrame(data_list, columns=['レビュー本文', 'プレイ時間(h)'])

            # excelファイルの一時的な出力先を用意
            textStream = BytesIO()

            # 先程の出力先にexcelファイルを書き出し
            df.to_excel(textStream, encoding='utf-8')

            # ユーザーへの戻り値としてexcelファイルを送信
            return Response(
                textStream.getvalue(),
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                headers={'Content-disposition':
                         'attachment; filename=reviews_{}.xlsx'.format(game_id)}
            )
        else:
            # POST以外が送信された場合のハンドリング
            # return を使用すると、続く文字列がwebページとして出力される
            return 'ERROR DATA'

    except Exception as e:
        # エラーが起きた場合のハンドリング
        return 'ERROR'


# webサーバーを起動させるためのおまじない
if __name__ == '__main__':
    app.run()
