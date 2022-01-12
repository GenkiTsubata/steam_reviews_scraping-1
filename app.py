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


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/download', methods=['GET', 'POST'])
def excel_file_download():
    try:
        if request.method == 'POST':
            request.parameter_storage_class = ImmutableOrderedMultiDict

            values = request.form
            game_id = values.get('gameid')

            url = f'https://store.steampowered.com/appreviews/{game_id}?json=1'

            params = {
                'filter': 'recent',
                'language': 'japanese',
                'num_per_page': '100',
            }

            response = requests.get(url, params)

            res = json.loads(response.text)
            total = res['query_summary']['total_reviews']

            pages = (total // 100) + 1
            print(total)

            data_list = []
            for i in range(pages):
                response = requests.get(url, params)
                res = json.loads(response.text)

                for j in range(100):
                    try:
                        data = [res['reviews'][j]['review'], round(res['reviews'][j]['author']['playtime_forever'] / 60, 1)]
                    except Exception as e:
                        break
                    data_list.append(data)
                params['cursor'] = res['cursor']

            df = pd.DataFrame(data_list, columns=['レビュー本文', 'プレイ時間'])

            textStream = BytesIO()

            df.to_excel(textStream, encoding='utf-8')

            return Response(
                textStream.getvalue(),
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                headers={'Content-disposition':
                         'attachment; filename=reviews_{}.xlsx'.format(game_id)}
            )
        else:
            return 'ERROR'

    except Exception as e:
        print('ERROR')
        return 'ERROR'


if __name__ == '__main__':
    app.run()

"""
game_id = '945360'

url = f'https://store.steampowered.com/appreviews/{game_id}?json=1'

params = {
    'filter': 'recent',
    'language': 'japanese',
    'num_per_page': '100',
}

response = requests.get(url, params)

res = json.loads(response.text)
total = res["query_summary"]['total_reviews']

pages = (total // 100) + 1
print(total)

data_list = []
for i in range(pages):
    response = requests.get(url, params)
    res = json.loads(response.text)

    for j in range(100):
        try:
            data = [res["reviews"][j]["review"], res["reviews"][j]["author"]["playtime_forever"]/60]
        except Exception as e:
            break
        data_list.append(data)
    params['cursor'] = res["cursor"]


df = pd.DataFrame(data_list, columns=['レビュー本文', 'プレイ時間'])
print(df)

df.to_excel(r'D:/steam_reviews.xlsx', encoding='utf-8')"""