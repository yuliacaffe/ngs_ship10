import win32com.client
import pandas as pd
import os
import sys
from tkinter import messagebox
import datetime
import json

# Pythonを実行した場所の取得
py_path = os.path.dirname(os.path.abspath(sys.argv[0]))

# 今日の日付を取得
today = datetime.date.today()

# 曜日を日本語に変換するためのリスト
weekdays = ["月", "火", "水", "木", "金", "土", "日"]

# フォーマット日付
formatted_date = today.strftime(f"%Y年%m月%d日({weekdays[today.weekday()]})")

# Excel起動
xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = False

# Excelファイル読み込み
workbook = xl.Workbooks.Open(f"{py_path}/メンバー一覧.xlsx")
worksheet = workbook.Worksheets("Member List")

# テーブルデータの取得
table = worksheet.ListObjects("Member_Table")
data = table.Range.Value
df = pd.DataFrame(data[1:], columns=data[0])

workbook.Close(SaveChanges=False)  # Excelを閉じる

# 条件(※在籍中)でフィルタリング
fil_df = df[df["Status"] == 1.0]

# 不要な列を除外
fil_df = fil_df.drop(columns=["Joining Route", "Status"])

# 改行を <br> に変換
fil_df = fil_df.replace("\n", "<br>", regex=True)

# HTMlテーブルを作成
html_table = fil_df.to_html(escape=False, index=False)  # index=Falseは、DataFrame各行の先頭インデックスを除外

# クミ設定ファイル読み込み
with open(f"{py_path}/kumi_config.json", "r", encoding="utf-8") as f:
    config = json.load(f)

labels = config["labels"]
colors = config["colors"]

# countMap の JS オブジェクトを Python が自動生成
count_map_js = "{ " + ", ".join([f'"{l}": 0' for l in labels]) + " }"

# JS 用の配列文字列生成
labels_js = "[" + ", ".join([f'"{l}"' for l in labels]) + "]"
colors_js = "[" + ", ".join([f'"{c}"' for c in colors]) + "]"

# CSSスタイルを定義
css_style = """
<style>
/* ページ全体のスタイル */
body {
    background-color: black;
    color: white;
}

/* 通常のリンク */
a {
    color: yellow;
}

/* 訪問済みのリンク */
a:visited {
    color: lime;
}

/* テーブル全体のスタイル */
table {
    width: 90%;
    border-collapse: collapse;
}

/* テーブルヘッダーのスタイル */
th {
    background-color: #f2f2f2;
    color: #333;
    font-weight: bold;
    padding: 8px;
    text-align: left;
    border: 1px solid #ddd;
}

/* テーブルデータのスタイル */
td {
    padding: 8px;
    text-align: left;
    border: 1px solid #ddd;
}

/* テーブルストライプのスタイル */
tr:nth-child(even) {
    background-color: #8b0000;
}

/* ソートアイコンのスタイル */
.sort-icon {
    float: right;
    margin-top: 3px;
    margin-left: 5px;
}
</style>
"""

# JavaScriptでのテーブルソート機能を定義
javascript = f"""
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/2.7.2/Chart.bundle.js"></script>
<script>
document.addEventListener('DOMContentLoaded', function() {{

    // 表のヘッダーと行を取得
    var thList = Array.from(document.querySelectorAll('th'));
    var tbody = document.querySelector('tbody');
    var trList = Array.from(tbody.querySelectorAll('tr'));

    // --- グラフ描画用関数 ---
    function drawKumiChart() {{
        const labels = {labels_js};
        const countMap = {count_map_js};

        trList.forEach(row => {{
            const cells = row.querySelectorAll('td');
            const kumi = cells[2].textContent.trim(); // 組列インデックス:"2"
            if (countMap.hasOwnProperty(kumi)) {{
                countMap[kumi]++;
            }}
        }});

        const data = labels.map(label => countMap[label] || 0);

        // 在籍者の割合のグラフ
        var ctx = document.getElementById("kumi").getContext('2d');
        var kumi = new Chart(ctx, {{
            type: 'pie',
            data: {{
                labels: labels,
                datasets: [{{
                    backgroundColor: {colors_js},
                    data: data,
                    borderColor: "#696969"
                }}]
            }},
            options: {{
                title: {{
                    display: true,
                    text: '◆組の勢力!?割合',
                    fontColor: '#ffffff' // タイトル文字色
                }},
                legend: {{
                    labels: {{
                        fontColor: '#ffffff' // 凡例ラベルの文字色
                    }}
                }}
            }}
        }});
    }}

    // グラフ描画
    drawKumiChart();


    // 各ヘッダーにクリックイベントを追加
    thList.forEach(function(th) {{
        th.addEventListener('click', function() {{

            // クリックされたヘッダーのインデックスを取得
            var hedIndex = thList.indexOf(th);

            // 昇順 or 降順判定
            var isAsc = !th.classList.contains('asc');

            // 行ソート
            trList.sort(function(a, b) {{
                var valueA = a.children[hedIndex].textContent.trim();
                var valueB = b.children[hedIndex].textContent.trim();
                return isAsc ? valueA.localeCompare(valueB, undefined, {{ numeric: true, sensitivity: 'base' }}) : valueB.localeCompare(valueA, undefined, {{ numeric: true, sensitivity: 'base' }});
            }});

            // ソートされた行を再配置
            trList.forEach(function(row) {{
                tbody.appendChild(row);
            }});

            // ソートアイコンとクラスをリセット
            resetSortIcons();
            resetSortClasses();

            // ソートされたヘッダーにクラスを追加し、アイコンの色を変更
            th.classList.add(isAsc ? 'asc' : 'desc', 'sorted');
            var sortIcon = th.querySelector('.sort-icon');
            sortIcon.style.color = isAsc ? 'red' : 'blue';
        }});

        // ソートアイコンを追加
        var sortIcon = document.createElement('span');
        sortIcon.className = 'sort-icon';
        sortIcon.innerHTML = '&#x25B2;&#x25BC;';
        th.appendChild(sortIcon);

    }});

    // ソートアイコンの色をリセットする関数
    function resetSortIcons() {{
        var iconList = Array.from(document.querySelectorAll('.sort-icon'))
        iconList.forEach(function(icon) {{
            icon.style.color = 'black';
        }});
    }}

    // ソート用のクラスをリセットする関数
    function resetSortClasses() {{
        thList.forEach(function(th) {{
            th.classList.remove('sorted', 'asc', 'desc');
        }});
    }}
}});
</script>
"""

# HTMLをTextファイルに書き込み
with open(f"{py_path}/TeamMemberList.html", "w", encoding="utf-8") as file:
    file.write(f'''<!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>チームメンバー一覧♬</title>
        {css_style}
    </head>
    <body>
    <h1>チームメンバー一覧♬</h1>
    <p>※ソート状態：デフォルト(黒)、昇順(赤)、降順(青)
    <br>※在籍者数：{(df["Status"] == 1.0).sum()}名
    <br>※更新日：<span id="update_date">{formatted_date}</span></p>
    {html_table}
    <br><br>
    <h3>◆組列の集計</h3>
    <div style="width: 50%;"><canvas id="kumi"></canvas></div>
    {javascript}
    <br><br>
    <h3>◆加入者の集計</h3>
    <p><small><small>◆◇◆◇◆◇◆◇◆◇◆◇◆◇ 集計 ◆◇◆◇◆◇◆◇◆◇◆◇◆◇
    <p>◆累計加入者数（在籍者数）：<span id="join_data1">{len(df)}</span>名（<span id="join_data2">{(df["Status"] == 1.0).sum()}</span>名）</p>
    <p>◇ユリア・ルート加入者数（在籍者数）：<span id="join_data3">{(df["Joining Route"] == "YR").sum()}</span>名（<span id="join_data4">{((df["Joining Route"] == "YR") & (df["Status"] == 1.0)).sum()}</span>名）</p>
    <p>◆メンバー・ルート加入者数（在籍者数）：<span id="join_data5">{(df["Joining Route"] == "FR").sum()}</span>名（<span id="join_data6">{((df["Joining Route"] == "FR") & (df["Status"] == 1.0)).sum()}</span>名）</p>
    <p>◇入団希望加入者数（在籍者数）：<span id="join_data7">{(df["Joining Route"] == "RR").sum()}</span>名（<span id="join_data8">{((df["Joining Route"] == "RR") & (df["Status"] == 1.0)).sum()}</span>名）</p>
    <p>◆体験のみ加入者数：<span id="join_data9">{(df["Joining Route"] == "EO").sum()}</span>名</p>
    <p>◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇</small></small></p>

    </body>
    </html>
    ''')

messagebox.showinfo("実行結果", "完了(*★ω☆*)!!。")  # ポップアップメッセージを表示

xl.Quit()
sys.exit()
