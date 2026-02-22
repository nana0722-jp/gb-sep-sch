import pandas as pd
import json

# あなたが見つけてくれたExcelの直URL
excel_url = "https://www.city.asaka.lg.jp/uploaded/attachment/98121.xlsx"

def convert_excel_to_json():
    print(f"Excelをダウンロードして解析中: {excel_url}")
    
    try:
        # 1. Excelを読み込む (1行目がヘッダー)
        # B列(頭文字)やC列(品名)があることを前提に読み込み
        df = pd.read_excel(excel_url, header=0)

        # 2. 列名の整理 (C列以降が品名・分別区分・出し方と想定)
        # Excelの実際の列名に合わせて適宜修正してください
        # 0=番号, 1=頭文字, 2=品名, 3=分別区分, 4=出し方 と仮定
        
        # 必要な列だけを抽出して名前を付け直す
        # columns[2]がC列(品名), [3]がD列(分別), [4]がE列(出し方)
        df_clean = df.iloc[:, [2, 3, 4]].copy()
        df_clean.columns = ['name', 'category', 'note']

        # 3. 空の行（品名が入っていない行）を除去



        df_clean = df_clean.dropna(subset=['name'])

        # --- ここを追加：空欄(NaN)を空文字("")に変換 ---
        df_clean = df_clean.fillna("")
        # --------------------------------------------

        # 4. JSONに変換
        data_list = df_clean.to_dict(orient='records')
        
        with open('asaka_gomi.json', 'w', encoding='utf-8') as f:
            json.dump(data_list, f, ensure_ascii=False, indent=4)
        
        print(f"成功！ {len(data_list)} 件のデータを保存しました。")

    except Exception as e:
        print(f"エラーが発生しました: {e}")
        print("※ pandas と openpyxl がインストールされているか確認してください。")

if __name__ == "__main__":
    convert_excel_to_json()