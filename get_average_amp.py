## Excelファイルから１周期あたりの平均振幅を計算するプログラム

import pandas as pd

##　データ読み込み＆計算を行う関数
def calculate_average_amplitude(file_path, frequency):
    # Excelファイルを読み込む
    df = pd.read_excel(file_path, engine='openpyxl')

    # 周波数に基づいてサンプル数を計算
    sample_count = int(1 / frequency * 10)  # 1周期あたりのセル数

    # 平均振幅を計算
    results = []
    for col in df.columns[0:7]:  # 最初の7つの列について処理
        values = df[col].values
        subsets = [values[i:i+sample_count] for i in range(0, len(values), sample_count)]  # サンプル数ごとにサブセットに分割
        amplitudes = [subset.max() - subset.min() for subset in subsets]  # 各サブセットの振幅を計算
        average_amplitude = sum(amplitudes) / len(amplitudes)  # 振幅の平均値を計算
        results.append(average_amplitude)

    return results

## 実行部分
# pathの指定
file_path = 'Excelファイルの絶対パスを指定'
# 周波数を指定
frequency = float(input('周波数を入力してください（単位: Hz）: '))
# 計算
results = calculate_average_amplitude(file_path, frequency)

# 結果をテキストファイルに出力
output_file = f'output_{frequency}Hz.txt'  # 周波数をファイル名に含める
with open(output_file, 'w') as f:
    f.write(f'周波数: {frequency} Hz\n')
    for col_index, average_amplitude in enumerate(results):
        col_name = chr(ord('A') + col_index)  # 列名の取得
        line = f'列{col_name}の1周期あたりの平均振幅: {average_amplitude}\n'
        f.write(line)
        print(line, end='')

print(f'結果を {output_file} に記録しました。')

