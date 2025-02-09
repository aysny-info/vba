import os

# 変更したいディレクトリのパス
# target_directory = r'\\AFnewT320-kyoyu\社内共有\個人フォルダ\笠間\test_namechange'
target_directory = r'\\Afnewt320-kyoyu\社内共有\AFSKS\ピッキング表\コープ事前入力csv\コープデリ'

# ディレクトリ内のすべてのファイルを再帰的に検索し、ファイル名を変更する関数
def rename_csv_files(directory):
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.csv') and '出荷日' in file:
                # 変更前と変更後のファイル名を設定
                old_file_path = os.path.join(root, file)
                new_file_name = file.replace('出荷日', '製造日')
                new_file_path = os.path.join(root, new_file_name)
                
                # ファイル名を変更
                os.rename(old_file_path, new_file_path)
                print(f"Renamed: {old_file_path} -> {new_file_path}")

# 関数を実行
rename_csv_files(target_directory)
