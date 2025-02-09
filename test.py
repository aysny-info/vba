from Cryptodome.Cipher import AES
import struct

def decrypt_xlsx(input_file, output_file, password):
    try:
        with open(input_file, "rb") as f:
            file_data = f.read()
        
        # AESの復号化設定
        key = password.encode("utf-8")[:16]  # パスワードを16バイトに切り詰め
        iv = b'\x00' * 16  # 初期ベクトル
        cipher = AES.new(key, AES.MODE_CBC, iv)
        
        decrypted_data = cipher.decrypt(file_data)
        
        with open(output_file, "wb") as f:
            f.write(decrypted_data)
        
        print(f"パスワード解除済みのファイルを '{output_file}' に保存しました。")
    except Exception as e:
        print(f"エラーが発生しました: {e}")

# 使用例
input_file = r"C:\Users\kasama9\Desktop\◆2024期 アイソニーフーズ組織図下期案 20241127VER.xlsx"
output_file = r"C:\Users\kasama9\Desktop\test.xlsx"
password = "your_password"
decrypt_xlsx(input_file, output_file, password)

