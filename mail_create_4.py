import csv
import os
from email.message import EmailMessage
from email.generator import Generator
from date import date
from date import subject
from date import att_date

ma_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'OneDrive\\デスクトップ')

#添付ファイルの有無を入れてもらう。
mail_attach = input("添付ファイルを入れる場合は、「1」を不要の場合は「2」を入れてください。")
mail_attach_int = int(mail_attach)

# CSVファイルを読み込む
with open('address2.csv', 'r') as f:
    # 以下のフォーマットで値をコマンドプロンプトに表示させてみよう
    # 会社名: <会社名>, 宛名: <宛名>
    # リストlist_dataに値を追加して最後にリストをコマンドプロンプトに表示させてみよう
    csv_data = csv.reader(f)

    list_data = []
    for line in csv_data:
        list_data.append(line)

# 変数の設定
# 送信元、宛先、CCのメールアドレスを変数で定義してメールを作成してみよう！
        company = line[2]
        name = line[3]
        from_mail = 'pymee-support@example.com'
        to_mail = line[0]
        cc_mail = line[1]

        mail_path = os.path.join(ma_path,company)
        if not os.path.isdir(mail_path):
            os.makedirs(mail_path)

# テキストファイルから本文を読み込み、<会社名>と<宛名>の部分を関数で書き換えてみよう
        message = date.replace('<会社名>', company).replace('<宛名>', name)

# メールを作成する
        mail_data = EmailMessage()

# 送信元アドレスを設定する
        mail_data['From'] = from_mail

# 宛先アドレスを設定する
        mail_data['To'] = to_mail

# CCアドレスを設定する
        mail_data['CC'] = cc_mail

# テキストファイルから件名を読み込む
        mail_data['subject'] = subject

# 本文を設定する
        mail_data.set_content(message)

# 設定した内容をファイルに書き込む
        with open(os.path.join(mail_path,'サービスメンテナンスのご連絡_<会社名>_<宛名>.eml'.replace('<会社名>',company).replace('<宛名>',name)), 'w') as eml:
            eml_file = Generator(eml)
            eml_file.flatten(mail_data)

        if mail_attach_int == 1:
            # 送信元、宛先、CCのメールアドレスを変数で定義してメールを作成してみよう！
            company = line[2]
            name = line[3]
            from_mail = 'pymee-support@example.com'
            to_mail = line[0]
            cc_mail = line[1]

            mail_path = os.path.join(ma_path, company)
            if not os.path.isdir(mail_path):
                os.makedirs(mail_path)

            # テキストファイルから本文を読み込み、<会社名>と<宛名>の部分を関数で書き換えてみよう
            att_message = att_date.replace('<会社名>', company).replace('<宛名>', name)

            # メールを作成する
            mail_data = EmailMessage()

            # 送信元アドレスを設定する
            mail_data['From'] = from_mail

            # 宛先アドレスを設定する
            mail_data['To'] = to_mail

            # CCアドレスを設定する
            mail_data['CC'] = cc_mail

            # テキストファイルから件名を読み込む
            mail_data['subject'] = subject

            # 本文を設定する
            mail_data.set_content(att_message)

            with open(os.path.join(mail_path,'【PW】サービスメンテナンスのご連絡_<会社名>_<宛名>.eml'.replace('<会社名>', company).replace('<宛名>', name)),'w') as eml:
                eml_file = Generator(eml)
                eml_file.flatten(mail_data)

        elif mail_attach_int == 2:
            pass

print("メールの作成が完了しました")