#メール本文用テキスト
main_bun = open('mail_honbun.txt','r',encoding='utf-8')
date = main_bun.read()
main_bun.close()

#件名用テキスト
mail_subject = open('subject.txt','r',encoding = 'utf-8')
subject = mail_subject.read()
mail_subject.close()

#添付ファイル用本文テキスト
main_att = open('mail_attach.txt','r',encoding='utf-8')
att_date = main_att.read()
main_att.close()



