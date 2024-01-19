import xlrd
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

#定义全局变量
nCols=4     # input表列数
colToAddr=0 #收件人地址所在列数
colToName=1 #收件人姓名所在列数
colSubject=2#主题所在列数
colText=3   #正文文本所在列数
colAttach=4 #附件名所在列数

def SendMail(NewMail):
    print('正发送给',NewMail[colToName],NewMail[colToAddr],'...')
    try:
        #读取本封邮件信息
        toName=NewMail[colToName]
        toAddr=NewMail[colToAddr]
        Subject=NewMail[colSubject]
        Text=NewMail[colText]
        Attach=NewMail[colAttach]
        #构造邮件对象
        Mail=MIMEMultipart('alternative')
        #编辑邮件头
        Mail['Subject']=Header(Subject,'utf-8')
        Mail['From']=Header(FromName,'utf-8')
        Mail['From'].append(fromAddr,'us-ascii')
        Mail['To']=Header(toName,'utf-8')
        Mail['To'].append(toAddr,'us-ascii')
        #构造邮件正文文本
        msgText=MIMEText(Text,'plain','utf-8')
        Mail.attach(msgText)
        #构造邮件附件
        msgAttach=MIMEBase('application','octet-stream')
        fp=open('attachments\\'+Attach,'rb')
        msgAttach.set_payload(fp.read())
        fp.close()
        msgAttach.add_header('Content-Disposition','attachment',filename=('gbk','',Attach))
        encoders.encode_base64(msgAttach)
        Mail.attach(msgAttach)
        #发送邮件
        server.sendmail(fromAddr,toAddr,Mail.as_string())
        print('发送成功')
    except smtplib.SMTPException as e:
        print('发送失败')
        print('异常信息：',e)
#读取my_mail.xls中的邮箱设置信息
try:
    my_mail=xlrd.open_workbook('my_mail.xls')[0]  #打开input表
except Exception as e:
    print('无法打开my_mail.xls文件')
    print('异常信息：',e)
    exit()
smtp_host=my_mail.cell(0,1).value   #服务器
FromName=my_mail.cell(1,1).value
fromAddr=my_mail.cell(2,1).value    #发件人邮箱地址
autho_code=my_mail.cell(3,1).value  #授权码

#读取input.xlsx文件中的邮件信息
try:
    excel=xlrd.open_workbook('input.xls')  #打开input表
except Exception as e:
    print('无法打开input.xls文件')
    print('异常信息：',e)
    exit()
worksheet=excel.sheets()[0]
nRows=worksheet.nrows
prompts=[]
for i in range(1,nRows):
    NewMail=worksheet.row_values(i)
    prompts.append(NewMail)

#初始化服务器
try:
    server=smtplib.SMTP_SSL(smtp_host,465)  #   实例化服务器
    server.connect(smtp_host,465)           #   连接服务器 
    server.login(fromAddr,autho_code)       #   登录邮箱
    server.set_debuglevel(1)
    print('服务器初始化成功')
    for i in range(0,len(prompts)):
        NewMail=prompts[i]
        SendMail(NewMail)
    server.quit()
except smtplib.SMTPException as e:
    print('服务器初始化失败')
    print('异常信息：',e)
    exit()
