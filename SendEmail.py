#!/usr/bin/env python3
# -*- coding: utf-8 -*-

' Send Email '

__author__ = 'ZOG'

import os
import sys
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import date, timedelta

# 获取当天日期
today = date.today().strftime("%Y-%m-%d")
# 获取昨天日期
# today = (date.today() - timedelta(days=1)).strftime("%Y-%m-%d")

# 检查附件文件是否存在
filename = f"\\\\附件文件地址\\{today}.7z"
if not os.path.exists(filename):
    print("文件不存在，程序将退出。")
    sys.exit()  # 文件不存在，退出程序

# 配置SMTP服务器凭据
username = "发件人邮箱地址"
password = "发件人邮箱密码"
smtp_server = "smtp服务器地址"
smtp_port = 25

# 发件人、收件人和主题信息
from_email = "发件人邮箱地址"
to_emails = ["收件人地址1", "收件人地址2"]  # 收件人列表
# 单个收件地址时使用，注释上一行
# to_emails = "邮件地址"
# 邮件标题
subject = f"{today}_RFID 生産LOG"

# 创建一个MIMEMultipart对象
msg = MIMEMultipart()

# 设置发件人和主题
msg["From"] = from_email
msg["Subject"] = subject
msg["To"] = ", ".join(to_emails)  # 将收件人列表转换为字符串
# 单个邮件地址时使用，注释上一行
# msg["To"] = to_emails

# 添加邮件内容
body = f"""​
邮件正文内容
"""
msg.attach(MIMEText(body, "plain"))

# 读取附件文件
filename = f"\\\\附件文件地址\\{today}.7z"
attachment = open(filename, "rb")

# 创建一个MIMEBase对象
mime_base = MIMEBase("application", "octet-stream")

# 设置附件头
mime_base.set_payload(attachment.read())
encoders.encode_base64(mime_base)
mime_base.add_header("Content-Disposition", "attachment", filename=f"{today}.7z")
mime_base.add_header("Content-Type", "application/x-7z-compressed")

# 添加附件到邮件
msg.attach(mime_base)
attachment.close()  # 关闭附件文件

# 连接到SMTP服务器
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(username, password)

# 发送邮件
server.sendmail(from_email, to_emails, msg.as_string())  # 注意这里把to_emails作为列表传入

# 关闭连接
server.quit()
