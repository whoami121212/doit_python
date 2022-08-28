import smtplib
import os
import json
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

class mailbot:
    def __init__(self):
        '''
        email_info.json에 들어있는 로그인 정보를 이용해서 mailbot 생성
        로그인 후 self.bot에 할당
        '''
        self.info = self.load_info()
        self.bot = self.log_in(self.info)

    def load_info(self):
        info_src = os.path.join(os.getcwd(),'email_info.json')
        with open(info_src,'r', encoding="UTF-8") as f:
            info = json.load(f)['info']
        return info
        
    def log_in(self, info):
        s = smtplib.SMTP_SSL('smtp.gmail.com')
        print(info['From'])
        print(info['Password'])
        s.login(info['From'], info['Password'])
        return s

    def send_mail(self, s, info, subject, content, to_email, path=None):
        '''
        
        '''
        # msg = MIMEMultipart('alternative')
        msg = MIMEMultipart()
        msg['Subject'] = subject
        msg['From'] = self.info['From']
        msg['To'] = to_email

        part2 = MIMEText(content, 'html')
        msg.attach(part2)

        part = MIMEBase('application', "octet-stream")
        if path is not None:
            with open(path, 'rb') as file:
                part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', "attachment", filename=os.path.basename(path))
            msg.attach(part)
        
        s.sendmail(info['From'], to_email, msg.as_string())

    def log_out(self, s):
        s.quit()


