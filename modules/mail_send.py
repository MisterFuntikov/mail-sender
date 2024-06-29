import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


class mailSend():

    SMTPOBJ = None

    def connect(self, host, port=25, user='user', password='', ssl=False) -> dict:
        try:
            self.SMTPOBJ = smtplib.SMTP(host, port, timeout=10000)
            if ssl:
                self.SMTPOBJ.starttls()
            self.SMTPOBJ.login(user, password)
        except Exception as Err:
            return {'status': False, 'msg': str(Err)}
        return {'status': True}

    def close(self) -> dict:
        try:
            self.SMTPOBJ.close()
        except Exception as Err:
            return {'status': False, 'msg': str(Err)}
        return {'status': True}

    def send(self, mfrom='', mto=[], title='', msg='', files=[]) -> dict:

        if mfrom == '':
            return {'status': False, 'msg': 'Нет отправителя.'}
        if len(mto) == 0:
            return {'status': False, 'msg': 'Не задан параметр кому направлено письмо.'}
        if type(files) != list:
            files = []

        try:

            message = MIMEMultipart()
            message['Subject'] = str(title)
            message['From'] = str(mfrom)
            message['To'] = ', '.join(mto)
            message.attach(MIMEText(msg, 'html', 'utf8'))

            if len(files) > 0:
                for f in files:
                    with open(f, 'rb') as file:
                        part = MIMEApplication(
                            file.read(),
                            Name=os.path.basename(f)
                        )
                    part.add_header('Content-Disposition', 'attachment',
                                    filename='%s' % f.split('\\')[-1])
                    message.attach(part)

            self.SMTPOBJ.sendmail(mfrom, mto, message.as_string())

        except Exception as Err:
            return {'status': False, 'msg': str(Err)}
        return {'status': True}

    pass
