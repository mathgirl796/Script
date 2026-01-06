import smtplib
import mimetypes
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from email import encoders
from email.header import Header



class EmailSender:
    def __init__(self, sender_email, sender_password, sender_name, 
                 smtp_server, smtp_port):
        """
        初始化邮件发送器
        :param sender_email: 发件人邮箱
        :param sender_password: 发件人密码（或应用专用密码）
        :param sender_name: 发件人显示名称（可选）
        :param smtp_server: SMTP服务器
        :param smtp_port: SMTP端口
        """
        self.sender_email = sender_email
        self.sender_password = sender_password
        self.sender_name = sender_name if sender_name else sender_email.split('@')[0]
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port

    def _create_email(self, receiver_email, subject, body, attachments, receiver_name=None):
        """创建单封邮件对象（内部方法）"""
        msg = MIMEMultipart()
        # 使用标准的邮件格式编码发件人信息
        from email.utils import formataddr
        msg['From'] = formataddr((self.sender_name, self.sender_email))
        if receiver_name:
            msg['To'] = formataddr((receiver_name, receiver_email))
        else:
            msg['To'] = Header(receiver_email)
        msg['Subject'] = Header(subject)

        # 添加正文
        msg.attach(MIMEText(body, 'plain'))

        # 添加附件
        for path in attachments:
            try:
                with open(path, 'rb') as f:
                    file_data = f.read()
                
                filename = path.split('/')[-1].split('\\')[-1]  # 兼容不同系统路径
                
                # 根据文件扩展名确定MIME类型
                ctype, encoding = mimetypes.guess_type(path)
                if ctype is None or encoding is not None:
                    ctype = 'application/octet-stream'
                
                maintype, subtype = ctype.split('/', 1)
                
                if maintype == 'text':
                    # 文本文件需要解码为字符串
                    try:
                        text_content = file_data.decode('utf-8')
                        part = MIMEText(text_content, _subtype=subtype)
                    except UnicodeDecodeError:
                        # 如果UTF-8解码失败，尝试其他编码
                        try:
                            text_content = file_data.decode('gbk')
                            part = MIMEText(text_content, _subtype=subtype)
                        except UnicodeDecodeError:
                            # 如果都失败，当作二进制文件处理
                            part = MIMEBase('application', 'octet-stream')
                            part.set_payload(file_data)
                            encoders.encode_base64(part)
                elif maintype == 'image':
                    from email.mime.image import MIMEImage
                    part = MIMEImage(file_data, _subtype=subtype)
                elif maintype == 'audio':
                    from email.mime.audio import MIMEAudio
                    part = MIMEAudio(file_data, _subtype=subtype)
                elif maintype == 'application':
                    part = MIMEApplication(file_data, _subtype=subtype)
                else:
                    # 其他类型使用MIMEBase
                    part = MIMEBase(maintype, subtype)
                    part.set_payload(file_data)
                    encoders.encode_base64(part)
                
                # 设置文件名，确保正确显示原始文件名
                # 使用标准的Content-Disposition头设置
                part.add_header('Content-Disposition', f'attachment', filename=filename)
                msg.attach(part)
                
            except Exception as e:
                raise Exception(f"附件处理失败（{path}）：{str(e)}")

        return msg

    def send_single_email(self, receiver_email, subject, body, attachments=[], receiver_name=None):
        """
        发送单封邮件
        :param receiver_email: 收件人邮箱
        :param subject: 邮件标题
        :param body: 邮件正文
        :param attachments: 附件路径列表（如['file1.txt', 'file2.pdf']，可选）
        :param receiver_name: 收件人名称（可选）
        :return: 发送结果（True/False）和信息
        """
        try:
            # 创建邮件
            msg = self._create_email(receiver_email, subject, body, attachments, receiver_name)
            
            # 发送邮件
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.sender_email, self.sender_password)
                server.sendmail(
                    from_addr=self.sender_email,
                    to_addrs=receiver_email,
                    msg=msg.as_string()
                )
            return True, f"邮件发送成功（收件人：{receiver_email}）"
        except Exception as e:
            return False, f"邮件发送失败（收件人：{receiver_email}）：{str(e)}"

if __name__ == '__main__':
    from ConfigManager import ConfigManager
    config_manager = ConfigManager()
    email_config = config_manager.get_email_sender_config()
    email_sender = EmailSender(**email_config)
    # 发送测试邮件给自己
    receiver_email = 'aduan456@163.com'
    subject = config_manager.get_config('email_subject') or '测试邮件'
    body = config_manager.get_config('email_body') or '这是一封测试邮件'
    success, message = email_sender.send_single_email(
        receiver_email, subject, body, [r'C:\Users\Administrator\Desktop\发邮件\测试.xlsx'],
        receiver_name="粑粑人"
    )
    print(success, message)
    