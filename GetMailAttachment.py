import poplib
import getpass
import email
import datetime
import time
import os
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr

email_addr = ''
email_pswd = ''
pop_server = ''
files_path = ''
date_limit = ''

def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value

def GetAttachment(msg):
    attachment_files = []
    
    for part in msg.walk():
        file_name = part.get_filename()
        
        if file_name:
            h = email.header.Header(file_name)
            dh = email.header.decode_header(h)
            filename = dh[0][0]
            if dh[0][1]:
                filename = decode_str(str(filename,dh[0][1]))
            data = part.get_payload(decode=True)
            att_file = open(files_path + filename, 'wb')
            attachment_files.append(filename)
            att_file.write(data)
            att_file.close()
    return attachment_files

def GetMail(server):
    resp, mails, octets = server.list()
    # 可以查看返回的列表类似[b'1 82923', b'2 2184', ...]
    index = len(mails)
    for i in range(index, 0, -1):
        resp, lines, octets = server.retr(i)
        msg_content = b'\r\n'.join(lines).decode('utf-8')
        msg = Parser().parsestr(msg_content)
        date1 = time.strptime(msg.get("Date")[0:24], '%a, %d %b %Y %H:%M:%S')
        date = time.strftime("%Y%m%d", date1)
        if date < date_limit:
            break
        subject = decode_str(msg.get('Subject', ''))
        print(subject)
        attachment_files = GetAttachment(msg)
        if attachment_files == []:
            print("No attachment in mail " + str(i))
        else:
            print("Successfully downloaded " + str(attachment_files) + " in mail " + str(i))

def MailLogin():
    server = poplib.POP3_SSL(pop_server)
    server.set_debuglevel(0)
    print(server.getwelcome().decode('utf-8'))
    server.user(email_addr)
    server.pass_(email_pswd)
    print('Message: %s. Size: %s' % server.stat())
    return server

#mail info & log in
email_addr = input("Your Email Address?\n")
pop_server = 'pop.' + email_addr.split('@')[1]
email_pswd = getpass.getpass("Your Email Password?(The password will not be shown on the screen!)\n")
try:
    server = MailLogin()
except poplib.error_proto:
    print("Log in failed")
    print("Your input Address: "+ email_addr + " Password: " + email_pswd)
else:
    print("Log in successfully")

#input the download path
files_path = input("Please input the path to download files:\n")
while os.path.isdir(files_path) == 0:
    print("The path is unavailable.")
    files_path = input("Please input the path to download files again:\n")
print("The path is available.")

#get the limit of date
date_limit = input("Get emails after: (Format like 20180101)\n")

#get mail
GetMail(server)

print("Successfully downloaded attachments!")

server.quit()