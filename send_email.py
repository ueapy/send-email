# -*- coding: utf-8 -*-
"""
Create email and send it
"""
#source: http://www.experts-exchange.com/questions/28654155/How-to-send-mail-with-office-365-mail-in-python.html
import argparse
import configparser as ConfigParser
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.utils import COMMASPACE, formatdate
from email.encoders import encode_base64
import smtplib
import os

parser = argparse.ArgumentParser(os.path.basename(__file__),
                                 description=__doc__,
                                 formatter_class=argparse.RawTextHelpFormatter)
parser.add_argument('-s', '--send_from_file', default='sender.txt', type=str,
                          help='file name with a sender email')
parser.add_argument('-S', '--send_from', nargs='*', type=str, help='email to send from')
parser.add_argument('-r', '--recipients_file', default='recipients.txt', type=str,
                          help='file name with a list of recipients')
parser.add_argument('-R', '--recipients', nargs='*', type=str, help='email(s) to send to')
parser.add_argument('-t', '--subject', type=str, help='message subject')
parser.add_argument('-b', '--body_file', default='message_body.txt', type=str, help='text file with message body')
parser.add_argument('-B', '--body', nargs='*', type=str, help='message body')
parser.add_argument('-n', '--sign', default='', type=str, help='name to sign')
parser.add_argument('-f', '--files', nargs='*', default=None, help='files to attach')
parser.add_argument('-i', '--images', nargs='*', default=None, help='images to attach')
parser.add_argument('-u', '--username', type=str, help='username')
parser.add_argument('-d', '--debug', action='store_true', default=False, help='print debug message')

def send_email(send_from, send_to, msg_sub, msg_body, files=None,
                          data_attachments=None, server='smtp.office365.com', port=587,
                          tls=True, html=False, images=None,
                          username=None, password=None,
                          config_file=None, config=None):

    if files is None:
        files = []

    if images is None:
        images = []

    if data_attachments is None:
        data_attachments = []

    if config_file is not None:
        config = ConfigParser.ConfigParser()
        config.read(config_file)

    if config is not None:
        server = config.get('smtp', 'server')
        port = config.get('smtp', 'port')
        tls = config.get('smtp', 'tls').lower() in ('true', 'yes', 'y')
        username = config.get('smtp', 'username')
        password = config.get('smtp', 'password')

    msg = MIMEMultipart('related')
    msg['From'] = send_from
    msg['To'] = send_to if isinstance(send_to, str) else COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = msg_sub

    msg.attach( MIMEText(msg_body, 'html' if html else 'plain') )

    for f in files:
        part = MIMEBase('application', "octet-stream")
        part.set_payload( open(f,"rb").read() )
        encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(f))
        msg.attach(part)

    for f in data_attachments:
        part = MIMEBase('application', "octet-stream")
        part.set_payload( f['data'] )
        encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="%s"' % f['filename'])
        msg.attach(part)

    for (n, i) in enumerate(images):
        fp = open(i, 'rb')
        msgImage = MIMEImage(fp.read())
        fp.close()
        msgImage.add_header('Content-ID', '<image{0}>'.format(str(n+1)))
        msg.attach(msgImage)

    smtp = smtplib.SMTP(server, int(port))
    if tls:
        smtp.starttls()

    if username is not None:
        smtp.login(username, password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.close()

def get_flist(filenames):
    if filenames:
        if isinstance(filenames, list):
            files = []
            for i in filenames:
                files += glob(i)
            files = sorted(set(files))
        elif isinstance(filenames, str):
            files = list(filenames)
    else:
        files = None
    return files


if __name__ == '__main__':
    from getpass import getpass
    from glob import glob

    myargs = dict()
    args = parser.parse_args()

    #debug
    #for i in vars(args):
    #    print(i, getattr(args,i), type(getattr(args,i)))
    if args.send_from:
        myargs['send_from'] = args.send_from
    elif args.send_from_file:
        sff = args.send_from_file
        if os.path.isfile(sff):
            with open(sff) as f:
                for l in f.readlines():
                    if not l.startswith('#'):
                        myargs['send_from'] = l.rstrip('\n')
                        break # only one sender
        else:
            raise FileNotFoundError('file {} with sender email not found'.format(sff))
    #else:
    #    parser.error('\nNo senders specified, add --send_from or --send_from_file')

    if args.recipients:
        myargs['send_to'] = args.recipients
    elif args.recipients_file:
        rff = args.recipients_file
        if os.path.isfile(rff):
            with open(rff) as f:
                myargs['send_to'] = []
                for l in f.readlines():
                    if not l.startswith('#'):
                        myargs['send_to'].append(l.rstrip('\n'))
        else:
            raise FileNotFoundError('file {} with recipients emails not found'.format(rff))
    #else:
    #    parser.error('\nNo recipients specified, add --recipients or --recipients_from_file')

    myargs['msg_sub'] = '[python-uea]'
    if args.subject:
        myargs['msg_sub'] = ' '.join([myargs['msg_sub'], args.subject])

    if args.body_file:
        with open(args.body_file) as f:
            myargs['msg_body'] = f.read()
    else:
        myargs['msg_body'] = """Hi pythonistas,
{body}

{sign}
-----------------------------------
On behalf of Python Users Group UEA
website: http://ueapy.github.io
github: http://github.com/ueapy
-----------------------------------""".format(sign=args.sign, body=args.body)

    myargs['files'] = get_flist(args.files)
    myargs['images'] = get_flist(args.images)

    myargs['server'] = 'outlook.office365.com'
    if args.username:
        myargs['username'] = args.username
    else:
        myargs['username'] = os.getenv('USER') + '@uea.ac.uk'


    if args.debug:
        debug_msg = """
Username: {username}
From: {send_from}
To: {send_to}
Subject: {msg_sub}

Body:
{msg_body}

Attachments:
    - files: {files}
    - images: {images}
    """.format(**myargs)
        print(debug_msg)
    else:
        password = getpass()
        print('Sending email...')
        send_email(password=password, **myargs)
