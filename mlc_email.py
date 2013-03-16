#
# mlc_email.py
# Copyright (c) 2012-2013 Aaron C Spike
# 
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
# copies of the Software, and to permit persons to whom the Software is 
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in 
# all copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, 
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
# 

import win32com.client
from win32com.server.exception import COMException
import winerror

import smtplib
import os

from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText

from email.utils import COMMASPACE, formatdate
from email import encoders

import email.generator
import email.iterators


class Email(object):
    _reg_progid_ = 'MLC.Email'
    _reg_verprogid_ = 'MLC.Email.1'
    _reg_clsid_ = '{88A05252-816B-4214-AC83-05190632D498}'
    _public_methods_ = ['setFrom','setSubject','addTo','addCc','addBcc','setHeader','addText','addFile','setServer','send']
    _public_attrs_ = []
    _readonly_attrs_ = []
    def __init__(self):
        self.From = ''
        self.Subject = ''
        self.To = set([])
        self.Cc = set([])
        self.Bcc = set([])
        self.Headers = {}
        self.Body = []
        self.Files = {}
        self.Server = 'localhost'
    def setFrom(self, addr):
        addr = str(addr)
        self.From = addr
    def setSubject(self, subj):
        self.Subject = str(subj)
    def addTo(self, addr):
        addr = str(addr)
        self.To.add(addr)
    def addCc(self, addr):
        addr = str(addr)
        self.Cc.add(addr)
    def addBcc(self, addr):
        addr = str(addr)
        self.Bcc.add(addr)
    def setHeader(self, name, val):
        self.Headers[str(name)] = str(val)
    def addText(self, text):
        self.Body.append(str(text))
    def addFile(self, filepath):
        try:
            filepath = str(filepath)
            fname = os.path.basename(filepath)
            fdata = open(filepath, 'rb').read()
            self.Files[fname] = fdata
        except:
            raise COMException('Unable to attach file: ' + str(filepath), winerror.E_FAIL)
    def setServer(self, server):
        self.Server = str(server)
    def send(self):
        msg = MIMEMultipart()
        try:
            msg['From'] = self.From
            msg['To'] = COMMASPACE.join(list(self.To))
            msg['Date'] = formatdate(localtime=True)
            msg['Subject'] = self.Subject
            if self.Cc:
                msg['Cc'] = COMMASPACE.join(list(self.Cc))
            
            for h in self.Headers:
                msg[h] = self.Headers[h]
            
            msg.attach(MIMEText('\n'.join(self.Body)))
            
            for f in self.Files.keys():
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(self.Files[f])
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment; filename="%s"' % (f,))
                msg.attach(part)
            
        except:
            raise COMException('Unable to build message', winerror.E_FAIL)
        
        try:
            envelopeAddresses = list(self.To | self.Cc | self.Bcc)
            smtp = smtplib.SMTP(self.Server)
            smtp.sendmail(self.From, envelopeAddresses, msg.as_string())
            smtp.close()
        except:
            raise COMException('Unable to send message', winerror.E_FAIL)
    
if __name__ == '__main__':
    import win32com.server.register
    win32com.server.register.UseCommandLine(Email)
