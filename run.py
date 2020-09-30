import win32com.client
import os
import sys

if getattr(sys, 'frozen', False):
    dirname = os.path.normpath((os.path.dirname(sys.executable)))
elif __file__:
    dirname = os.path.normpath(os.path.dirname(__file__))

mail_path = os.path.join(dirname, 'mails')

if not os.path.exists(mail_path):
    input('Creating directory called mails. Put .msg files here and run this program again. Press any key to close')
    
    os.makedirs(mail_path)

else:
    reletive_files = os.listdir(mail_path)
    files = [os.path.join(mail_path, f) for f in reletive_files]

    print(files)
    for file in files:
        if file.endswith(".msg"):
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            msg = outlook.OpenSharedItem(file)
            att = msg.Attachments
            for i in att:
                attachment_path = os.path.join(dirname, 'attachments')
                if not os.path.exists(attachment_path):
                    os.makedirs(attachment_path)

                i.SaveAsFile(os.path.join(attachment_path, i.FileName))