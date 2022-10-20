#done: send email with attachments - dynamic list, send emails to cc recipients - dynamic list, content can be HTML encoded
#only demonstrate my method
#you have to modify this code to make it work. First you have to register an app in MS Azure and add delegated permission: Mail.Send, Mail.ReadWrite
#then paste app ID in self.APP_ID
#If it is not working please check which accounts can use your registered app, maybe your type of account is not on the list
#You have to install https://pypi.org/project/msgraph-core/
#This is a small part of program modified to make it works so probably you have to adjust it for your purposes, class is not neccessery here
#Maybe I will develop this code to make tool in the future but I do not see universal way to implement it.

import os
import base64
import json
from azure.identity import DeviceCodeCredential, ClientSecretCredential, UsernamePasswordCredential
from msgraph.core import GraphClient
import logging
import ctypes
import getpass


class Send_email_by_MS_Graph(object):
    def __init__(self, user_password, to_recipient, subject, body, cc_recipient: list=None, attachments: list=None):
        self.log = logging.getLogger()
        self.log.info("Starting MS_GRAPH")
        self.APP_ID = '0e23213e-9fsd-4f323-9cds-2221213fa'  #YOU HAVE TO CHANGE THIS -- IT IS RANDOM NUMBER AND YOU NEED TO PASTE HERE YOUR APP ID
        self.SCOPES = ['Mail.Send', 'Mail.ReadWrite']
        self.user_email = input("Enter email (login to MS_Graph) :")
        self.password = user_password
        self.initialize_graph_for_user_auth()
        self.send_mail(to_recipient, cc_recipient, subject, body, attachments)

    def initialize_graph_for_user_auth(self):
        self.UsernamePasswordCredential = UsernamePasswordCredential(client_id=self.APP_ID, username=self.user_email, password=self.password)
        self.user_client = GraphClient(credential=self.UsernamePasswordCredential, scopes=self.SCOPES)

    def draft_attachment(self, file_path):
        if not os.path.exists(file_path):
            self.log.error('file is not found')
            return

        with open(file_path, 'rb') as upload:
            media_content = base64.b64encode(upload.read())

        data_body = {
            '@odata.type': '#microsoft.graph.fileAttachment',
            'contentBytes': media_content.decode('utf-8'),
            'name': os.path.basename(file_path)
        }
        return data_body

    def send_mail(self, to_recipient: str, cc_recipient: list, subject: str, body: str, attachments: list):

        request_body = {
            'message': {
                'subject': subject,
                'body': {
                    'contentType': 'HTML',
                    'content': body
                },
                'toRecipients': [
                    {
                        'emailAddress': {
                            'address': to_recipient
                        }
                    }
                ]
            }
        }

        if cc_recipient is not None:
            request_body['message']['ccRecipients'] = []
            for i in range(len(cc_recipient)):
                request_body['message']['ccRecipients'].append({'emailAddress': {'address': cc_recipient[i]}})

        if attachments is not None:
            request_body['message']['attachments'] = []
            for i in range(len(attachments)):
                request_body['message']['attachments'].append(self.draft_attachment(attachments[i]))

        request_url = '/me/sendmail'

        response = self.user_client.post(request_url,
                              data=json.dumps(request_body),
                              headers={'Content-Type': 'application/json'})
        if response.status_code == 202:
            #self.log.info("Response 202. Email Sent")
            print("Email Sent")
            return True
        else:
            #self.log.error("Email not sent!")
            #self.log.error(response)
            print("Email not sent")
            return False


def main():
    to_recipient = 'write.recipient@email.addres'
    #cc_recipients = ['', '']
    content = r"My mom always said life was like a box of <b>chocolates.</b> <i>You never know what you're gonna get.</i>"
    #attachments = ['']
    password = getpass.getpass(prompt='Password: ', stream=None)
    Send_email_by_MS_Graph(user_password=password, to_recipient=to_recipient, subject='Run Forest!', body=content)

if __name__ == "__main__":
    main()




