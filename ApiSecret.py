# -*- coding: UTF-8 -*-
import os
import json
import sys
import random
import time
import logging
import requests as req
#先註冊azure應用,確保應用有以下權限:
#files:	Files.Read.All、Files.ReadWrite.All、Sites.Read.All、Sites.ReadWrite.All
#user:	User.Read.All、User.ReadWrite.All、Directory.Read.All、Directory.ReadWrite.All
#mail:  Mail.Read、Mail.ReadWrite、MailboxSettings.Read、MailboxSettings.ReadWrite
#註冊後一定要再點代表xxx授予管理員同意,否則outlook api無法調用

def gettoken(token_path, client_id, client_secret):
    fo = open(refresh_token_path, "r+", encoding="utf-8")
    refresh_token = fo.read()
    fo.close()

    headers={'Content-Type':'application/x-www-form-urlencoded'}
    data={'grant_type': 'refresh_token',
          'refresh_token': refresh_token,
          'client_id':client_id,
          'client_secret':client_secret,
          'redirect_uri':'http://localhost:53682/'
         }
    logging.debug('gettoken data %s', data)
    html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', data=data,headers=headers)
    jsontxt = json.loads(html.text)
    logging.debug('gettoken result %s', html.text)
    refresh_token = jsontxt['refresh_token']
    ms_access_token = jsontxt['access_token']
    with open(token_path, 'w+', encoding="utf-8") as f:
        f.write(refresh_token)
    return ms_access_token
    
def exec_api(token, url):
    headers={
        'Authorization': token,
        'Content-Type':'application/json'
    }
    if req.get(url, headers=headers, timeout=60).status_code == 200:
        logging.info('調用成功 %s', url)
    else:
        logging.error('調用失敗 %s', url)

if __name__ == '__main__':
    logging.basicConfig(format='%(asctime)s %(levelname)-8s %(message)s', level=logging.DEBUG)

    refresh_token_path=sys.path[0]+r'/refresh_token.txt'
    client_id = os.getenv('CLIENT_ID')
    client_secret = os.getenv('CLIENT_SECRET')

    access_token=gettoken(refresh_token_path, client_id, client_secret)

    api_list = [
            r'https://graph.microsoft.com/v1.0/me/',
            r'https://graph.microsoft.com/v1.0/users',
            r'https://graph.microsoft.com/v1.0/me/people',
            r'https://graph.microsoft.com/v1.0/groups',
            r'https://graph.microsoft.com/v1.0/me/contacts',
            r'https://graph.microsoft.com/v1.0/me/drive/root',
            r'https://graph.microsoft.com/v1.0/me/drive/root/children',
            r'https://graph.microsoft.com/v1.0/drive/root',
            r'https://graph.microsoft.com/v1.0/me/drive',
            r'https://graph.microsoft.com/v1.0/me/drive/recent',
            r'https://graph.microsoft.com/v1.0/me/drive/sharedWithMe',
            r'https://graph.microsoft.com/v1.0/me/calendars',
            r'https://graph.microsoft.com/v1.0/me/events',
            r'https://graph.microsoft.com/v1.0/sites/root',
            r'https://graph.microsoft.com/v1.0/sites/root/sites',
            r'https://graph.microsoft.com/v1.0/sites/root/drives',
            r'https://graph.microsoft.com/v1.0/sites/root/columns',
            r'https://graph.microsoft.com/v1.0/me/onenote/notebooks',
            r'https://graph.microsoft.com/v1.0/me/onenote/sections',
            r'https://graph.microsoft.com/v1.0/me/onenote/pages',
            r'https://graph.microsoft.com/v1.0/me/messages',
            r'https://graph.microsoft.com/v1.0/me/mailFolders',
            r'https://graph.microsoft.com/v1.0/me/outlook/masterCategories',
            r'https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages/delta',
            r'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',
            r"https://graph.microsoft.com/v1.0/me/messages?$filter=importance eq 'high'",
            r'https://graph.microsoft.com/v1.0/me/messages?$search="hello world"',
            r'https://graph.microsoft.com/beta/me/messages?$select=internetMessageHeaders&$top',
            r'https://api.powerbi.com/v1.0/myorg/apps'
            ]
    for _ in range(random.randint(10,100)):
        random.shuffle(api_list)
        for api in api_list:
            exec_api(access_token, api)
            time.sleep(random.randint(0,5))
        logging.info("完成 %d", _)
