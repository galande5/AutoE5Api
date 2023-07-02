# -*- coding: UTF-8 -*-
import os
import requests as req
import json,sys,time,random

app_num=os.getenv('APP_NUM')
if app_num == '':
    app_num = '1'
access_token_list=['wangziyingwen']*int(app_num)
###########################
# config選項說明
# 0：關閉  ， 1：開啟
# api_rand：是否隨機排序api （開啟隨機抽取12個，關閉默認初版10個）。默認1開啟
# rounds: 輪數，即每次啟動跑幾輪。
# rounds_delay: 是否開啟每輪之間的隨機延時，後面兩參數代表延時的區間。默認0關閉
# api_delay: 是否開啟api之間的延時，默認0關閉
# app_delay: 是否開啟帳號之間的延時，默認0關閉
########################################
config = {
         'api_rand': 1,
         'rounds': 3,
         'rounds_delay': [0,60,120],
         'api_delay': [0,2,6],
         'app_delay': [0,30,60],
         }
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
           ]

#微軟refresh_token獲取
def getmstoken(ms_token,appnum):
    headers={'Content-Type':'application/x-www-form-urlencoded'
            }
    data={'grant_type': 'refresh_token',
        'refresh_token': ms_token,
        'client_id':client_id,
        'client_secret':client_secret,
        'redirect_uri':'http://localhost:53682/'
        }
    html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',data=data,headers=headers)
    jsontxt = json.loads(html.text)
    if 'refresh_token' in jsontxt:
        print(r'帳號/應用 '+str(appnum)+' 的微軟金鑰獲取成功')
    else:
        print(r'帳號/應用 '+str(appnum)+' 的微軟金鑰獲取失敗'+'\n'+'請檢查secret裡 CLIENT_ID , CLIENT_SECRET , MS_TOKEN 格式與內容是否正確，然後重新設定')
    refresh_token = jsontxt['refresh_token']
    access_token = jsontxt['access_token']
    return access_token

#呼叫函數
def runapi(apilist,a):
    localtime = time.asctime( time.localtime(time.time()) )
    access_token=access_token_list[a-1]
    headers={
            'Authorization': 'bearer ' + access_token,
            'Content-Type': 'application/json'
            }
    for b in range(len(apilist)):	
        if req.get(api_list[apilist[b]],headers=headers).status_code == 200:
            print('第'+str(apilist[b])+"號api呼叫成功")
        else:
            print("pass")
        if config['api_delay'][0] == 1:
            time.sleep(random.randint(config['api_delay'][1],config['api_delay'][2]))

#一次性獲取access_token，降低獲取率
for a in range(1, int(app_num)+1):
    client_id=os.getenv('CLIENT_ID_'+str(a))
    client_secret=os.getenv('CLIENT_SECRET_'+str(a))
    ms_token=os.getenv('MS_TOKEN_'+str(a))
    access_token_list[a-1]=getmstoken(ms_token,a)

#隨機api序列
fixed_api=[0,1,5,6,20,21]
#保證抽取到outlook,onedrive的api
ex_api=[2,3,4,7,8,9,10,22,23,24,25,26,27,13,14,15,16,17,18,19,11,12]
#額外抽取填充的api
fixed_api.extend(random.sample(ex_api,6))
random.shuffle(fixed_api)
final_list=fixed_api

#實際運行
if int(app_num) > 1:
    print('多帳戶/應用模式下，日誌報告裡可能會出現一堆***，屬於正常情況')
print("如果api數量少於規定值，則是api賦權沒有弄好，或者是onedrive還沒有初始化成功。前者請重新賦權並獲取微軟金鑰替換，後者請稍等幾天")
print('共 '+str(app_num)+r' 帳號/應用，'+r'每個帳號/應用 '+str(config['rounds'])+' 輪') 
for r in range(1,config['rounds']+1):
    if config['rounds_delay'][0] == 1:
        time.sleep(random.randint(config['rounds_delay'][1],config['rounds_delay'][2]))		
    for a in range(1, int(app_num)+1):
        if config['app_delay'][0] == 1:
            time.sleep(random.randint(config['app_delay'][1],config['app_delay'][2]))
        client_id=os.getenv('CLIENT_ID_'+str(a))
        client_secret=os.getenv('CLIENT_SECRET_'+str(a))
        print('\n'+'應用/帳號 '+str(a)+' 的第'+str(r)+'輪 '+time.asctime(time.localtime(time.time()))+'\n')
        if config['api_rand'] == 1:
            print("已開啟隨機順序,共十二個api,自己數")
            apilist=final_list
        else:
            print("原版順序,共十個api,自己數")
            apilist=[5,9,8,1,20,24,23,6,21,22]
        runapi(apilist,a)