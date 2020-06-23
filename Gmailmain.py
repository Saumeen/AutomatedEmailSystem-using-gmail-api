from __future__ import print_function
import pickle
import os.path
import os
import datetime
import base64
import email
import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from apiclient import errors
from requests_oauthlib.oauth1_auth import unicode


SCOPES = ['https://www.googleapis.com/auth/gmail.modify']



def ListMessagesMatchingQuery(service, user_id, query=''):
  try:
    response = service.users().messages().list(userId=user_id,
                                               q=query).execute()
    messages = []
    if 'messages' in response:
      messages.extend(response['messages'])

    while 'nextPageToken' in response:
      page_token = response['nextPageToken']
      response = service.users().messages().list(userId=user_id, q=query,
                                         pageToken=page_token).execute()
      messages.extend(response['messages'])

    return messages
  except errors.HttpError:
    print( 'An error occurred: %s')

##############################################################################
def GetMessage(service, user_id, msg_id):
  try:
    message = service.users().messages().get(userId=user_id, id=msg_id).execute()
    return message
  except errors.HttpError:
    print ('An error occurred: %s')

##############################################################################
def GetAttachments(service, user_id, msg_id, store_dir):
  try:
    message = service.users().messages().get(userId=user_id, id=msg_id).execute()
    for part in message['payload']['parts']:
      if part['filename']:
        print("attachmnet is heres")
        attachment = service.users().messages().attachments().get(userId='me', messageId=message['id'], id=part['body']['attachmentId']).execute()
        file_data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))
        path = ''.join([store_dir, part['filename']])
        f = open(path,'wb')
        f.write(file_data)
        f.close()
  except errors.HttpError:
    print ('An error occurred in messaage attactment' )

###############################################################################
def Convert(tup, di): 
    for a, b in tup: 
        di.setdefault(a, []).append(b) 
    return di 
def check(dlist):
    count = 0
    temp = 0
    max_key, max_value = max(dlist.items(), key = lambda x: len(set(x[1])))
    for i in dlist:
        
        if len(dlist[i]) != len(max_value):
            count =  max(len(max_value) - len(dlist[i]),temp) 
        temp = count
    print(count)
    return count

def ExportToexcel(msg,compnyname,path):
    d = dict()
    dlist={}
    for k in msg:
        lines=k.split('\n')
        lines_capital_char = [x.upper() for x in lines]
        res=[]
        for line in lines_capital_char:
            if ':' in  line:
                temp = tuple(line.split(':'))
                k = tuple(a.strip() for a in temp)
                res.append(k)
        dlist = Convert(res,d)
    if dlist:
        print(dlist)
        count =check(dlist) 
        print('In export function')
        if count > 0:
            
            df = pd.DataFrame(dlist['ID'],columns=['ID'])
        else:
            df= pd.DataFrame.from_dict(dlist)
        print(df)
        if os.path.exists(path+compnyname+'/'+compnyname+'.xlsx'):
            os.remove(path+compnyname+'/'+compnyname+'.xlsx')

        df.to_excel(path+compnyname+'/'+compnyname+'.xlsx', engine='xlsxwriter')
        print("export excel")
        return count 

#################################FIlename == file.xlsx###############################################
def FinalSheet(filename,compnyname,database,path,idfield):
    from datetime import datetime
    db = pd.read_excel(path+database,sheet_name='Sheet1')
    idf = pd.read_excel(path+compnyname+'/'+filename,sheet_name='Sheet1')
    ids = []
    for SR,t in idf.iterrows():
        ids.append(t.ID.strip())
    out = db.loc[db[idfield].isin(ids)]
    if os.path.exists(path+compnyname+'/'+compnyname+'fullsheet'+'.xlsx'):
         os.remove(path+compnyname+'/'+compnyname+'fullsheet'+'.xlsx')
    out.to_excel(path+compnyname+'/'+compnyname+'fullsheet'+ '.xlsx')
#################################################################################

def checkemail(compnyname,database,path):
    data = pd.read_excel(path+compnyname+'/'+compnyname+'fullsheet.xlsx',sheet_name='Sheet1')
    email = pd.read_excel(path+compnyname+'/'+'Email.xlsx',sheet_name='Sheet1')
    db = pd.read_excel(path+database,sheet_name='Sheet1')

    presentemails=[]
    for SR,t in data.iterrows():
        presentemails.append(t.Email.strip())
    print(presentemails)
    misssingdata=[]
    for i in email['Email']:
        if i not in presentemails:
            misssingdata.append(i)
    print(misssingdata)
    out = db.loc[db['Email'].isin(misssingdata)]

    out.to_excel(path+compnyname+'/'+'errorinemail.xlsx',sheet_name='Sheet1')
##################################################################################
def ExportEmail(fromemaillist,compny_name_from_query,path):
    emaillistfinal=[]
    
    for s in fromemaillist:
        emaillistfinal.append(s[s.find("<")+1:s.find(">")])
    if fromemaillist:
        df = pd.DataFrame(emaillistfinal,columns=['Email'])
        df.to_excel(path+compny_name_from_query+'/Email.xlsx')
##################################################################################
def main(query,date,db1,path,idfield):
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
       
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)
    message = ListMessagesMatchingQuery(service,'me','')
    todaydate =datetime.date.today().strftime('%Y-%m-%d')
    querydate = date
    counter = 0
    #open('Email.txt','w').close()
    ecxel_sheet_data=[]
    fromemaillist=[]
    compny_name_from_query = query.strip().lower()
    
    for m in message:
        print('********************',counter,'****************************')
        counter =counter+1
        messageid = m['id']
        print(messageid)
        messagejson =  GetMessage(service,'me',messageid)
        dateinternal = int(messagejson['internalDate']) / 1000.0
        messagedate = datetime.datetime.fromtimestamp(dateinternal).strftime('%Y-%m-%d')
       
       ########################### Mathching the Date for fatching the email ##################
        if querydate <= messagedate <= todaydate :
            headers = messagejson['payload']['headers']
            subject= [i['value'] for i in headers if i["name"]=="Subject"]
            mailsubject = subject[0]
            
            compny_name_from_mail = mailsubject.split()[-1].lower()
            
            
            ############################## Attchment download ####################################
            if compny_name_from_mail == compny_name_from_query:  

                #########################Fatching body part#########################################
                for part in messagejson['payload']['parts']:
                    if part['mimeType'] == 'text/plain':
                        data1 = part['body']['data']
                        msg_str = base64.urlsafe_b64decode(data1.encode('utf-8','ignore'))
                        mime_msg = email.message_from_string(msg_str.decode('utf-8','ignore'))
                        print(mime_msg)
                        msg = str(mime_msg)
                        print(msg)
                        ecxel_sheet_data.append(msg)
                ##################################Appending (From) Email in the List##################################################         
                data = dict()
                for h in headers:
                    data[h['name']]=h['value']
                From = "From :" +data['From'] + '\n'
                fromemaillist.append(data['From'])

                if not os.path.exists(path+compny_name_from_mail):
                    os.makedirs(path+compny_name_from_mail)
                ################################## Download the attachment ########################
                GetAttachments(service,'me',messageid,path+compny_name_from_query+'/')
                ###################################################################################
        else:
            print(ecxel_sheet_data)    
            break  

    errorcount = ExportToexcel(ecxel_sheet_data,compny_name_from_query,path)
    print("-------------------------",errorcount)
    
    if errorcount !=None:
        ExportEmail(fromemaillist,compny_name_from_query,path)
        FinalSheet(compny_name_from_query+'.xlsx',compny_name_from_query,db1,path,idfield)
    
        if errorcount>0:
            checkemail(compny_name_from_query,db1,path) 
    else:
        errorcount = 1000
    return errorcount
    


    
