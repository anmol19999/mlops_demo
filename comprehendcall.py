import win32com.client
import boto3
import json
import pandas as pd



#Fetching Email
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) 
messages = inbox.Items
#feedback from yatin
messages = messages.Restrict("[Subject] = 'Client Requirements for xyz project'")
f= open("temp_yatin.txt","w+")
for message in messages:
    body_content = message.body
    body_subject = message.subject
    body_sendername = message.SenderName
    body_content=body_content[body_content.find('Anmol'):]
    body_content=body_content[:body_content.find('Sincerely')]
    #print(body_content)
    f.write(body_content)
f.close()





#Amazon Comprehend
with open('temp_yatin.txt', 'r') as file:
    data = file.read().replace('\n', '')
comprehend = boto3.client('comprehend')

response = comprehend.detect_entities(
    Text = data,
    LanguageCode='en',
)
lis = []
entlist = []
lis = response['Entities']
for i in lis:
    for j in i:
        if j == "Text":
            entlist.append(i[j])

#print(entlist)
#entlist= ['Anmol', 'AWS', 'Azure']
df = pd.DataFrame(pd.read_excel("Devops Skill Set.xlsx"))
df.drop([1], axis=0, inplace=True)
df = df.fillna(0)
new_header = df.iloc[0]
df = df[1:]
df.columns = new_header
ls = df.columns
ls = ls[2:]

for i in ls:
    df[i] = pd.to_numeric(df[i])
temp = []
for i in df.columns:
    if i in entlist:
        temp.append(i)
df['total']=0
for i in temp:
    df['total'] = df['total'] + df[i]
    #print(i)
top=df['total'].max()
top_data=df.loc[df['total']==top]
print(top_data['Name'])






#JSON output
json_object = json.dumps(response, indent = 4) 
#print(json_object)
f= open("outputo.json","w+")
f.write(json_object)
f.close()





