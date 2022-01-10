import pandas as pd

data=pd.read_excel('With_Emails.xlsx')
#print(data['City'],type(data['City']),type(data))
if 'Email' in data.columns:
    emails=list(data['Email'])
    #print(emails)
    c=[]
    for i in emails:
        #print(i)
        if pd.isnull(i)==False:
            #print(i)
            c.append(i)
    emails=c
    print(emails)
else:
    print("Not Exist")