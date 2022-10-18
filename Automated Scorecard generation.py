# import the required libraries
import gspread
import pandas as pd
import numpy as np
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
from fpdf import FPDF
from pretty_html_table import build_table
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from smtplib import SMTP
import smtplib
from email.message import EmailMessage
import sys
import smtplib

##Defining the scope of api and sheet

scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]


creds = ServiceAccountCredentials.from_json_keyfile_name("#insert_json_filename#.json", scope)

client = gspread.authorize(creds)

##Opening the sheet

sheet = client.open("#googlesheet_name#").sheet1

##Calling in all records of the sheet

data = sheet.get_all_records()
print(data)

df = pd.DataFrame(data)


ind=len(df.index)-1
print(ind)
i=0
df1 =pd.DataFrame(columns=['Component','Level Achieved','Score Achieved','Minimum Level Required'],index=[0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23])
flag_df2= False
flag_df3= False

    #for C1
df1.loc[df1.index[i], 'Component'] = 'C1'
if(df.loc[df.index[ind], 'C1.1'] == 'Less than 50%'):
        df1.loc[df1.index[i],'Level Achieved'] = 'Level 0'
        df1.loc[df1.index[i],'Score Achieved'] = 0
        df1.loc[df1.index[i],'Minimum Level Required'] = 'Level 1'

elif(df.loc[df.index[ind], 'C1.1']  == '50% to 70%'):
        df1.loc[df1.index[i], 'Level Achieved'] = 'Level 1'
        df1.loc[df1.index[i], 'Score Achieved'] = 150
        df1.loc[df1.index[i], 'Minimum Level Required'] = 'Level 1'

elif((((df.loc[df.index[ind], 'C1.1']  == '70% to 90%') | (df.loc[df.index[ind], 'C1.1']  == '90% to 100%')) | (df.loc[df.index[ind], 'C1.1'] == '100%')) & (df.loc[df.index[ind], 'C1.2'] == 'No')):
            df1.loc[df1.index[i], 'Level Achieved'] = 'Level 1'
            df1.loc[df1.index[i], 'Score Achieved'] = 150
            df1.loc[df1.index[i], 'Minimum Level Required'] = 'Level 1'

elif((df.loc[df.index[ind], 'C1.1'] == '70% to 90%') & (df.loc[df.index[ind], 'C1.2'] == 'Yes')):
        df1.loc[df1.index[i], 'Level Achieved'] = 'Level 2'
        df1.loc[df1.index[i], 'Score Achieved'] = 200
        df1.loc[df1.index[i], 'Minimum Level Required'] = 'Level 1'

elif((df.loc[df.index[ind], 'C1.1'] == '90% to 100%') & (df.loc[df.index[ind], 'C1.2'] == 'Yes')):
        df1.loc[df1.index[i], 'Level Achieved'] = 'Level 3'
        df1.loc[df1.index[i], 'Score Achieved'] = 250
        df1.loc[df1.index[i], 'Minimum Level Required'] = 'Level 1'

elif((df.loc[df.index[ind], 'C1.1'] == '100%') & (df.loc[df.index[ind], 'C1.2'] == 'Yes')):
        df1.loc[df1.index[i], 'Level Achieved'] = 'Level 4'
        df1.loc[df1.index[i], 'Score Achieved'] = 300
        df1.loc[df1.index[i], 'Minimum Level Required'] = 'Level 1'


i=i+1

    #C2
df1.loc[df1.index[i], 'Component'] = 'C2'
if (df.loc[df.index[ind], 'C2.1'] == 'Less than 40%'):
        df1.loc[df1.index[i], 'Level Achieved'] = 'Level 0'
        df1.loc[df1.index[i], 'Score Achieved'] = 0
        df1.loc[df1.index[i], 'Minimum Level Required'] = 'Level 1'
elif(df.loc[df.index[ind], 'C2.1'] == '40% to 60%'):
        df1.loc[df1.index[i], 'Level Achieved'] = 'Level 1'
        df1.loc[df1.index[i], 'Score Achieved'] = 350
        df1.loc[df1.index[i], 'Minimum Level Required'] = 'Level 1'
elif(df.loc[df.index[ind], 'C2.1'] == '60% to 80%'):
        df1.loc[df1.index[i], 'Level Achieved'] = 'Level 2'
        df1.loc[df1.index[i], 'Score Achieved'] = 450
        df1.loc[df1.index[i], 'Minimum Level Required'] = 'Level 1'
elif(df.loc[df.index[ind], 'C2.1'] == '80% to 90%'):
        df1.loc[df1.index[i], 'Level Achieved'] = 'Level 3'
        df1.loc[df1.index[i], 'Score Achieved'] = 575
        df1.loc[df1.index[i], 'Minimum Level Required'] = 'Level 1'
elif(df.loc[df.index[ind], 'C2.1'] == 'Greater than or equal to 90%'):
        df1.loc[df1.index[i], 'Level Achieved'] = 'Level 4'
        df1.loc[df1.index[i], 'Score Achieved'] = 700
        df1.loc[df1.index[i], 'Minimum Level Required'] = 'Level 1'
else:
        df1.loc[df1.index[i], 'Level Achieved'] = 'Level 0'
        df1.loc[df1.index[i], 'Score Achieved'] = 0
        df1.loc[df1.index[i], 'Minimum Level Required'] = 'Level 1'

i=i+1
    #C3
df1.loc[df1.index[i], 'Component'] = 'C3'
str1=df.loc[df.index[ind], 'C18.1']
str_arr=str1.split(',')
length=len(str_arr)
if((length == 3)|(length == 4)):
        df1.loc[df1.index[i], 'Level Achieved'] = 'Level 1'
        df1.loc[df1.index[i], 'Score Achieved'] = 100
        df1.loc[df1.index[i], 'Minimum Level Required'] = 'Level 1'
elif((length==5)|(length==6)):
        df1.loc[df1.index[i], 'Level Achieved'] = 'Level 2'
        df1.loc[df1.index[i], 'Score Achieved'] = 150
        df1.loc[df1.index[i], 'Minimum Level Required'] = 'Level 1'
elif((length == 7)|(length == 8)):
        df1.loc[df1.index[i], 'Level Achieved'] = 'Level 3'
        df1.loc[df1.index[i], 'Score Achieved'] = 225
        df1.loc[df1.index[i], 'Minimum Level Required'] = 'Level 1'
elif(length == 9):
        df1.loc[df1.index[i], 'Level Achieved'] = 'Level 4'
        df1.loc[df1.index[i], 'Score Achieved'] = 300
        df1.loc[df1.index[i], 'Minimum Level Required'] = 'Level 1'
else:
        df1.loc[df1.index[i], 'Level Achieved'] = 'Level 0'
        df1.loc[df1.index[i], 'Score Achieved'] = 0
        df1.loc[df1.index[i], 'Minimum Level Required'] = 'Level 1'

df1 = df1.replace(np.nan, '-')

if(i==23): ##Here i=23 denotes the index number of the last question in the questionnaire.
        sum = df1['Score Achieved'].sum()
        obj='Your Total Score Achieved is ' + format(sum)
        print(obj)
        str=''
        if((sum < 2400) | ('Level 0' in df1.iloc[0:16,1].values.tolist())):
            str='You have scored "0" Star.'
            print(str)
        elif(((sum >= 2400) & (sum < 3600)) & ('Level 0' not in df1.iloc[0:16, 1].values.tolist())):
            str='You have scored "1" Star.'
            print(str)
        elif(((sum >= 3600) & (sum <6300)) & ('Level 0' not in df1.iloc[0:16, 1].values.tolist())):
            str = 'You have scored "3" Star.'
            print(str)
        elif((sum >= 6300) & ('Level 0' in df1.iloc[16:, 1].values.tolist())):
            df1.loc[df1.index[i], 'Star Rating obtained'] = 3
            str = 'You have scored "3" Star.'
            print(str)
        elif(((sum >=6300) & (sum <7500)) & ('Level 0' not in df1.iloc[0:, 1].values.tolist())):
            str = 'You have scored "5" Star.'
            print(str)
        elif((sum == 7500) & ('Level 0' not in df1.iloc[0:, 1].values.tolist())):
            str = 'You have scored "7" Star.'
            print(str)
        print(df1)

        #Way forward matrix
        msg=''
        sen=''
        if(str=='You have scored "0" Star'):
            msg='You can aim to achieve 1-star & 3-star, the requirements of which are given below:'
            print(msg)
            data1=[['C1','Level 1',150,'"At-least 50% to 70%'],['C2','Level 1',350,'At-least 40% to 60%'],['C3','Level 1',150,'At-least 40% to 60%']]
            df2 = pd.DataFrame(data1,columns=['Component', 'Compulsory Level Required', 'Marks To be obtained', 'Requirements'])
            sen = '3-Star:'
            print(sen)
            data2 = [['C1', 'Level 2', 200,'At-least 70% to 90%'],['C2', 'Level 2', 450, 'At-least 60% to 80%'],['C3', 'Level 2', 200, 'At-least 60% to 80%']]
            df3 = pd.DataFrame(data2, columns=['Component', 'Compulsory Level Required', 'Marks To be obtained',
                                               'Requirements'])

        elif(str=='You have scored "1" Star'):
            msg='You can aim to achieve 3-star & 5-star, the requirements of which are given below:'
            print(msg)
            data1=[['C1','Level 2',200,'At least 70% to 90%'],['C2','Level 2',450,'At-least 60% to 80%'],['C3','Level 2',200,'At-least 60% to 80%']]
            df2 = pd.DataFrame(data1, columns=['Component', 'Compulsory Level Required', 'Marks To be obtained', 'Requirements'])
            sen = '5-Star:'
            print(sen)
            data2 = [['C1', 'Level 3', 250,'At-least 70% to 90%'],['C2', 'Level 3', 575, 'At-least 60% to 80% source segregation'],['C3', 'Level 3', 250,'At-least 60% to 80%']]
            df3 = pd.DataFrame(data2, columns=['Component', 'Compulsory Level Required', 'Marks To be obtained',
                                               'Requirements'])

        elif(str=='You have scored "3" star'):
            msg='You can aim to achieve 5-star & 7-star:'
            print(msg)
            data1 = [['C1', 'Level 3', 250,'At-least 70% to 90% '],['C2', 'Level 3', 575, 'At-least 60% to 80%'], ['C3', 'Level 3', 250,'At-least 60% to 80%']]
            df2 = pd.DataFrame(data1, columns=['Component', 'Compulsory Level Required', 'Marks To be obtained',
                                              'Requirements'])
            sen = '7-Star:'
            print(sen)
            data2 = [['C1', 'Level 4', 300, '100%'],['C2', 'Level 4', 700, '>=90%'], ['C3', 'Level 4', 300, '> 90%']]
            df3 = pd.DataFrame(data2, columns=['Component','Compulsory Level Required','Marks To be obtained',
                                              'Requirements'])

        elif(str=='You have scored "5" Star'):
            msg='You can aim to achieve 7-star:'
            print(msg)
            flag_df3 = True
            data1 = [['C1', 'Level 4', 300,
                     '100%'],
                    ['C2', 'Level 4', 700, '>=90% '], ['C3', 'Level 4', 300, '> 90%']]
            df2 = pd.DataFrame(data1, columns=['Component', 'Compulsory Level Required', 'Marks To be obtained',
                                               'Requirements'])

        elif(str=='You have scored "7" Star:'):
            msg='Congratulations! You have achieved highest star rating. You should retain your rating by continuous improvement.'
            print(msg)
            flag_df2 = True


        import smtplib, ssl

## For the autogeneration and mailing of scorecard to respondents:
        port = 465  # For SSL
        smtp_server = "smtp.gmail.com"
        sender_email = "sender_address@gmail.com"  # Enter your address
        receiver_email = df.loc[df.index[ind], 'Email']  # This Enters receiver address which was there in the questionnaire and hence the google sheet
        password = "16 digit google code" #For the 2-step verification of our account, we can generate a 16 digit code which can then be used to login into our email id
        email = EmailMessage()
        email['Subject'] = f"Subject: Scorecard "
        email['From'] = sender_email
        email['To'] = receiver_email
        if (flag_df2):
                email.set_content(f"""\
                          <html><head></head><body>
                          <p><b>Dear respondent,</b></p>
                          <p><b>Please have a look at your scorecard:</b></p>
                          <p>{obj}</p>
                          <p>{str}</p>
                          {df1.to_html()}
                          <p>{msg}</p>
                          </body></html>""", subtype='html')
        elif (flag_df3):
                email.set_content(f"""\
                          <html><head></head><body>
                          <p><b>Dear respondent,</b></p>
                          <p><b>Please have a look at your scorecard:</b></p>
                          <p>{obj}</p>
                          <p>{str}</p>
                          {df1.to_html()}
                          <p>{msg}</p>
                          {df2.to_html()}
                          </body></html>""", subtype='html')
        else:
                email.set_content(f"""\
                          <html><head></head><body>
                          <p><b>Dear respondent,</b></p>
                          <p><b>Please have a look at your scorecard:</b></p>
                          <p>{obj}</p>
                          <p>{str}</p>
                          {df1.to_html()}
                          <p>{msg}</p>
                          {df2.to_html()}
                          <p>{sen}</p>
                          {df3.to_html()}
                          </body></html>""", subtype='html')
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
                server.login(sender_email, password)
                server.send_message(email)
















































































