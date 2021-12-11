import pandas as pd
import datetime
import smtplib
import os

from pandas.tseries.offsets import BDay

current_path = os.getcwd()
print(current_path)
os.chdir(current_path)

GMAIL_ID = input("Enter your Gmail ID: ")
GMAIL_PSWD = input("Enter you Gmail Password: ")

def sendEmail(to, sub, msg):
    print(f"Email to {to} sent: \nSubject: {sub}, \nMessage: {msg}")
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PSWD)
    s.sendmail(GMAIL_ID, to, f"Subject: {sub} \n\n {msg}")
    s.quit()

if __name__ == "__main__":
    df = pd.read_excel("friends_data.xlsx")
    today = datetime.datetime.now().strftime("%d-%m")
    yearNow = datetime.datetime.now().strftime("%Y")

    writeInd = []
    for index, item in df.iterrows():
        bday = item['Birthday']
        bday = datetime.datetime.strptime(bday, "%d-%m-%Y")
        bday = bday.strftime("%d-%m")
        if (today == bday) and yearNow not in str(item['LastWishedYear']):
            sendEmail(item['Email'], "Happy Birthday", item['Dialogue'])
            writeInd.append(index)

    if writeInd != None:
        for i in writeInd:
            oldYear = df.loc[i, 'LastWishedYear']
            df.loc[i, 'LastWishedYear'] = str(oldYear) +", " + str(yearNow)

    df.to_excel('friends_data.xlsx', index = False)
    