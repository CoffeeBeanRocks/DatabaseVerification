# @Author: Ethan Meyers
# @Email: ewm1230@gmail.com
# @Phone: 847-212-2264
# @Created Date: 07/13/2022

import os
import sys
import pandas as pd
import pyodbc as pyodbc
import win32com.client
import zipfile
from pathlib import Path

mailTo = 'emeyers@whimsytrucking.com'
inboxEmail = 'emeyers@whimsytrucking.com'


# @Param reason (String): Reason that will be inserted into email regarding why the program failed.
# @Description: Sends email to specified recipient about reason for program failure.
def sendFailureEmail(reason: str):
    olApp = win32com.client.Dispatch("Outlook.Application")
    olNS = olApp.GetNamespace("MAPI")

    mailItem = olApp.CreateItem(0)
    mailItem.Subject = 'Automatic Task Failure'
    mailItem.BodyFormat = 1
    mailItem.Body = "This is an automated email informing you that the task 'Default_TEST " \
                    "Upload' has failed for the following reason:\n\"{}\"\n\n".format(reason)
    mailItem.To = mailTo
    mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item(inboxEmail)))

    mailItem.Save()
    mailItem.Send()

    sys.exit(1)

# @Return DataFrame: Dataframe containing the information from Default_TEST CSV
# @Description: Retrieves Default_TEST file from Outlook email
def getFileFromEmail():
    # Find desired email in inbox
    dir_path = '%s\\DefaultTestAuto\\' % os.environ['APPDATA']
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)

    olApp = win32com.client.Dispatch("Outlook.Application")
    olNS = olApp.GetNamespace("MAPI")

    inbox = olNS.GetDefaultFolder(6)
    item = None
    for messages in inbox.Items:
        if 'AUTO DEFAULT TEST' in messages.Subject:
            item = messages
            break

    if item is None:
        raise Exception('Default_TEST email not found in inbox')
    else:
        # Extracts CSV from email, places it in a local directory, and converts info from CSV to dataframe
        dir_path = Path(dir_path)
        csvPath = ''
        for attachment in item.Attachments:
            attachment.SaveAsFile(dir_path / str(attachment))

        for i in os.listdir(dir_path):
            if '.zip' in i:
                with zipfile.ZipFile(dir_path / i, 'r') as zip_ref:
                    zip_ref.extractall(dir_path)
                os.remove(dir_path / i)

        for i in os.listdir(dir_path):
            csvPath = dir_path / i

        if '.csv' in str(csvPath):
            df = pd.read_csv(csvPath, header=0, encoding='unicode_escape')
            df.replace({pd.NaT: None}, inplace=True)
            df = df.fillna('')
            # TODO: Delete mail item when done (MailItem.Delete)
            return df
        else:
            # TODO: Delete mail item when done (MailItem.Delete)
            raise Exception('Unknown attachment found')


# @Description: Returns a pandas dataframe consisting of some old and some new records from a .csv file.
def getFileFromLocal():
    FilePath = r'C:\Users\emeyers\Desktop\default_test_2.csv'
    df = pd.read_csv(FilePath, header=0, encoding='unicode_escape')
    df.replace({pd.NaT: None}, inplace=True)
    df = df.fillna('')
    end = []
    for i in range(0, len(df.index)):
        string = df['Order #'][i]
        zero = string.index('0')
        slash = string.index('/')
        df['Unnamed: 19'][i] = int(string[zero:slash])
        end.append(int(string[slash + 1:]))
    df['End'] = end
    df = df.sort_values(by=['Unnamed: 19', 'End'], ignore_index=True)
    return df


# @Description: Inserts new records into designated microsoft access database
def updateAccess():
    # Setup
    filePath = r"C:\Users\emeyers\Desktop\EthanAccess.accdb"
    driver = pyodbc.dataSources()
    driver = driver['MS Access Database']
    connection = pyodbc.connect(driver=driver, dbq=filePath)
    cursor = connection.cursor()
    tableName = 'Pick Up 2022 Cont'

    df = getFileFromEmail()
    lastUpdate = 0
    maxItem = 0

    # Finds the last record in the dataframe that's also in Access
    for i in range(0, len(df.index)):
        if int((i / len(df.index)) * 100) != lastUpdate:
            lastUpdate = int((i / len(df.index)) * 100)
            print("{}%".format(lastUpdate))
        find = "SELECT * FROM [{}] WHERE [{}] = '{}'".format(tableName, 'Order #', df.iloc[i]['Order #'])
        cursor.execute(find)
        if cursor.fetchone() is not None:
            maxItem = max(maxItem, int(df.iloc[i]['Unnamed: 19']))

    df2 = df[df['Unnamed: 19'] > maxItem]
    df2.drop('Unnamed: 19', axis=1, inplace=True)
    df2.drop('End', axis=1, inplace=True)

    # Inserts the new records into Access using MySQL syntax (';' not required)
    for i in range(0, len(df2.index)):
        cursor.execute("INSERT INTO [Pick Up 2022 Cont] ([User], [EDI], [Order Date], [Order #], [Container #], "
                       "[Master BOL/Booking Ref], [Customer], [Customer Ref], [Pick Up], [Delivery], "
                       "[DL City]) VALUES (?,?,?,?,?,?,?,?,?,?,?)", df2.iloc[i]['User'], df2.iloc[i]['EDI'],
                       df2.iloc[i]['Order Date'], df2.iloc[i]['Order #'], df2.iloc[i]['Container #'],
                       str(df2.iloc[i]['Master BOL/Booking Ref']), df2.iloc[i]['Customer'],
                       str(df2.iloc[i]['Customer Ref']),
                       df2.iloc[i]['Pick Up'], df2.iloc[i]['Delivery'], df2.iloc[i]['DL City'])

    print('Commit in progress...')
    connection.commit()
    print('Finished')


if __name__ == '__main__':
    try:
        updateAccess()
    except Exception as e:
        print(e)
        sendFailureEmail(e)
