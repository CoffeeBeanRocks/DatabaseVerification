# @Author: Ethan Meyers
# @Email: ewm1230@gmail.com
# @Phone: 847-212-2264
# @Created Date: 07/13/2022

import os
import sys
import numpy as np
import pandas as pd
import pyodbc as pyodbc
import win32com.client
import zipfile
from pathlib import Path
import traceback
import io
from contextlib import redirect_stderr


class Data:
    skippedLines = ''
    mailTo = 'emeyers@whimsytrucking.com'
    inboxEmail = 'emeyers@whimsytrucking.com'
    tableName = 'Pick Up 2022 Cont'


# @Param reason (String): Reason that will be inserted into email regarding why the program failed.
# @Description: Sends email to specified recipient about reason for program failure.
def sendFailureEmail(reason: Exception, trace: str):
    olApp = win32com.client.Dispatch("Outlook.Application")
    olNS = olApp.GetNamespace("MAPI")

    mailItem = olApp.CreateItem(0)
    mailItem.Subject = 'Automatic Task Failure'
    mailItem.BodyFormat = 1
    mailItem.Body = "This is an automated email informing you that the task 'Default_TEST " \
                    "Upload' has failed for the following reason:\n\"{}\"\n\n{}\n\n".format(reason, trace)
    mailItem.To = Data.mailTo
    mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item(Data.inboxEmail)))

    mailItem.Save()
    mailItem.Send()

    sys.exit(1)


# @Param message: Tuple that contains any messages that should be displayed in the success email
# @Param df: Dataframe containing all the data that was entered into access
# @Description: Sends email to specified recipient about reason for program success.
def sendSuccessEmail(lines: str, df: pd.DataFrame):
    olApp = win32com.client.Dispatch("Outlook.Application")
    olNS = olApp.GetNamespace("MAPI")

    mailItem = olApp.CreateItem(0)
    mailItem.BodyFormat = 1
    if len(lines) > 0:
        mailItem.Subject = 'Automatic Task Success (Warning: Skipped Lines)'
        mailItem.Body = "This is an automated email informing you that the task 'Default_TEST " \
                        "Upload' has been completed successfully but with the following " \
                        "skipped lines:\n\n{}\n\n".format(lines)
    else:
        mailItem.Subject = 'Automatic Task Success'
        mailItem.Body = "This is an automated email informing you that the task 'Default_TEST " \
                        "Upload' has been completed successfully!\n\n"

    mailItem.To = Data.mailTo
    mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item(Data.inboxEmail)))

    dir_path = '%s\\DefaultTestAuto\\' % os.environ['APPDATA']
    dir_path = Path(dir_path)
    if not os.path.exists(dir_path / 'LinesAdded'):
        os.makedirs(dir_path / 'LinesAdded')
    dir_path = dir_path / 'LinesAdded'
    w2Path = dir_path / 'LinesAdded.xlsx'

    writer = pd.ExcelWriter(w2Path, engine='openpyxl')
    df.to_excel(writer, sheet_name='Output', index=False)
    writer.save()

    mailItem.Attachments.Add(Source=str(w2Path))

    mailItem.Save()
    mailItem.Send()

    sys.exit(0)


# @Return DataFrame: Dataframe containing the information from Default_TEST CSV
# @Description: Retrieves Default_TEST file from Outlook email
def getFileFromEmail() -> pd.DataFrame:
    # Find desired email in inbox
    dir_path = '%s\\DefaultTestAuto\\' % os.environ['APPDATA']
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)
    dir_path = Path(dir_path)
    if not os.path.exists(dir_path / 'DownloadedEmailAttachments'):
        os.makedirs(dir_path / 'DownloadedEmailAttachments')

    dir_path = dir_path / 'DownloadedEmailAttachments'

    # Removing attachments from previous run
    for i in os.listdir(dir_path):
        os.remove(dir_path / i)

    olApp = win32com.client.Dispatch("Outlook.Application")
    olNS = olApp.GetNamespace("MAPI")

    inbox = olNS.GetDefaultFolder(6)
    item = None
    for messages in inbox.Items:
        if 'DEFAULT_TEST AUTO' in messages.Subject:
            item = messages
            break

    if item is None:
        raise Exception('Default_TEST email not found in inbox')
    else:
        # Extracts CSV from email, places it in a local directory, and converts info from CSV to dataframe
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

        # Find .csv in email and format new pandas dataframe
        if '.csv' in str(csvPath):
            try:
                df = pd.read_csv(csvPath, header=0, encoding='unicode_escape')
            except pd.errors.ParserError:
                # Saves any skipped lines to str myOut
                with io.StringIO() as buf, redirect_stderr(buf):
                    df = pd.read_csv(csvPath, header=0, encoding='unicode_escape', error_bad_lines=False)
                    myOut = buf.getvalue()
                    try:
                        if len(myOut) > 0:
                            myOut = myOut[myOut.index("b'") + 1:]
                            myOut = myOut.replace('\\n', '] [')
                            myOut = myOut.replace("'", '')
                            myOut = '[' + myOut + ']'
                            myOut = myOut[0:len(myOut) - 3]  # TODO: Testing needed
                    except:
                        myOut = buf.getvalue()
                    Data.skippedLines = myOut
            except Exception as e:
                raise Exception('Encountered a fatal error while reading file\n{}'.format(e))

            df.drop('Cost', axis=1, inplace=True)
            df.drop('Inv', axis=1, inplace=True)
            df.drop('Site', axis=1, inplace=True)
            df.drop('Status', axis=1, inplace=True)
            df.drop('OWT', axis=1, inplace=True)
            df.drop('Live', axis=1, inplace=True)
            df.drop('Revenue', axis=1, inplace=True)

            df.replace({pd.NaT: None}, inplace=True)
            df = df.fillna('')
            end = []
            df['Unnamed: 19'] = np.NAN
            for i in range(0, len(df.index)):
                string = df['Order #'][i]
                zero = string.index('0')
                slash = string.index('/')
                df['Unnamed: 19'][i] = int(string[zero:slash])
                end.append(int(string[slash + 1:]))
            df['End'] = end
            df = df.sort_values(by=['Unnamed: 19', 'End'], ignore_index=True)
            # TODO: Delete mail item when done (MailItem.Delete)
            # TODO: Delete .csv item when done (os.remove(csvPath))
            return df
        else:
            # TODO: Delete mail item when done (MailItem.Delete)
            raise Exception('Incompatible attachment {}'.format(csvPath))


# @Param filePath: represents the filePath to the csv holding the pertinent data
# @Description: Returns a pandas dataframe consisting of some old and some new records from a .csv file.
def getFileFromLocal(filePath: Path) -> pd.DataFrame:
    df = pd.read_csv(filePath, header=0, encoding='unicode_escape')
    df.replace({pd.NaT: None}, inplace=True)
    df = df.fillna('')
    end = []
    df['Unnamed: 19'] = np.NAN
    for i in range(0, len(df.index)):
        string = df['Order #'][i]
        zero = string.index('0')
        slash = string.index('/')
        df['Unnamed: 19'][i] = int(string[zero:slash])
        end.append(int(string[slash + 1:]))
    df['End'] = end
    df = df.sort_values(by=['Unnamed: 19', 'End'], ignore_index=True)
    return df


# @Param value: The string that will be adjusted
# @Return str: The newly formatted string
# @Description: Takes an input string, removes all instances of '=' and '"' then returns the result
def normalizeStr(value: str) -> str:
    temp = value
    if '=' in temp:
        temp = temp.replace('=', '')
    if '"' in temp:
        temp = temp.replace('"', '')
    return temp


# @Description: Inserts new records into designated microsoft access database
def updateAccess():
    # Setup
    filePath = r"C:\Users\emeyers\Desktop\EthanAccess.accdb"  # TODO: Change to live version
    driver = pyodbc.dataSources()
    driver = driver['MS Access Database']
    connection = pyodbc.connect(driver=driver, dbq=filePath)
    cursor = connection.cursor()

    df = getFileFromEmail()
    # df = getFileFromLocal(r"C:\Users\emeyers\Desktop\2022071826852.csv")
    lastUpdate = 0  # TODO: Remove completion percentage for final build
    maxItem = 0

    # Finds the last record in the dataframe that's also in Access
    for i in range(0, len(df.index)):  # TODO: Change to search from the bottom up
        if int((i / len(df.index)) * 100) != lastUpdate:
            lastUpdate = int((i / len(df.index)) * 100)
            print("{}%".format(lastUpdate))
        find = "SELECT * FROM [{}] WHERE [{}] = '{}'".format(Data.tableName, 'Order #', df.iloc[i]['Order #'])
        cursor.execute(find)
        if cursor.fetchone() is not None:
            maxItem = max(maxItem, int(df.iloc[i]['Unnamed: 19']))

    df2 = df[df['Unnamed: 19'] > maxItem]
    df2.drop('Unnamed: 19', axis=1, inplace=True)
    df2.drop('End', axis=1, inplace=True)

    # Inserts the new records into Access using MySQL syntax (';' not required)
    for i in range(0, len(df2.index)):
        customerRef = normalizeStr(df2.iloc[i]['Customer Ref'])
        masterBOL = normalizeStr(df2.iloc[i]['Master BOL/Booking Ref'])
        containerNum = normalizeStr(df2.iloc[i]['Container #'])

        cursor.execute("INSERT INTO [{}] ([User], [EDI], [Order Date], [Order #], [Container #], "
                       "[Master BOL/Booking Ref], [Customer], [Customer Ref], [Pick Up], [Delivery], "
                       "[DL City]) VALUES (?,?,?,?,?,?,?,?,?,?,?)".format(Data.tableName), df2.iloc[i]['User'],
                       df2.iloc[i]['EDI'], df2.iloc[i]['Order Date'], df2.iloc[i]['Order #'], containerNum, masterBOL,
                       df2.iloc[i]['Customer'], customerRef, df2.iloc[i]['Pick Up'], df2.iloc[i]['Delivery'],
                       df2.iloc[i]['DL City'])

    print('Commit in progress...')
    connection.commit()
    sendSuccessEmail(Data.skippedLines, df2)
    print('Finished')


if __name__ == '__main__':
    try:
        if len(sys.argv) > 1:
            Data.tableName = sys.argv[1]
            Data.mailTo = sys.argv[2]
        updateAccess()
    except Exception as e:
        print(traceback.format_exc())
        # sendFailureEmail(e, traceback.format_exc())
