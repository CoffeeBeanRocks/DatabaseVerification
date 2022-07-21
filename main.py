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

# @Var skippedLines: holds a string of any warnings generated by pd.read_csv()
# @Var mailTo: The email address where success and failure emails will be sent
# @Var inboxEmail: The email address where success and failure emails will be generated and sent from
# @Var tableName: The name of the table within the access file
# @Var csvPath: Path to the downloaded Default_TEST csv locally
# @Description: Class for necessary global variables
class Data:
    skippedLines = ''
    mailTo = 'emeyers@whimsytrucking.com'
    inboxEmail = 'emeyers@whimsytrucking.com'
    tableName = 'Pick Up 2022 Cont'
    csvPath = ''


# @Param reason (String): Reason that will be inserted into email regarding why the program failed.
# @Description: Sends email to specified recipient about reason for program failure.
def sendFailureEmail(reason: Exception, trace: str):
    # Connecting to outlook
    olApp = win32com.client.Dispatch("Outlook.Application")
    olNS = olApp.GetNamespace("MAPI")

    # Creating failure email
    mailItem = olApp.CreateItem(0)
    mailItem.Subject = 'Automatic Task Failure'
    mailItem.BodyFormat = 1
    mailItem.Body = "This is an automated email informing you that the task 'Default_TEST " \
                    "Upload' has failed for the following reason:\n\"{}\"\n\n{}\n\n".format(reason, trace)
    mailItem.To = Data.mailTo
    mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item(Data.inboxEmail)))

    # Sending
    mailItem.Save()
    mailItem.Send()

    # Error occurred, end program
    sys.exit(1)


# @Param message: Tuple that contains any messages that should be displayed in the success email
# @Param df: Dataframe containing all the data that was entered into access
# @Description: Sends email to specified recipient about reason for program success.
def sendSuccessEmail(lines: str, df: pd.DataFrame):
    # Connecting to outlook
    olApp = win32com.client.Dispatch("Outlook.Application")
    olNS = olApp.GetNamespace("MAPI")

    # Creating success email
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

    # Creating Lines Added attachment
    dir_path = '%s\\DefaultTestAuto\\' % os.environ['APPDATA']
    dir_path = Path(dir_path)
    if not os.path.exists(dir_path / 'LinesAdded'):
        os.makedirs(dir_path / 'LinesAdded')
    dir_path = dir_path / 'LinesAdded'
    w2Path = dir_path / 'LinesAdded.xlsx'
    writer = pd.ExcelWriter(w2Path, engine='openpyxl')
    df.to_excel(writer, sheet_name='Output', index=False)
    writer.save()

    # Adding attachments to email
    mailItem.Attachments.Add(Source=str(w2Path))
    mailItem.Attachments.Add(Source=str(Data.csvPath))

    # Sending
    mailItem.Save()
    mailItem.Send()

    # End program
    sys.exit(0)


# @Return DataFrame: Dataframe containing the information from Default_TEST CSV
# @Description: Retrieves Default_TEST file from Outlook email
def getFileFromEmail() -> pd.DataFrame:
    # Creating proper directory structure TODO: Consolidate all the references to dir_path in Data class
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

    # Connecting to outlook
    olApp = win32com.client.Dispatch("Outlook.Application")
    olNS = olApp.GetNamespace("MAPI")

    # Finding Default_TEST email in inbox
    inbox = olNS.GetDefaultFolder(6)
    item = None
    for messages in inbox.Items:
        if 'DEFAULT_TEST AUTO' in messages.Subject:
            item = messages
            break

    # Validates if required document has been received
    if item is None:
        raise Exception('Default_TEST email not found in inbox')
    else:
        # Downloads all attachments contained in email
        for attachment in item.Attachments:
            attachment.SaveAsFile(dir_path / str(attachment))

        # If necessary file is in a .zip file, it extracts the file from the .zip and removes the .zip file
        for i in os.listdir(dir_path):
            if '.zip' in i:
                with zipfile.ZipFile(dir_path / i, 'r') as zip_ref:
                    zip_ref.extractall(dir_path)
                os.remove(dir_path / i)

        # Finds .csv in list of attachments
        for i in os.listdir(dir_path):
            if '.csv' in i:
                Data.csvPath = dir_path / i

        # Validates .csv was found
        if '.csv' in str(Data.csvPath):
            # Converting data from .csv to pandas dataframe
            try:
                df = pd.read_csv(Data.csvPath, header=0, encoding='unicode_escape')
            except pd.errors.ParserError:
                # Records location of any skipped lines in .csv
                with io.StringIO() as buf, redirect_stderr(buf):
                    df = pd.read_csv(Data.csvPath, header=0, encoding='unicode_escape', error_bad_lines=False)
                    myOut = buf.getvalue()
                    # Attempts to format stderr output
                    try:
                        if len(myOut) > 0:
                            myOut = myOut[myOut.index("b'") + 1:]
                            myOut = myOut.replace('\\n', '] [')
                            myOut = myOut.replace("'", '')
                            myOut = '[' + myOut + ']'
                            myOut = myOut[0:len(myOut) - 3]
                    except:
                        myOut = buf.getvalue()
                    Data.skippedLines = myOut
            except Exception as e:
                raise Exception('Encountered a fatal error while reading file: {}'.format(e))

            # Dropping unnecessary data from .csv
            df.drop('Cost', axis=1, inplace=True)
            df.drop('Inv', axis=1, inplace=True)
            df.drop('Site', axis=1, inplace=True)
            df.drop('Status', axis=1, inplace=True)
            df.drop('OWT', axis=1, inplace=True)
            df.drop('Live', axis=1, inplace=True)
            df.drop('Revenue', axis=1, inplace=True)

            # Sorting data by Order #
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

            return df
        else:
            # TODO: Delete mail item when done (MailItem.Delete)
            raise Exception('No .csv found!')


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
def updateAccess(filePath):
    # Pyodbc setup
    driver = pyodbc.dataSources()
    driver = driver['MS Access Database']
    connection = pyodbc.connect(driver=driver, dbq=filePath)
    cursor = connection.cursor()

    # Get Default_TEST dataframe
    df = getFileFromEmail()

    # Finds the last record in the dataframe that's also in Access
    maxItem = 0
    for i in reversed(range(0, len(df.index))):
        find = "SELECT * FROM [{}] WHERE [{}] = '{}'".format(Data.tableName, 'Order #', df.iloc[i]['Order #'])
        cursor.execute(find)
        if cursor.fetchone() is not None:
            maxItem = max(maxItem, int(df.iloc[i]['Unnamed: 19']))
            break

    # Limits the dataframe to only the new records that will be placed in access
    df2 = df[df['Unnamed: 19'] > maxItem]
    df2.drop('Unnamed: 19', axis=1, inplace=True)
    df2.drop('End', axis=1, inplace=True)

    # Inserts the new records into Access using SQL syntax (';' not required)
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

    # Saving updates
    connection.commit()
    sendSuccessEmail(Data.skippedLines, df2)


# @Description: Main runner for the program
if __name__ == '__main__':
    # Attempts to update access, sends failure email if any errors occur.
    try:
        if len(sys.argv) < 2:
            raise Exception('Not enough arguments!')
        elif len(sys.argv) > 2:
            Data.tableName = sys.argv[2]
            Data.mailTo = sys.argv[3]
            Data.inboxEmail = sys.argv[4]
        updateAccess(sys.argv[1])
    except Exception as e:
        sendFailureEmail(e, traceback.format_exc())
