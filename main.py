# @Author: Ethan Meyers
# @Email: ewm1230@gmail.com
# @Phone: 847-212-2264
# @Created Date: 07/13/2022

import pandas as pd
import pyodbc as pyodbc

def getFile():
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
        end.append(int(string[slash+1:]))
    df['End'] = end
    df = df.sort_values(by=['Unnamed: 19', 'End'], ignore_index=True)
    return df


def updateAccess():
    filePath = r"C:\Users\emeyers\Desktop\EthanAccess.accdb"
    driver = pyodbc.dataSources()
    driver = driver['MS Access Database']
    connection = pyodbc.connect(driver=driver, dbq=filePath)
    cursor = connection.cursor()
    tableName = 'Pick Up 2022 Cont'

    dfColumns = 'User', 'EDI', 'Order Date', 'Order #', 'Container #', \
                'Master BOL/Booking Ref', 'Customer', 'Customer Ref', 'Pick Up', \
                'Delivery', 'DL City', 'Revenue', 'Cost', 'Inv', 'Live', 'OWT', \
                'Status', 'Site'

    df = getFile()
    lastUpdate = 0
    maxItem = 0
    for i in range(0, len(df.index)):
        if int((i / len(df.index)) * 100) != lastUpdate:
            lastUpdate = int((i / len(df.index)) * 100)
            print("{}%".format(lastUpdate))
        find = "SELECT * FROM [{}] WHERE [{}] = '{}'".format(tableName, 'Order #', df.iloc[i]['Order #'])
        cursor.execute(find)
        if cursor.fetchone() is not None:
            maxItem = max(maxItem, int(df.iloc[i]['Unnamed: 19']))

    # maxIndex = max(df.index[df['Unnamed: 19'] == maxItem])
    df2 = df[df['Unnamed: 19'] > maxItem]
    df2.drop('Unnamed: 19', axis=1, inplace=True)
    df2.drop('End', axis=1, inplace=True)


if __name__ == '__main__':
    updateAccess()




