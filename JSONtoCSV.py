import pandas as pd
import json
import numpy as np
from openpyxl import load_workbook


def read_json(filename: str) -> dict:

    try:
        with open(filename, "r") as f:
            data = json.loads(f.read())
    except:
        raise Exception(f"Reading {filename} file encountered an error")
    return data

def create_dataframe(data: list) -> pd.Dataframe:

    # declare empty dataframe to append records
    dataframe = pd.DataFrame()

    # loop through each record
    for d in data:

        # normalize the column levels

        # record = pandas.json_normalize(d)
        record1 = pd.json_normalize(d, ['credentials', 'certs'],
                                    errors='ignore', record_prefix='')
        record1.head()

        # use series instead of datafram for dict
        record2 = pd.Series(d, ['name'], name='Connection_Name')
        # convert to dataframe
        record2.reset_index(drop=True, inplace=True)
        # record2.head()

        # concatenate both dataframes
        record = pd.concat([record1, record2], axis=1)

        # append it to the datafram
        dataframe = dataframe.append(record, ignore_index=True)
    
    return dataframe

def add_excel_sheet():
    filepath = "C:/Users/845730829/Documents/Pycharm/Python Script/Connect SSO QA&Prod active connections.xlsx"

    # generate workbook
    workbook = load_workbook(filepath)

    # generate writer engine
    writer = pd.ExcelWriter(filepath, engine='openpyxl', mode='a')

    # assigning the workbook to the writer engine
    writer.book = workbook

    # create dataframe
    df1 = pd.read_csv("certs.csv")

    # add the excel file as a new sheet
    df1.to_excel(writer, sheet_name='Certs')
    writer.save()
    writer.close()

def main():
    # read json file as python dictionary
    data = read_json(filename="C:/Users/845730829/Documents/Pycharm/Python Script/response.json")

    # generate the dataframe for the array items in key
    df = create_dataframe(data=data['items'])

    # rename columns of the dataframe
    # print("Normalized Columns:", df.columns.to_list())

    # remove rose with null values
    # df.dropna(inplace=True)

    # change the column order by passing a list
    df = df[['Connection_Name', 'primaryVerificationCert,', 'secondaryVerificationCert', 'activeVerificationCert', 'encryptionCert', 'certview.id', 'certView.serialNumber', 'certView.subjectDN',
             'certView.subjectAlternativeNames', 'certView.issuerDN', 'certView.validForm', 'certView.expires', 'certView.keyAlgorithm', 'certView.keySize', 'certView.signatureAlgorithm',
             'certView.version', 'certView.sha1Fingerprint', 'certView.sha256Fingerprint', 'certView.status', 'x509File.id', 'x509File.fileData']]
    
    # convert dataframe to csv
    # df.drop_duplicates()
    df.to_csv("certs.csv", index=False)

    add_excel_sheet()


if __name__ == '__main__':
    main()