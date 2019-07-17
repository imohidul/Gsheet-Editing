import gspread
import csv
import os
from oauth2client.service_account import ServiceAccountCredentials
import shutil
import logging
import time


companies = ['Rupali_Life_Insurance',
             ]

worksheets = ['Policy_File_Valid', 'Policy_Invalid_1', 'Policy_Invalid_2',
              'Payment_Valid', 'Payment_Invalid_1', 'Payment_Invalid_2']


def logging_config():
    logging.basicConfig(filename='gsheet.log', level=logging.INFO,
                        format='%(levelname)s:%(message)s')


def authorize_credential():
    scope = ['https://www.googleapis.com/auth/drive']

    credentials = ServiceAccountCredentials. \
        from_json_keyfile_name('credential.json', scope)
    client = gspread.authorize(credentials)
    return client


def create_csv(cl):
    print("{} will be read".format(len(companies)))
    for i, company in enumerate(companies):
        print("{} {}".format(i, company))
        start_time = int(round(time.time()*1000))
        sheet = cl.open(company)
        if os.path.exists('./' + company):
            shutil.rmtree('./' + company)
        os.makedirs(company)
        for worksheet in worksheets:
            file_name = worksheet + ".csv"
            f = open(company + "/" + file_name, 'w')
            csv.register_dialect('myDialect', quoting=csv.QUOTE_ALL, skipinitialspace=True, lineterminator='\r\n')
            writer = csv.writer(f,dialect='myDialect')
            writer.writerows(sheet.worksheet(worksheet).get_all_values())
            f.close()
        end_time = int(round(time.time() * 1000))
        total_time = start_time-end_time
        ""


if __name__ == '__main__':
    clnt = authorize_credential()
    create_csv(clnt)

