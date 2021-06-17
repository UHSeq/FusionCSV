# sort and merge multiple csv files to one single file for import into excel.

import os
from zipfile import ZipFile
import pandas
import openpyxl
from datetime import datetime

cwd = os.getcwd()
now = datetime.now()
nowtime = now.strftime('%H%M%S')
nowdate = now.strftime('%m%d%Y')
excel_file_name = f'fusions_{nowtime}_{nowdate}.xlsx'
excel_path = os.path.join(cwd, 'Samples', excel_file_name)
count_col = 3

zip_files = []
csv_files = set()
csv_objs = {}


class CSVFiles():
    def __init__(self, path):
        self.path = path
        _, self.name = os.path.split(self.path)
        self.name, _ = os.path.splitext(self.name)
        self.read_data()
        self.find_groups()
        self.process_groups()
        self.data.to_csv(os.path.join(cwd, 'Samples', self.name + '.csv'))

    def read_data(self):
        self.data = pandas.read_csv(self.path)

    def find_groups(self):
        self.groupby = self.data.groupby(['Name'])
        self.size = self.groupby.size()
        self.size = dict(self.size)
        self.data.insert(loc=count_col, column='Counts', value=0)

    
    def process_groups(self):
        for index, name in enumerate(self.data['Name']):
            self.data.iat[index, count_col] = self.size[name]

# class Uniques():
#     def __init__(self, name) -> None:
#         self.name = name


def write_pathname(*strings):
    return os.path.join(*strings)

def add_zip(directory, name, ziplist=zip_files):
    zip_path = write_pathname(directory, name)
    ziplist.append(zip_path)

def add_csv(directory, name, csvset=csv_files):
    csv_path = write_pathname(directory, name)
    csvset.add(csv_path)

def unzipper(zipperlist, csvset):
    for zipperfile in zipperlist:
        zipperpath = os.path.dirname(zipperfile)
        with ZipFile(zipperfile, 'r') as Zipper:
            tempset = set(ZipFile.namelist(Zipper))
            for file in tempset:
                csvset.add(os.path.join(zipperpath, file))
            Zipper.extractall(os.path.dirname(zipperfile))

def initialize_csv():
    # workbook = openpyxl.Workbook()
    # workbook.save(excel_path)
    for index, csv in enumerate(csv_files, 1):
        print(f'#### Initializing {index} ####')
        csv_objs[index] = CSVFiles(csv)
    # workbook = openpyxl.load_workbook(excel_path)
    # workbook.remove(workbook['Sheet'])
    # workbook.save(excel_path)

def main_loop():
    for dir, _, file in os.walk(cwd):
        for name in file:
            if name.endswith('.csv'):
                add_csv(dir, name)
            elif name.endswith('.zip'):
                add_zip(dir, name)
    if len(zip_files) > 0:
        unzipper(zip_files, csv_files)
        # pass
    if len(csv_files) > 0:
        initialize_csv()
        # pass

if __name__ in '__main__':
    main_loop()
