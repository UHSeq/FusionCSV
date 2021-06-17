# sort and merge multiple csv files to one single file for import into excel.

import os
from zipfile import ZipFile
import pandas
# import zipfile

cwd = os.getcwd()
# print(cwd, type(cwd))

zip_files = []
csv_files = set()
csv_objs = {}

class CSVFiles():
    def __init__(self, path):
        # pass
        self.path = path
        self.read_data()
        self.find_unique()

    def read_data(self):
        self.data = pandas.read_csv(self.path)
        # print(type(self.data))

    def find_unique(self):
        # print(type(self.data))
        self.unique_names = self.data.Name.unique()

def write_pathname(*strings):
    return os.path.join(*strings)

def add_zip(directory, name, ziplist=zip_files):
    zip_path = write_pathname(directory, name)
    ziplist.append(zip_path)
    # pass

def add_csv(directory, name, csvset=csv_files):
    csv_path = write_pathname(directory, name)
    csvset.add(csv_path)
    # pass

def unzipper(zipperlist, csvset):
    for zipperfile in zipperlist:
        zipperpath = os.path.dirname(zipperfile)
        with ZipFile(zipperfile, 'r') as Zipper:
            tempset = set(ZipFile.namelist(Zipper))
            for file in tempset:
                csvset.add(os.path.join(zipperpath, file))
            Zipper.extractall(os.path.dirname(zipperfile))

def initialize_csv():
    # pass
    for index, csv in enumerate(csv_files):
        csv_objs[index] = CSVFiles(csv)
        # print(index, csv)


def test_loop():
    for dir, _, file in os.walk(cwd):
        for name in file:
            if name.endswith('.csv'):
                add_csv(dir, name)
            elif name.endswith('.zip'):
                add_zip(dir, name)
                # print(ZipFile.namelist(os.path.join(dir, name)))
    if len(zip_files) > 0:
        unzipper(zip_files, csv_files)
    if len(csv_files) > 0:
        initialize_csv()
        # pass

if __name__ in '__main__':
    test_loop()
    # print(zip_files)
    # print(csv_files)
