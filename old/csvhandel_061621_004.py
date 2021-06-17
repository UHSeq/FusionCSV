# sort and merge multiple csv files to one single file for import into excel.

import os
from zipfile import ZipFile
import pandas
import openpyxl

cwd = os.getcwd()
excel_path = os.path.join(cwd, 'Samples', 'fusions.xlsx')

zip_files = []
csv_files = set()
csv_objs = {}


class CSVFiles():
    def __init__(self, path):
        self.path = path
        _, self.name = os.path.split(self.path)
        self.name, _ = os.path.splitext(self.name)
        self.read_data()
        self.find_unique()

    def read_data(self):
        self.data = pandas.read_csv(self.path)

    def find_unique(self):
        self.unique_counts = self.data.groupby(['Name']).size()
        with pandas.ExcelWriter(excel_path, mode='a') as Excel_File:
            self.unique_counts.to_excel(Excel_File, self.name)


class Uniques():
    def __init__(self, name) -> None:
        self.name = name


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
    workbook = openpyxl.Workbook()
    workbook.save(excel_path)
    for index, csv in enumerate(csv_files):
        csv_objs[index] = CSVFiles(csv)
    workbook = openpyxl.load_workbook(excel_path)
    workbook.remove(workbook['Sheet'])
    workbook.save(excel_path)

def main_loop():
    for dir, _, file in os.walk(cwd):
        for name in file:
            if name.endswith('.csv'):
                add_csv(dir, name)
            elif name.endswith('.zip'):
                add_zip(dir, name)
    if len(zip_files) > 0:
        unzipper(zip_files, csv_files)
    if len(csv_files) > 0:
        initialize_csv()

if __name__ in '__main__':
    main_loop()
