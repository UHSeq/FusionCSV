
import os
from zipfile import ZipFile
import pandas
import openpyxl
from datetime import datetime
# import time

cwd = os.getcwd()
now = datetime.now()
nowtime = now.strftime('%H%M%S')
nowdate = now.strftime('%m%d%Y')
excel_file_name = f'fusions_{nowtime}_{nowdate}.xlsx'
output_dir = 'Output'
csv_hold = 'csv_hold'
excel_path = os.path.join(cwd, output_dir, excel_file_name)
count_col = 3
fusion_sheet = 'Total Fusions'

zip_files = []
csv_files = set()
csv_objs = {}
outter_fusion_class = {}

class CSVFiles():
    def __init__(self, path):
        self.path = path
        _, self.name = os.path.split(self.path)
        self.name, _ = os.path.splitext(self.name)
        self.read_data()
        # self.find_groups()
        self.isolate_uniques()

    def read_data(self):
        self.data = pandas.read_csv(self.path)

    def find_groups(self):
        self.groupby = pandas.DataFrame(self.data.groupby(['Name']))
        self.groupby.to_csv(os.path.join(cwd, csv_hold, self.name + "_dump.csv"))

    def isolate_uniques(self):
        self.unique_fusions = self.data.drop_duplicates(subset=['Name'])
        self.unique_fusions = self.unique_fusions['Name']
        self.unique_fusions = pandas.DataFrame(self.unique_fusions)
        self.unique_fusions.insert(loc=1, column=self.name, value=1)

class MasterData():
    def __init__(self, path):
        self.path = path
        self.create_excel()
        self.csv_files = {}
        self.fusions = pandas.DataFrame()
        self.csv_dict()
        self.fusion_count()
        with pandas.ExcelWriter(self.path, mode='a', engine='openpyxl') as Excel_file:
            self.print2excel(Excel_file, fusion_sheet, self.fusions)
            for csv in self.csv_files.values():
                self.print2excel(Excel_file, csv.name, csv.data)
        self.remove_sheet()

    def csv_dict(self):
        for index, csv in enumerate(csv_files, 1):
            self.csv_files[index] = CSVFiles(csv)

    def create_excel(self):
        workbook = openpyxl.Workbook()
        workbook.save(self.path)

    def remove_sheet(self):
        self.workbook = openpyxl.load_workbook(self.path)
        self.workbook.remove(self.workbook['Sheet'])
        self.workbook.save(self.path)

    def rename_sheet(self):
        self.wso = self.workbook['Sheet']
        self.wso.title = fusion_sheet
        self.workbook.save(self.path)
        # Would prefer to rename and use the sheet as opposed to just deleting it
        # but for some reason this doesn't working that way. Keeping it in anyway.
    
    def fusion_count(self):
        for index, csv in self.csv_files.items():
            if index == 1:
                self.fusions = csv.unique_fusions
            else:
                self.fusions = pandas.merge(self.fusions, csv.unique_fusions, how='outer')
        self.fusions = self.fusions.set_index('Name')
        self.fusions.insert(0, 'Total Occurance', self.fusions.sum(axis=1))

    def print2excel(self, excel_obj, excel_sheet, input_obj):
        input_obj.to_excel(excel_obj, excel_sheet)

    # def print_excel_sheet(self, excel_output, obj, sheet_name):
    #     print(f"Printing to excel data set {sheet_name}")
    #     startime = time.time()
    #     with pandas.ExcelWriter(excel_output, mode='a', engine='openpyxl') as Excel_file:
    #         obj.to_excel(Excel_file, sheet_name)
    #     totaltime = time.time() - startime
    #     print(f'It took {totaltime} seconds to complete')
    # Keep this here becuase it's useful, not using it because it's slow.
    # Doing this one sheet at time took over a minute during testing, while opening
    # the excel file once and plugging in all components at once took about 5 seconds.
    # Moral of the story: open a file and add to it a minimum number of times.

def add_zip(directory, name, ziplist=zip_files):
    zip_path = write_pathname(directory, name)
    ziplist.append(zip_path)

def add_csv(directory, name, csvset=csv_files):
    csv_path = write_pathname(directory, name)
    csvset.add(csv_path)

def clear_folder(dir_path):
    pass
    for root, _, files in os.walk(dir_path):
        for file in files:
            os.remove(os.path.join(root, file))

def unzipper(zipperlist, csvset):
    for zipperfile in zipperlist:
        zipperpath = os.path.dirname(zipperfile)
        with ZipFile(zipperfile, 'r') as Zipper:
            tempset = set(ZipFile.namelist(Zipper))
            for file in tempset:
                csvset.add(os.path.join(zipperpath, file))
            Zipper.extractall(os.path.dirname(zipperfile))

def write_pathname(*strings):
    return os.path.join(*strings)

def main_loop():
    clear_folder(output_dir)
    clear_folder(csv_hold)
    for dir, _, file in os.walk(cwd):
        for name in file:
            if name.endswith('.csv'):
                add_csv(dir, name)
            elif name.endswith('.zip'):
                add_zip(dir, name)
    if len(zip_files) > 0:
        unzipper(zip_files, csv_files)
    if len(csv_files) > 0:
        m = MasterData(excel_path)
    else:
        print("####\nOpps... did you mean to add some csv files?\nNothing here...\n####")

if __name__ in '__main__':
    main_loop()
