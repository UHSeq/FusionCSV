# sort and merge multiple csv files to one single file for import into excel.

import os
from zipfile import ZipFile
import zipfile

cwd = os.getcwd()
# print(cwd, type(cwd))

zip_files = []
csv_files = set()

def write_pathname(*strings):
    return os.path.join(*strings)

# def find_files(path, ending, hold):
#     pass
#     for dir, _, file in os.walk(path):
#         for name in file:
#             if name.endswith(ending):
                

def test_loop():
    for dir, _, file in os.walk(cwd):
        for name in file:
            if name.endswith('.zip'):
                add_zip(dir, name)
                # print(ZipFile.namelist(os.path.join(dir, name)))
            elif name.endswith('.csv'):
                add_csv(dir, name)
    if len(zip_files) > 0:
        unzipper(zip_files, csv_files)
        # pass
    if len(csv_files) > 0:
        pass

def add_zip(directory, name, ziplist=zip_files):
    zip_path = write_pathname(directory, name)
    ziplist.append(zip_path)
    # pass

def add_csv(directory, name, csvset=csv_files):
    csv_path = write_pathname(directory, name)
    # csvset.add(csv_path)
    pass

def unzipper(zipperlist, csvset):
    # pass
    for zipperfile in zipperlist:
        # print(zipperfile)
        zipperpath = os.path.dirname(zipperfile)
        with ZipFile(zipperfile, 'r') as Zipper:
            # zipperpath = os.path.dirname(zipperfile)
            tempset = set(ZipFile.namelist(Zipper))
            for file in tempset:
                csvset.add(os.path.join(zipperpath, file))
            # csvset.update()
            # tempset = set(templist)
            # print(ZipFile.namelist(Zipper), type(ZipFile.namelist(Zipper)))
            Zipper.extractall(os.path.dirname(zipperfile))
            # Zipper.extractall(os.path.join())
    # for file in 
    # print(f'the last zip file is named {zipfile}')


if __name__ in '__main__':
    test_loop()
    print(zip_files)
    print(csv_files)
