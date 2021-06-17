# sort and merge multiple csv files to one single file for import into excel.

import os
from zipfile import ZipFile

cwd = os.getcwd()
# print(cwd, type(cwd))

zip_files = []
csv_files = {}

def write_pathname(*strings):
    return os.path.join(*strings)


def test_loop():
    # csv_dir = None
    for dir, _, file in os.walk(cwd):
        for name in file:
            # print(f'root directory is {dir} for {file} and type is {type(dir)}')
            if name.endswith('.zip'):
                add_zip(dir, name)
                # print(os.path.join(cwd, name))
                # zip_path = write_pathname(dir, name)
                # zip_files.append(zip_path)
                # print(os.path.)
            elif name.endswith('.csv'):
                add_csv(dir, name)
                # print(os.path.join(cwd, name))
                # csv_path = write_pathname(dir, name)
                # csv_files.add(csv_path)
    if len(zip_files) > 0:
        unzipper(zip_files)
        # pass
        # print(zip_files)
    if len(csv_files) > 0:
        pass

def add_zip(directory, name, ziplist=zip_files):
    zip_path = write_pathname(directory, name)
    ziplist.append(zip_path)
    # pass

def add_csv(directory, name, csvset=csv_files):
    csv_path = write_pathname(directory, name)
    csvset.add(csv_path)
    pass

def unzipper(zipperlist, csvset):
    # pass
    for zipfile in zipperlist:
        with ZipFile(zipfile, 'r') as Zipper:
            Zipper.extractall(os.path.dirname(zipfile))
            # Zipper.extractall(os.path.join())
    # for file in 

if __name__ in '__main__':
    test_loop()
    print(csv_files)
    print(zip_files)