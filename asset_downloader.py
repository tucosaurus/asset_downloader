import sys
import requests
import openpyxl
import validators
from six.moves import urllib
from urllib.parse import urlparse
from validators.utils import ValidationFailure


def get_file_from_url(url):
    filename = urlparse(url).path[1:].split('/')[-1]
    print("Downloading file --> ", filename)
    urllib.request.urlretrieve(url, filename)
    print("Finished downloading --> ", filename)


def download_files_from_excel(file_path, column_name):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    for row in sheet.iter_rows():
        for cell in row:
            if cell.column_letter.lower() == column_name.lower():
                cell_value = cell.value
                if not isinstance(cell_value, str):
                    continue
                else:
                    if not cell_value.startswith("http"):
                        continue
                try:
                    validators.url(cell_value)
                    get_file_from_url(cell_value)
                except (ValidationFailure, ValueError):
                    print("*** {} is not a valid url *****".format(cell_value))

if  __name__ == '__main__':
    if len(sys.argv) == 1:
        print("Usage: python asset_downloader.py <file_path> <column_name>")
        sys.exit(1)

    file_path = sys.argv[1]
    column_name = sys.argv[2]
    print("Downloading files from urls in column ", column_name)
    download_files_from_excel(file_path, column_name)
    print("#"*20)
    print("Download complete")
    print("#"*20)
    