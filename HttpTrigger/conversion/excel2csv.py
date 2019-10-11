import csv
import os
import sys

import xlrd


class File:
    def __init__(self, filename, filepath):
        self.filename = filename
        self.filepath = filepath


class Excel2Csv:
    def __init__(self, block_blob_service):
        self.block_blob_service = block_blob_service             
        self.downloaded_dir = os.getcwd() + "\\downloadedFromBlob\\"

    def convert(self, excel_file, csv_file_base_path):
        workbook = xlrd.open_workbook(excel_file)    
        csv_files = []

        for sheet_name in workbook.sheet_names():
            print('processing - ' + sheet_name)
            csv_file = self.convert_worksheet(workbook, sheet_name, csv_file_base_path)
            csv_files.append(csv_file)
        return csv_files

    def convert_worksheet(self, workbook, sheet_name, csv_file_base_path):        
        worksheet = workbook.sheet_by_name(sheet_name)
        csv_file_name = sheet_name.lower().replace(" - ", "_").replace(" ","_") + '.csv'
        csv_file_full_path = csv_file_base_path + csv_file_name
        csv_file = open(csv_file_full_path, 'w', encoding='utf8')
        write_to_csv = csv.writer(csv_file, quoting = csv.QUOTE_ALL)
        for row_num in range(worksheet.nrows):
            write_to_csv.writerow(worksheet.row_values(row_num))
        csv_file.close()
        return File(csv_file_name, csv_file_full_path)

    def convert_and_upload(self, source_container, dest_container):
        generator = self.block_blob_service.list_blobs(source_container)
        
        for blob in generator:
            self.convert_and_upload_blob(source_container, dest_container, blob.name)

    def convert_and_upload_blob(self, source_container, dest_container, blob_name):
        downloaded_excel = self.download(source_container, blob_name)
        # conversion
        csv_files = self.convert(excel_file = downloaded_excel.filename, csv_file_base_path = downloaded_excel.filepath)     
        self.upload(dest_container, csv_files)
        
    def download(self, source_container, blob_name):
        if not os.path.exists(self.downloaded_dir):
            os.makedirs(self.downloaded_dir)

        downloaded_file = self.downloaded_dir + blob_name
        # download excel from blob
        self.block_blob_service.get_blob_to_path(source_container, blob_name, downloaded_file)
        return File(downloaded_file, self.downloaded_dir)

    def upload(self, dest_container, csv_files):
        for csv_file in csv_files:
            self.block_blob_service.create_blob_from_path(dest_container, csv_file.filename, csv_file.filepath)