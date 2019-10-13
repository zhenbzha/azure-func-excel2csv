import json
import logging

import azure.functions as func
from azure.keyvault import KeyVaultClient
from azure.storage.blob import BlockBlobService
from msrestazure.azure_active_directory import MSIAuthentication

from .conversion.excel2csv import Excel2Csv


def main(req: func.HttpRequest) -> func.HttpResponse:

    req_body = req.get_json()
    source_container = req_body.get("source_container")
    source_blob_name = req_body.get("source_blob")
    dest_folder = req_body.get("dest_folder")
    dest_container = req_body.get("dest_container")

    credentials = MSIAuthentication(resource='https://vault.azure.net')
    kvclient = KeyVaultClient(credentials)
    key = kvclient.get_secret("https://zhenzh-databricks-kv.vault.azure.net/", "zhenzh-python-func-account-key", "").value
    
    converter = Excel2Csv(BlockBlobService(account_name='zhenzhadfblobsource', account_key=key))

    if source_blob_name is None:        
        converter.convert_and_upload(source_container, dest_container, "raw", "curated")
    else:
        converter.convert_and_upload_blob("apachelog-analysis", "apachelog-analysis", source_blob_name, dest_folder)     
    
    return json.dumps({"result": "Conversion Finished!"})