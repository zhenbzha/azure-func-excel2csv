import json
import logging

import azure.functions as func
from azure.storage.blob import BlockBlobService

from .conversion.excel2csv import Excel2Csv

from azure.keyvault import KeyVaultClient
from msrestazure.azure_active_directory import MSIAuthentication
 

def main(req: func.HttpRequest) -> func.HttpResponse:
    credentials = MSIAuthentication(resource='https://vault.azure.net')
    kvclient = KeyVaultClient(credentials)
    key = kvclient.get_secret("https://zhenzh-databricks-kv.vault.azure.net/", "zhenzh-python-func-account-key", "").value
    
    converter = Excel2Csv(BlockBlobService(account_name='zhenzhadfblobsource', account_key=key))
    converter.convert_and_upload()         
    
    return json.dumps({"result": "Conversion Finished!"})