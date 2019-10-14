from azure.storage.blob import BlockBlobService
from conversion.excel2csv import Excel2Csv


converter = Excel2Csv(BlockBlobService(account_name='zhenzhadfblobsource', account_key='ZNqO5INeOxwUd9RyK0kbOy/Z5HhEKOBgW5l1Jg59t/IKvKSyvqTjwSHfXF3Oyuab3cVACcHOhB/fPIvekA7new=='))
converter.convert_and_upload_blob("apachelog-analysis", "apachelog-analysis", "raw/apache.access.log.xlsx", "curated")  