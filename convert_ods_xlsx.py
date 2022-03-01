from __future__ import print_function
import time
import cloudmersive_convert_api_client
from cloudmersive_convert_api_client.rest import ApiException
from pprint import pprint




def convent_ods_to_xlsx(input_file, outputfile, KEY):

    # Configure API key authorization: Apikey
    configuration = cloudmersive_convert_api_client.Configuration()
    configuration.api_key['Apikey'] = KEY
    # create an instance of the API class
    api_instance = cloudmersive_convert_api_client.ConvertDocumentApi(cloudmersive_convert_api_client.ApiClient(configuration))

    try:
        # Convert ODS Spreadsheet to XLSX
        api_response = api_instance.convert_document_ods_to_xlsx(input_file)
        time.sleep(0.1)
        # pprint(api_response)
        with open(outputfile, 'wb') as fp:
            fp.write(api_response)
    except ApiException as e:
        print("Exception when calling ConvertDocumentApi->convert_document_ods_to_xlsx: %s\n" % e)

if __name__ == '__main__':
    convent_ods_to_xlsx(input_file, outputfile, KEY)
    