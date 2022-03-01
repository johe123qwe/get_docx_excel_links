from __future__ import print_function
import time
import cloudmersive_convert_api_client
from cloudmersive_convert_api_client.rest import ApiException
from pprint import pprint





def convent_odt_to_docx(input_file, outputfile, KEY):

    # Configure API key authorization: Apikey
    configuration = cloudmersive_convert_api_client.Configuration()
    configuration.api_key['Apikey'] = KEY

    # create an instance of the API class
    api_instance = cloudmersive_convert_api_client.ConvertDocumentApi(cloudmersive_convert_api_client.ApiClient(configuration))

    try:
        # Convert Office Open Document ODT to Word DOCX
        api_response = api_instance.convert_document_odt_to_docx(input_file)
        time.sleep(0.1)
        with open(outputfile, 'wb') as fp:
            fp.write(api_response)
    except ApiException as e:
        print("Exception when calling ConvertDocumentApi->convert_document_odt_to_docx: %s\n" % e)

if __name__ == '__main__':
    convent_odt_to_docx(input_file, outputfile, KEY)
    