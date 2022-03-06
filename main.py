# -*- coding:UTF-8 -*-

# @AUTHOR: xiaoyu
# @DATE: 2022/03/05 Sat
# @TIME: 10:50:00
#
# @DESCRIPTION: 获取odt、ods、docx、xlsx中的超链接,命令行版。
# modified by xy 20220216 1.0.0
# modified by xy 20220225 1.0.1
# modified by xy 20220301 1.0.2  把 KEY 设置为参数导入，而不是写在脚本里。
# modified by xy 20220305 1.0.3  Modify comments。
# modified by xy 20220306 1.0.4  添加默认 KEY，可不用每次都使用 -k 参数。

version = '1.0.4'

import argparse
import shutil
import GetWordLinks
import GetExcelLinks
import convert_ods_xlsx
import convert_odt_docx
import shutil
import os


'''
pip install cloudmersive-validate-api-client
pip install cloudmersive_convert_api_client
pip install bs4
pip install lxml
pip install pandas
pip install unicodecsv
'''

docxs_list = []
xlsxs_list = []


def get_file(args_path, KEY):
    '''Get a list of files to deposit'''
    for root, dirs, files in os.walk(args_path):
        files = [f for f in files if not f[0] == '.']
        dirs = [d for d in dirs if not[0] == '.']
        for eachfile in files:
            if eachfile.endswith(".docx") and not eachfile.startswith("~$"):
                docx_path = os.path.join(root, eachfile)
                docxs_list.append(docx_path)
            if eachfile.endswith('.odt') and not eachfile.startswith('~$'):
                pass # Convert to docx
                tmp_path = os.path.join(root, '.~$__tmp__')
                if not os.path.exists(tmp_path):
                    os.makedirs(tmp_path)
                docx_path = os.path.join(root, eachfile)
                new_docx = os.path.join(tmp_path, os.path.basename(docx_path) + '.docx')
                convert_odt_docx.convent_odt_to_docx(docx_path, new_docx, KEY)
                docxs_list.append(new_docx)
            if eachfile.endswith('.xlsx') and not eachfile.startswith("~$"):
                xlsx_path = os.path.join(root, eachfile)
                xlsxs_list.append(xlsx_path)
            if eachfile.endswith('.ods') and not eachfile.startswith("~$"): # 转 ods to xlsx
                tmp_path = os.path.join(root, '.~$__tmp__')
                if not os.path.exists(tmp_path):
                    os.makedirs(tmp_path)
                xlsx_path = os.path.join(root, eachfile)
                new_xlsx = os.path.join(tmp_path, os.path.basename(xlsx_path) + '.xlsx')
                convert_ods_xlsx.convent_ods_to_xlsx(xlsx_path, new_xlsx, KEY)
                xlsxs_list.append(new_xlsx)

def delete_tmp_path(args_path):
    for x in os.walk('.'):
        if (os.path.basename(x[0])) == '.~$__tmp__':
            try:
                shutil.rmtree(x[0])
            except Exception as e:
                print('Unable to delete temporary directory')

def main():
    parser = argparse.ArgumentParser(description="Get document hyperlink widget")
    parser.add_argument('-d', '--directory', dest='args_path', action='store', required=True, help='路径')
    parser.add_argument('-k', '--key', dest='KEY', action='store', required=False, default='YOUR_KEY', help='cloudmersiv_key, \
                        如果不添加 -k 则使用默认值 KEY。')

    args = parser.parse_args()
    global KEY
    KEY = args.KEY
    get_file(args.args_path, KEY)
    if docxs_list != []:
        GetWordLinks.getWordLinks(docxs_list, args.args_path)
    if xlsxs_list != []:
        GetExcelLinks.getExcelLinks(xlsxs_list, args.args_path)
    delete_tmp_path(args.args_path)

if __name__ == '__main__':
    main()