# -*- coding:UTF-8 -*-

# @AUTHOR: xiaoyu
# @DATE: 2022/02/25 Fri
# @TIME: 16:09:00
#
# @DESCRIPTION: 获取odt、ods、docx、xlsx中的超链接,采用 Gooey 打包成 GUI 工具。
# modified by xy 20220216 1.0.0
# modified by xy 20220225 1.0.1

version = '1.0.1'
import shutil
import GetWordLinks
import GetExcelLinks
import convert_ods_xlsx
import convert_odt_docx
import shutil
import os
from gooey import Gooey, GooeyParser
import sys

utf8_stdout = os.fdopen(sys.stdout.fileno(), mode='w', encoding='utf-8', closefd=False)
sys.stdout = utf8_stdout


'''
# cloudmersiv API 每月 800 次，
https://cloudmersive.medium.com/how-to-convert-ods-to-xlsx-in-python-de3ae2930cd1
https://api.cloudmersive.com/python-client.asp
https://github.com/Cloudmersive/Cloudmersive.APIClient.Python.Validate/blob/master/demo/demo.py

备用：https://developers.zamzar.com/docs

另一种获取超链接方式：http://webcache.googleusercontent.com/search?q=cache:iawH2oQHZOMJ:www.cppcns.com/jiaoben/python/379590.html&hl=zh-CN&gl=sa&strip=0&vwsrc=0
'''

docxs_list = []
xlsxs_list = []

def get_file(args_path):
    '''获取文件存入列表'''
    for root, dirs, files in os.walk(args_path):
        files = [f for f in files if not f[0] == '.']
        dirs = [d for d in dirs if not[0] == '.']
        for eachfile in files:
            if eachfile.endswith(".docx") and not eachfile.startswith("~$"):
                docx_path = os.path.join(root, eachfile)
                docxs_list.append(docx_path)
            if eachfile.endswith('.odt') and not eachfile.startswith('~$'):
                pass # 此处转换为 docx
                tmp_path = os.path.join(root, '.~$__tmp__')
                if not os.path.exists(tmp_path):
                    os.makedirs(tmp_path)
                docx_path = os.path.join(root, eachfile)
                new_docx = os.path.join(tmp_path, os.path.basename(docx_path) + '.docx')
                convert_odt_docx.convent_odt_to_docx(docx_path, new_docx)
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
                convert_ods_xlsx.convent_ods_to_xlsx(xlsx_path, new_xlsx)
                xlsxs_list.append(new_xlsx)

def delete_tmp_path(args_path):
    for x in os.walk('.'):
        if (os.path.basename(x[0])) == '.~$__tmp__':
            try:
                shutil.rmtree(x[0])
            except Exception as e:
                print('无法删除临时目录')

@Gooey(
    program_name='获取文档超链接小工具' + version,
    default_size=(700, 550),
    # 调整界面颜色
    header_bg_color='#599171',
    body_bg_color='#599171',
    footer_bg_color='#599171',
    encoding='utf-8',
    language='chinese',
    menu=[{
        'name': 'File',
        'items': [{
                'type': 'AboutDialog',
                'menuTitle': 'About',
                'name': '获取文档超链接小工具',
                'description': '获取文档超链接小工具',
                'version': '1.0',
                'copyright': '@2022',
                'website': 'https://github.com/chriskiehl/Gooey',
                'developer': 'xxx',
                'license': 'MIT'
            }, {
                'type': 'Link',
                'menuTitle': 'other tools',
                'url': 'https://drive.google.com/drive/folders/'
            }]
        },{
        'name': 'Help',
        'items': [{
            'type': 'Link',
            'menuTitle': '小工具链接',
            'url': 'https://drive.google.com/drive/folders/'
        }]
    }]
)

def main():
    parser = GooeyParser(description="获取文档超链接小工具")
    parser.add_argument('-d', '--directory', dest='args_path', metavar='路径', action='store', required=True, widget='DirChooser', help='路径')
    parser.add_argument('-k', '--key', dest='KEY', metavar='cloudmersiv_key', action='store', required=True, widget='TextField', help='cloudmersiv_key')

    args = parser.parse_args()

    get_file(args.args_path)
    if docxs_list != []:
        GetWordLinks.getWordLinks(docxs_list, args.args_path)
    if xlsxs_list != []:
        GetExcelLinks.getExcelLinks(xlsxs_list, args.args_path)
    delete_tmp_path(args.args_path)

if __name__ == '__main__':
    main()