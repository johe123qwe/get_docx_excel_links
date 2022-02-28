## 项目说明

读取文件夹下的文档和表格，获取odt、ods、docx、xlsx中的超链接。通过 [cloudmersive](https://cloudmersive.com/tools) 进行转换。

## deploy
`pip3 install -r requirements.txt`

### register api
[cloudmersive api](https://cloudmersive.com/convert-api)

`main.py` 为命令行版。

`main_gui.py` 为通过 [Geooey](https://github.com/chriskiehl/Gooey) 生成的桌面版本，通过 `pyinstaller` 打包生成桌面包。

MacOS or Linux

`pyinstaller -w -F main_gui.py`

Windows

[Auto PY to EXE](https://pypi.org/project/auto-py-to-exe/)


## 使用

`python3 main.py -d /to/path/`

## 特别鸣谢
[hx-util](https://github.com/Colin-Fredericks/hx-util)

[关于cloudmersive安全文档](https://cloudmersive.com/security)