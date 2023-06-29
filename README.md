# 项目介绍

本项目实现了文档的解析功能，能够基本满足`PDF/Microsoft Word`的文档目录生成、版面解析、表格抽取以及图片抽取等功能。

# 环境依赖

```
pandas==1.4.4
pdfminer.six==20221105
pdfplumber==0.9.0
PyMuPDF==1.22.3
python_docx==0.8.11
PyPDF2==3.0.1
```

# 部署方式

### 安装依赖

在本项目目录下执行

```
pip install -r requirements.txt
```

所需要的依赖库将会自动安装成功

### 项目部署

本项目仅采用文件的形式，因而仅需要在`.py`文件目录下打开命令行执行命令就可以

# 使用方式

在项目目录下执行对应文件即可。对于PDF文件采用`pdfExtractor.py`，对于Microsoft Word文件采用`docExtractor.py`，可选项可通过`-h`确定。

如下为`pdfExtractor.py`的可选项以及介绍，其中`OPTION`的默认选项为`text`

```
usage: python pdfExtractor.py [-h] [-o OPTION] [-p PATH] [-r RESULT_PATH]

PDF Extraction

optional arguments:
  -h, --help            show this help message and exit
  -o OPTION, --option OPTION
                        The function will be served, optional choice will be text(Extract text), image(Extract
                        images), table(Extract tables) and toc(Extract tocs)
  -p PATH, --path PATH  Path to the pdf file
  -r RESULT_PATH, --result_path RESULT_PATH
                        Path to store the result images
```

如下为`docExtractor.py`的可选项以及介绍，其中`OPTION`的默认选项为`text`

```
usage: python docExtractor.py [-h] [-o OPTION] [-p PATH] [-r RESULT_PATH]

Microsoft Word Extraction

optional arguments:
  -h, --help            show this help message and exit
  -o OPTION, --option OPTION
                        The function will be served, optional choice will be text(Extract text), image(Extract
                        images), table(Extract tables) and toc(Extract tocs)
  -p PATH, --path PATH  Path to the doc file
  -r RESULT_PATH, --result_path RESULT_PATH
                        Path to store the result images
```

# 作者列表

Heart_1ess (https://Heart_1ess@bupt.edu.cn)