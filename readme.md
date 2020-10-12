## 1、使用前准备

	1. 安装python3
 	2. 安装wps

## 2、使用须知

该脚本运行在`python 3.8.3`测试通过

需要安装的依赖:

```
openpyxl==3.0.5
python-docx==0.8.10
pywin32==228
```

可以直接通过`python -m pip install -r requirements.txt`安装

## 3、使用步骤

1. 先再表格里把所需要的地方打上大写A-Z标记。例如：

   | 姓名 | A    | 性别 | B    |
   | ---- | ---- | ---- | ---- |
   | 星座 | C    | 年龄 | D    |

   然后另存为`模板.docx`，注意一定要是docx，并且和下面设置的**模板文件路径**一致。

2. 先在`setting.py`里设置，修改相应路径

``` python
# 模板文件路径
TEMPLATE_FILE = r'模板.docx'

# 要收集的docx文件夹
DOCX_PATH = r'F:\Python\docx2xlsx\code\docx'

# 要转换为docx的doc文件目录
DOC_PATH = r'F:\Python\docx2xlsx\code\doc'
```

3. docx转excel：执行`python run.py`，文件结果将会保存在`docx2xlsx收集结果.xlsx`文件里。

4. doc转docx：如果文件为doc，需要先将doc 转换为 docx格式：执行`python doc2docx.py`，转换的docx保存在doc所在文件夹下，该转换过程需要一点时间，请耐心等候；然后再执行第二步。

