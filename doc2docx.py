from win32com import client as wc  # doc转docx
import os  # 读取文件

from setting import DOC_PATH


def save_doc_to_docx(rawpath):  # doc转docx
    '''
    :param rawpath: 传入和传出文件夹的路径# 
        注意：目录的格式必须写成双反斜杠（这里是使用os.path替代）
        path = 'C:\\Users\\Admin\\Desktop\\'
    :return: None
    '''
    # print(wc.Dispatch("wps.Application"))
    # print("....")
    # 自适应wps和word
    try:
        word = wc.Dispatch("wps.Application")
    except:
        word = wc.Dispatch("kwps.Application")
    else:
        word = wc.Dispatch("word.Application")

    word.Visible = 0   # 后台运行
    word.DisplayAlerts = 0    # 不警告

    # 不能用相对路径，老老实实用绝对路径
    # 需要处理的文件所在文件夹目录
    filenamelist = os.listdir(rawpath)
    for i in os.listdir(rawpath):
        # 找出文件中以.doc结尾并且不以~$开头的文件（~$是为了排除临时文件的）
        if i.endswith('.doc') and not i.startswith('~$'):
            # print(i)
            # try
            # 打开文件
            doc = word.Documents.Open(os.path.join(rawpath, i))
            # # 将文件名与后缀分割
            rename = os.path.splitext(i)[0] + '.docx'
            # 将文件另存为.docx
            # print('另存为')
            # print(os.path.join(rawpath, rename))
            doc.SaveAs(os.path.join(rawpath, rename), 12)  # 12表示docx格式
            doc.Close()
    word.Quit()


if __name__ == "__main__":
    print('开始转换...请耐心等候')
    save_doc_to_docx(DOC_PATH)
    print('转换完成')
