import tkinter

from tkinter import simpledialog
import time

import win32api
import win32com.client
import os
import shutil
import tkinter.simpledialog
import tkinter.filedialog

#today=time.time()
today_str=time.strftime('%Y.%m.%d',time.localtime(time.time()))
print(today_str)
Review_path = 'D:/个人文件/桌面/审查单/合理审查单（公路运输）.docx'

def askname():
    result = tkinter.simpledialog.askstring(title='业务编号', prompt='请输入业务编号：', initialvalue='在此处输入')
    print(result)
    return result

def file_rename(path, todir):
    r=askname()
    filelist = os.listdir (path)
    i = 0
    for files in filelist:

        Olddir = os.path.join(path, files)

        if os.path.isdir(Olddir):

            sub_path = Olddir
            sub_filelist = os.listdir(sub_path)
            for sub_files in sub_filelist:
                if sub_files.find ('Invoice_') >= 0:
                    filename = os.path.splitext (sub_files)[0]  # 文件名
                    newname = files
                    newpath = sub_path + '/'
                    # os.rename(newpath + filename + ".rtf", newpath + )
                    ne = shutil.copy(newpath + filename + ".rtf", os.path.join(todir, newname + ".rtf"))
                    word = win32com.client.DispatchEx('kwps.Application')
                    #sc_doc = word.Documents.Open (paths)
                    word.DisplayAlerts = 0
                    word.Visible = 0
                    print(ne)

            #         (file_path, temp_file_name) = os.path.split (ne)
            #         (short_name, extension) = os.path.splitext (temp_file_name)
            #         print(short_name)
            #
            #         doc = word.Documents.Open(ne)
            #         save = doc.SaveAs(todir + '/' + short_name + ".docx", 16)  # 另存为后缀为".doc"的文件，其中参数0指doc文件
            #         doc.Close()
            # #word.Quit ()
            #         newdoc = Document(save)
            #         sections = newdoc.sections
            #         for section in sections:
            #             section.top_margin = Cm(1)
            #             section.bottom_margin = Cm(1)
            #             section.left_margin = Cm(1)
            #             section.right_margin = Cm(1)
            #         #newdoc.PrintOut(newdoc)
            #         newdoc.save(todir + '/' + short_name + ".doc")
                    Review_doc = word.Documents.Open(Review_path)
                    word.Selection.Find.Execute("文件名", False, False, False, False, False, True, 1, True, files, 2)
                    word.Selection.Find.Execute("接单编号", False, False, False, False, False, True, 1, True, r, 2)
                    word.Selection.Find.Execute("日期：", False, False, False, False, False, True, 1, True, '日期：'+today_str, 2)

                    Review_doc.PrintOut()
                    word.Selection.Find.Execute(files, False, False, False, False, False, True, 1, True, '文件名', 2)
                    word.Selection.Find.Execute(r, False, False, False, False, False, True, 1, True,"接单编号", 2)
                    word.Selection.Find.Execute ('日期：'+today_str, False, False, False, False, False, True, 1, True, "日期：", 2)
                    Review_doc.Close()
                    Invoice_doc = word.Documents.Open(ne)
                    Invoice_doc.PrintOut()
                    Invoice_doc.Close()
                    word.Quit()
                    i=i+1
    s = str(i)
    win32api.MessageBox(0, '此次共打印'+ s +'票','统计' )
                    #printer_loading(ne)



default_dir = r"文件路径"
path = tkinter.filedialog.askdirectory(title=u'选择文件', initialdir=(os.path.expanduser((default_dir))))  # 全部文件的路径
# todir = tkinter.filedialog.askdirectory(title=u'选择文件',initialdir=(os.path.expanduser((default_dir))))
todir = "D:/个人文件/桌面/运单/"  # 存放复制文件的路径"C:/Users/Administrator/Desktop/运单/"
# folder_rename(path)
file_rename(path, todir)


def excelFilesPath(path):
    '''
    path: 目录文件夹地址
    返回值：列表，pdf文件全路径
    '''
    filePaths = []  # 存储目录下的所有文件名，含路径
    for root, dirs, files in os.walk(path):
        for file in files:
            filePaths.append(os.path.join(root, file))
    return filePaths
