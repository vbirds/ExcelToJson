#!/usr/bin/python
# -*- coding: UTF-8 -*-

import os
import threading
import xlrd
from Tkinter import *
from FileDialog import *
import tkMessageBox



class ExcelToJson(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.pack()
        self.createWidgets()

    def createWidgets(self):
        """
        初始化窗体
        :return:
        """
        self.singleLabel = Label(self, text="单个转换：")
        self.batchLabel  = Label(self, text="批量转换：")
        self.singleConvertButton = Button(self, text="选择单文件", command=self.singleConvert)
        self.batchConvertButton = Button(self, text="选择文件夹", command=self.batchConvert)

        self.singleLabel.grid(row=1, column=0)
        self.singleConvertButton.grid(row=1, column=1)
        self.batchLabel.grid(row=2, column=0)
        self.batchConvertButton.grid(row=2, column=1)

    def singleConvert(self):
        """
        转换单个文件为json
        """
        fd = LoadFileDialog(self)
        filename = fd.go()

        if filename:
            self.doConvertBase(filename)
            tkMessageBox.showinfo("Excel To Json", "转换成功")

    def batchConvert(self):
        """
        批量转换文件为json，自动获取选择文件夹下的xls文件，转为json
        """
        fd = FileDialog(self)
        dir = fd.go()
        filenames = self.getFilesFromDir(dir, '.xls')
        # print(filenames)
        # 获取开启的线程数
        threadnum = self.getThreadNum(len(filenames))
        # 根据线程数对原有 filenames列表进行拆分，分配给不同线程
        threadlist = self.splitList(filenames, threadnum)

        for list in threadlist:
            try:
                t1 = threading.Thread(target=ExcelToJson.doConvert, args=(self, list))
                t1.start()
                t1.join()
                #threads.append(t1)
            except:
                print("Error: unable create thread")

        tkMessageBox.showinfo("Excel To Json", "转换成功")


    def getFilesFromDir(self, dir, filter):
        """
        获取当前文件夹下指定包含filter字符串的文件列表
        :param dir: 文件夹路径
        :param filter: 过滤字符串
        :return: 文件路径列表
        """
        filenames = []
        list = os.listdir(dir)

        for file in list:
            filepath = os.path.join(dir, file)
            if os.path.isdir(filepath):
                continue
            if filepath.find(filter) == -1:
                continue
            filenames.append(filepath)

        return filenames


    def getThreadNum(self, filenum):
        """
        计算需要开启的线程数
        :param filenum: 文件列表长度
        :return: 线程数
        """
        if filenum <= 2:
            return 1
        threadnum = (filenum / 3) + 1
        if threadnum > 5:
            return 5
        return threadnum

    def splitList(self, filelist, num):
        """
        根据线程数划分每个线程执行的文件路径列表
        :param filelist: 文件路径列表
        :param num: 线程数
        :return: 线程执行的列表  : [[file1, file2, file3], [file4, file5, file6]]
        """
        threadlist = []
        listnum = []
        remaindernum = len(filelist) - 3 * (num - 1)
        for i in range(1, num):
            listnum.append([(i - 1) * 3, 3 * i])

        for list in listnum:
            threadlist.append(filelist[list[0]:list[1]])

        threadlist.append(filelist[(0 - remaindernum):])
        return threadlist;


    def doConvert(self, filelist):
        """
        转换函数，调用核心转换函数doConvertBase
        :param filelist: 文件列表
        :return:
        """
        for file in filelist:
            self.doConvertBase(file)

    def doConvertBase(self, filename):
        """
        核心转换函数
        :param filename: 文件路径
        :return:
        """

        excel_file = xlrd.open_workbook(filename)
        (filename, exten) = os.path.splitext(filename)
        outputfile = filename + '.json'
        output = open(outputfile, 'w+')

        table = excel_file.sheet_by_index(0)
        nrows = table.nrows
        ncols = table.ncols
        title_table = table.row_values(0)

        # 写开头格式
        output.writelines('[\n')
        # 写json对象
        for i in range(1, nrows):
            output.writelines('  {\n')

            for j in range(ncols):
                temp = ''
                value = table.row(i)[j].value

                if  isinstance(value, float) or isinstance(value, int):
                    temp = "    \"%s\":%d,\n" % (title_table[j], value)
                elif isinstance(value, unicode):
                    temp = "    \"%s\":\"%s\",\n" % (title_table[j], value.encode('utf-8'))
                else:
                    temp = "    \"%s\":\"%s\",\n" % (title_table[j], value)

                output.writelines(temp)

            if i == (nrows - 1):
                output.writelines('  }\n')
            else:
                output.writelines('  },\n')

        # 写结尾']'
        output.writelines(']\n')
        output.close()


app = ExcelToJson()
# 设置窗口标题:
app.master.title('Excel To Json')
# 主消息循环:
app.mainloop()

