#!/bin/python3
import datetime
import os
import platform
import re
from urllib.parse import urlparse

import xlrd

import util
from openpyxl import load_workbook, Workbook
import docx


class Spider:
    conf = ""
    src_file_name = ""
    data = []
    data_res = []

    def __init__(self, conf_fp):
        self.conf = util.read_conf(conf_fp)
        self.conf["g_src_dir"] = os.path.join(self.conf["data_dir"], self.conf["src_dir"])
        self.conf["g_dst_dir"] = os.path.join(self.conf["data_dir"], self.conf["dst_dir"])
        self.src_file_name = util.select_file(self.conf["g_src_dir"])
        self.abs_src_file_name = os.path.join(self.conf["g_src_dir"], self.src_file_name)
        self.conf["f_dst_dir"] = os.path.join(self.conf["g_dst_dir"], os.path.splitext(self.src_file_name)[0])
        self.conf["dst_file"] = os.path.join(self.conf["f_dst_dir"],
                                             "核_" + os.path.splitext(self.src_file_name)[0] + ".xlsx")
        if not os.path.exists(self.conf["f_dst_dir"]):
            os.mkdir(self.conf["f_dst_dir"])
        if self.conf["SCREEN"]:
            self.conf["screenshot_dir"] = os.path.join(self.conf["f_dst_dir"], "screenshot")
            if not os.path.exists(self.conf["screenshot_dir"]):
                os.mkdir(self.conf["screenshot_dir"])
        else:
            self.conf["screenshot_dir"] = None
        if self.conf["SCREEN"]:
            self.conf["pagesource_dir"] = os.path.join(self.conf["f_dst_dir"], "pagesource")
            if not os.path.exists(self.conf["pagesource_dir"]):
                os.mkdir(self.conf["pagesource_dir"])
        else:
            self.conf["pagesource_dir"] = None

    def init_data(self):
        import shutil
        if self.data.__len__() > 0:
            return
        if os.path.splitext(self.src_file_name)[1] in ['.txt']:
            file_object = open(os.path.join(self.conf["g_src_dir"], self.src_file_name), 'r', encoding='utf-8')
            lines = file_object.readlines()
            file_object.close()
            for line in lines:
                new_line = line.replace('\n', '')
                match_obj = re.search(
                    r"([hH][tT]{2}[pP]://|[hH][tT]{2}[pP][sS]://|[wW]{3}.|[wW][aA][pP].|[fF][tT][pP].|[fF][iI][lL][eE].)[-A-Za-z0-9+&@#/%?=~_|!:,.;]+[-A-Za-z0-9+&@#/%=~_|]",
                    new_line)
                if match_obj:
                    row_t = [self.data.__len__() + 1, self.data.__len__() + 1, match_obj.group()]
                    self.data.append(row_t)
        elif os.path.splitext(self.src_file_name)[1] in ['.xlsx']:
            wb = load_workbook(self.abs_src_file_name)
            ws = wb.worksheets[0]
            for r in range(1, ws.max_row + 1):
                row_t = [r]
                for i in range(1, 3):
                    row_t.append(ws.cell(row=r, column=i).value)
                self.data.append(row_t)
        elif os.path.splitext(self.src_file_name)[1] in ['.xls']:
            book = xlrd.open_workbook(self.abs_src_file_name)
            ws = book.sheet_by_index(0)
            for r in range(0, ws.nrows):
                row_t = [r]
                for i in range(0, 2):
                    row_t.append(ws.cell(r, i).value)
                self.data.append(row_t)
        elif os.path.splitext(self.src_file_name)[1] in ['.docx']:
            d = docx.opendocx(self.abs_src_file_name)
            doc = docx.getdocumenttext(d)
            for line in doc:
                new_line = line.replace('\n', '')
                match_obj = re.search(r"https?://(?:[-\w.]|(?:%[\da-fA-F]{2}))+", new_line)
                if match_obj:
                    row_t = [self.data.__len__() + 1, self.data.__len__() + 1, match_obj.group()]
                    self.data.append(row_t)
        else:
            print("输入文件错误，系统不支持...")

        for i in self.data:
            (r_code, f_url) = util.get_url_normalize(i[2], short_urls=self.conf["short_domain"])  # formal URL
            if r_code == 404:
                (r_code, f_url) = util.get_url_normalize(i[2], short_urls=self.conf["short_domain"])
            i.append(f_url)
            if f_url == "异常":
                i.append("异常")
            else:
                i.append(urlparse((i[3])).hostname)
            len_t = 11
            if i[3] == '异常':
                for ii in range(0, len_t):
                    i.append('异常')
            else:
                for ii in range(0, len_t):
                    i.append('')
            i[14] = r_code
        if os.path.splitext(self.src_file_name)[1] in ['.xlsx']:
            if not os.path.exists(self.conf["dst_file"]):
                shutil.copyfile(self.abs_src_file_name, self.conf["dst_file"])
                wb = load_workbook(self.conf["dst_file"])
                wb.save(self.conf["dst_file"])
        else:
            wb = Workbook()
            if os.path.splitext(self.src_file_name)[1] in ['.xlsx']:
                wb_s = load_workbook(self.abs_src_file_name)
                for sheet in wb_s.sheetnames:
                    ws_s = wb_s[sheet]
                    ws = wb[sheet]
                    for i, row in enumerate(ws_s.iter_rows()):
                        for j, cell in enumerate(row):
                            ws.cell(row=i + 1, column=j + 1, value=cell.value)
            wb.save(self.conf["dst_file"])
        print("*****获取网址：%d 条" % self.data.__len__())




def main():
    # git rm -r --cached .
    # git add .
    # git commit -m "update"
    # [2, 2, 'www.baidu.com', 'http://www.baidu.com', 'www.baidu.com', ['39.156.66.14', '39.156.66.18'], '境内',
    # '中国·北京', '百度一下，你就知道', 'https://www.baidu.com/', '是', 'https://www.baidu.com', '', '是', '', 1]
    # pyinstaller -F main.py
    start = datetime.datetime.now()
    if platform.system() == "Windows":
        os.system("ipconfig/flushdns")
    # 初始化待测试数据
    spd = Spider("./conf/config.json")
    spd.init_data()
    end = datetime.datetime.now()
    print("====================\n任务结束，耗时 %d 秒, \n拨测结果：%d条" % ((end - start).seconds, len(spd.data)))


if __name__ == "__main__":
    main()
