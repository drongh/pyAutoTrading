# -*- encoding: utf8 -*-
# QQ群： 486224275
__author__ = '人在江湖'

import tkinter.messagebox
from tkinter import *
from tkinter.ttk import *
import datetime
import threading
import pickle
import time

import tushare as ts

from winguiauto import (dumpWindow, dumpWindows, getWindowText,
                        getListViewInfo, setEditText, clickWindow,
                        click, closePopupWindows, findTopWindow,
                        restoreFocusWindow)

is_start = False
is_monitor = True
set_stock_info = []
actual_stock_info = []
consignation_info = []
is_ordered = [1] * 5  # 1：未下单  0：已下单


class Operation:
    def __init__(self, top_hwnd):

        self.__top_hwnd = top_hwnd
        self.__button = {'refresh': 180, 'position': 145, 'deal': 112, 'withdrawal': 83, 'sell': 50, 'buy': 20}
        windows = dumpWindows(self.__top_hwnd)
        temp_hwnd = 0
        for window in windows:
            child_hwnd, window_text, window_class = window
            if window_class == 'AfxMDIFrame42':
                temp_hwnd = child_hwnd
                break
        temp_hwnds = dumpWindow(temp_hwnd)
        temp_hwnds = dumpWindow(temp_hwnds[1][0])
        self.__menu_hwnds = dumpWindow(temp_hwnds[0][0])
        self.__buy_sell_hwnds = dumpWindow(temp_hwnds[4][0])

    def __buy(self, code, quantity):
        """
        买入函数
        :param code: 股票代码，字符串
        :param quantity: 数量， 字符串
        """
        setEditText(self.__buy_sell_hwnds[0][0], code)
        time.sleep(0.3)
        if quantity != '0':
            setEditText(self.__buy_sell_hwnds[3][0], quantity)
            time.sleep(0.3)
        click(self.__buy_sell_hwnds[5][0])
        time.sleep(0.3)

    def __sell(self, code, quantity):
        """
        卖出函数
        :param code: 股票代码， 字符串
        :param quantity: 数量， 字符串
        """
        setEditText(self.__buy_sell_hwnds[24][0], code)
        time.sleep(0.3)
        if quantity != '0':
            setEditText(self.__buy_sell_hwnds[27][0], quantity)
            time.sleep(0.3)
        click(self.__buy_sell_hwnds[29][0])
        time.sleep(0.3)

    def order(self, code, direction, quantity):
        """
        下单函数
        :param code: 股票代码， 字符串
        :param direction: 买卖方向
        :param quantity: 数量， 字符串，数量为‘0’时，就交易软件指定数量
        """
        restoreFocusWindow(self.__top_hwnd)
        self.clickRefreshButton()
        if direction == 'B':
            self.__buy(code, quantity)
        if direction == 'S':
            self.__sell(code, quantity)
        closePopupWindows(self.__top_hwnd)

    def clickRefreshButton(self):
        """
        点击刷新按钮
        """
        restoreFocusWindow(self.__top_hwnd)
        clickWindow(self.__menu_hwnds[0][0], self.__button['refresh'])

    def getMoney(self):
        """
        :return:可用资金
        """
        restoreFocusWindow(self.__top_hwnd)
        self.clickRefreshButton()
        setEditText(self.__buy_sell_hwnds[24][0], '999999')  # 测试时获得资金情况
        time.sleep(0.3)
        money = getWindowText(self.__buy_sell_hwnds[12][0]).strip()
        return float(money)

    def getPosition(self):
        """获取持仓股票信息
        """
        # self.clickRefreshButton()
        restoreFocusWindow(self.__top_hwnd)
        return getListViewInfo(self.__buy_sell_hwnds[64][0], 5)


def pickCodeFromItems(items_info):
    """
    提取股票代码
    :param items_info: UI下各项输入信息
    :return:股票代码列表
    """
    stock_codes = []
    for item in items_info:
        stock_codes.append(item[0])
    return stock_codes


def getStockData(items_info):
    """
    获取股票实时数据
    :param items_info:UI下各项输入信息
    :return:股票实时数据
    """
    code_name_price = []
    stock_codes = pickCodeFromItems(items_info)
    try:
        df = ts.get_realtime_quotes(stock_codes)
        df_len = len(df)
        for stock_code in stock_codes:
            is_found = False
            for i in range(df_len):
                actual_code = df['code'][i]
                if stock_code == actual_code:
                    code_name_price.append((actual_code, df['name'][i], df['price'][i]))
                    is_found = True
                    break
            if is_found is False:
                code_name_price.append(('', '', ''))
    except:
        code_name_price = [('', '', '')] * 5  # 网络不行，返回空
    return code_name_price


def monitor():
    """
    实时监控函数
    :return:
    """
    global actual_stock_info, consignation_info, is_ordered, set_stock_info
    count = 1

    top_hwnd = findTopWindow(wantedClass='TdxW_MainFrame_Class')
    if top_hwnd == 0:
        tkinter.messagebox.showerror('错误', '请先打开通达信交易软件，再运行本软件')
    else:
        operation = Operation(top_hwnd)

    while is_monitor and top_hwnd:

        if is_start:
            actual_stock_info = getStockData(set_stock_info)
            for row, (actual_code, actual_name, actual_price) in enumerate(actual_stock_info):
                if is_start and actual_code and is_ordered[row] == 1 \
                        and set_stock_info[row][1] and set_stock_info[row][2] > 0 \
                        and set_stock_info[row][3] and set_stock_info[row][4] \
                        and datetime.datetime.now().time() > set_stock_info[row][5]:
                    if set_stock_info[row][1] == '>' and float(actual_price) > set_stock_info[row][2]:
                        operation.order(actual_code, set_stock_info[row][3], set_stock_info[row][4])
                        dt = datetime.datetime.now()
                        consignation_info.append(
                            (dt.strftime('%x'), dt.strftime('%X'), actual_code,
                             actual_name, set_stock_info[row][3],
                             actual_price, set_stock_info[row][4], '已委托'))
                        is_ordered[row] = 0

                    if set_stock_info[row][1] == '<' and float(actual_price) < set_stock_info[row][2]:
                        operation.order(actual_code, set_stock_info[row][3], set_stock_info[row][4])
                        dt = datetime.datetime.now()
                        consignation_info.append(
                            (dt.strftime('%x'), dt.strftime('%X'), actual_code,
                             actual_name, set_stock_info[row][3],
                             actual_price, set_stock_info[row][4], '已委托'))
                        is_ordered[row] = 0

        if count % 200 == 0:
            operation.clickRefreshButton()
        time.sleep(3)
        count += 1


class StockGui:
    def __init__(self):
        self.window = Tk()
        self.window.title("自动化股票交易")
        self.window.resizable(0, 0)

        frame1 = Frame(self.window)
        frame1.pack(padx=10, pady=10)

        Label(frame1, text="股票代码", width=8, justify=CENTER).grid(
            row=1, column=1, padx=5, pady=5)
        Label(frame1, text="股票名称", width=8, justify=CENTER).grid(
            row=1, column=2, padx=5, pady=5)
        Label(frame1, text="实时价格", width=8, justify=CENTER).grid(
            row=1, column=3, padx=5, pady=5)
        Label(frame1, text="关系", width=4, justify=CENTER).grid(
            row=1, column=4, padx=5, pady=5)
        Label(frame1, text="设定价格", width=8, justify=CENTER).grid(
            row=1, column=5, padx=5, pady=5)
        Label(frame1, text="方向", width=4, justify=CENTER).grid(
            row=1, column=6, padx=5, pady=5)
        Label(frame1, text="数量", width=8, justify=CENTER).grid(
            row=1, column=7, padx=5, pady=5)
        Label(frame1, text="时间可选", width=8, justify=CENTER).grid(
            row=1, column=8, padx=5, pady=5)
        Label(frame1, text="状态", width=6, justify=CENTER).grid(
            row=1, column=9, padx=5, pady=5)

        self.rows = 5
        self.cols = 9

        self.variable = []
        for row in range(self.rows):
            self.variable.append([])
            for col in range(self.cols):
                temp = StringVar()
                self.variable[row].append(temp)

        for row in range(self.rows):
            Entry(frame1, textvariable=self.variable[row][0],
                  width=8).grid(row=row + 2, column=1, padx=5, pady=5)
            Entry(frame1, textvariable=self.variable[row][1], state=DISABLED,
                  width=8).grid(row=row + 2, column=2, padx=5, pady=5)
            Entry(frame1, textvariable=self.variable[row][2], state=DISABLED, justify=RIGHT,
                  width=8).grid(row=row + 2, column=3, padx=5, pady=5)
            Combobox(frame1, values=('<', '>'), textvariable=self.variable[row][3],
                     width=2).grid(row=row + 2, column=4, padx=5, pady=5)
            Spinbox(frame1, from_=0, to=1000, textvariable=self.variable[row][4],
                    increment=0.01, width=6).grid(row=row + 2, column=5, padx=5, pady=5)
            Combobox(frame1, values=('B', 'S'), textvariable=self.variable[row][5],
                     width=2).grid(row=row + 2, column=6, padx=5, pady=5)
            Spinbox(frame1, from_=0, to=100000, textvariable=self.variable[row][6],
                    increment=100, width=6).grid(row=row + 2, column=7, padx=5, pady=5)
            Entry(frame1, textvariable=self.variable[row][7],
                  width=8).grid(row=row + 2, column=8, padx=5, pady=5)
            Entry(frame1, textvariable=self.variable[row][8], state=DISABLED, justify=CENTER,
                  width=6).grid(row=row + 2, column=9, padx=5, pady=5)

        frame3 = Frame(self.window)
        frame3.pack(padx=10, pady=10)
        self.start_bt = Button(frame3, text="开始", command=self.start)
        self.start_bt.pack(side=LEFT)
        self.set_bt = Button(frame3, text='重置买卖', command=self.setFlags)
        self.set_bt.pack(side=LEFT)
        Button(frame3, text="历史记录", command=self.displayHisRecords).pack(side=LEFT)
        Button(frame3, text='保存', command=self.save).pack(side=LEFT)
        self.load_bt = Button(frame3, text='载入', command=self.load)
        self.load_bt.pack(side=LEFT)

        self.window.protocol(name="WM_DELETE_WINDOW", func=self.close)
        self.window.after(100, self.updateControls)
        self.window.mainloop()

    def displayHisRecords(self):
        """
        显示历史信息
        :return:
        """
        global consignation_info
        tp = Toplevel()
        tp.title('历史记录')
        tp.resizable(0, 1)
        scrollbar = Scrollbar(tp)
        scrollbar.pack(side=RIGHT, fill=Y)
        col_name = ['日期', '时间', '证券代码', '证券名称', '方向', '价格', '数量', '备注']
        tree = Treeview(
            tp, show='headings', columns=col_name, height=30, yscrollcommand=scrollbar.set)
        tree.pack(expand=1, fill=Y)
        scrollbar.config(command=tree.yview)
        for name in col_name:
            tree.heading(name, text=name)
            tree.column(name, width=70, anchor=CENTER)

        for msg in consignation_info:
            tree.insert('', 0, values=msg)

    def save(self):
        """
        保存设置
        :return:
        """
        global set_stock_info, consignation_info
        self.getItems()
        with open('stockInfo.dat', 'wb') as fp:
            pickle.dump(set_stock_info, fp)
            pickle.dump(consignation_info, fp)

    def load(self):
        """
        载入设置
        :return:
        """
        global set_stock_info, consignation_info
        try:
            with open('stockInfo.dat', 'rb') as fp:
                set_stock_info = pickle.load(fp)
                consignation_info = pickle.load(fp)
        except FileNotFoundError as error:
            tkinter.messagebox.showerror('错误', error)

        for row in range(self.rows):
            for col in range(self.cols):
                if col == 0:
                    self.variable[row][col].set(set_stock_info[row][0])
                elif col == 3:
                    self.variable[row][col].set(set_stock_info[row][1])
                elif col == 4:
                    self.variable[row][col].set(set_stock_info[row][2])
                elif col == 5:
                    self.variable[row][col].set(set_stock_info[row][3])
                elif col == 6:
                    self.variable[row][col].set(set_stock_info[row][4])
                elif col == 7:
                    temp = set_stock_info[row][5].strftime('%X')
                    if temp == '01:00:00':
                        self.variable[row][col].set('')
                    else:
                        self.variable[row][col].set(temp)

    def setFlags(self):
        """
        重置买卖标志
        :return:
        """
        global is_start, is_ordered
        if is_start is False:
            is_ordered = [1] * 5

    def updateControls(self):
        """
        实时股票名称、价格、状态信息
        :return:
        """
        global actual_stock_info, is_start
        if is_start:
            for row, (actual_code, actual_name, actual_price) in enumerate(actual_stock_info):
                if actual_code:
                    self.variable[row][1].set(actual_name)
                    self.variable[row][2].set(actual_price)
                    if is_ordered[row] == 1:
                        self.variable[row][8].set('监控中')
                    elif is_ordered[row] == 0:
                        self.variable[row][8].set('已委托')
                else:
                    self.variable[row][1].set('')
                    self.variable[row][2].set('')
                    self.variable[row][8].set('')

        self.window.after(3000, self.updateControls)

    def start(self):
        """
        启动停止
        :return:
        """
        global is_start
        if is_start is False:
            is_start = True
        else:
            is_start = False

        if is_start:
            self.getItems()
            self.start_bt['text'] = '停止'
            self.set_bt['state'] = DISABLED
            self.load_bt['state'] = DISABLED
        else:
            self.start_bt['text'] = '开始'
            self.set_bt['state'] = NORMAL
            self.load_bt['state'] = NORMAL

    def close(self):
        """
        关闭程序时，停止monitor线程
        :return:
        """
        global is_monitor
        is_monitor = False
        self.window.quit()

    def getItems(self):
        """
        获取UI上用户输入的各项数据，
        """
        global set_stock_info
        set_stock_info = []

        # 获取买卖价格数量输入项等
        for row in range(self.rows):
            set_stock_info.append([])
            for col in range(self.cols):
                temp = self.variable[row][col].get().strip()
                if col == 0:
                    if len(temp) == 6 and temp.isdigit():  # 判断股票代码是否为6位数
                        set_stock_info[row].append(temp)
                    else:
                        set_stock_info[row].append('')
                elif col == 3:
                    if temp in ('>', '<'):
                        set_stock_info[row].append(temp)
                    else:
                        set_stock_info[row].append('')
                elif col == 4:
                    try:
                        price = float(temp)
                        if price > 0:
                            set_stock_info[row].append(price)  # 把价格转为数字
                        else:
                            set_stock_info[row].append(0)
                    except ValueError:
                        set_stock_info[row].append(0)
                elif col == 5:
                    if temp in ('B', 'S'):
                        set_stock_info[row].append(temp)
                    else:
                        set_stock_info[row].append('')
                elif col == 6:
                    if temp.isdigit() and int(temp) >= 0:
                        set_stock_info[row].append(str(int(temp) // 100 * 100))
                    else:
                        set_stock_info[row].append('')
                elif col == 7:
                    try:
                        set_stock_info[row].append(datetime.datetime.strptime(temp, '%H:%M:%S').time())
                    except ValueError:
                        set_stock_info[row].append(datetime.datetime.strptime('1:00:00', '%H:%M:%S').time())


if __name__ == '__main__':
    t1 = threading.Thread(target=StockGui)
    t1.start()
    t1.join(2)
    t2 = threading.Thread(target=monitor)
    t2.start()
