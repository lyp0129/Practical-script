# coding:utf-8

import xlrd
from xlutils.copy import copy

'''
这个程序的目的是把excle里的部分元素取出来做计算

思路就是，先写好几个方法操作excel
然后遍历一下excle把想要的元素取出来，并拿到他们的下标
把这个下标对应的其他列的值再计算一下

把正常的下标拿到，对应到支出收入，把支出收入的再分开，分别做和



'''


class Calculate:

    def __init__(self, excel_path=None, sheet_id=None):  # sheet_id是excel中的第几张表的下标
        if excel_path:
            self.excel_path = excel_path
            self.sheet_id = sheet_id
        else:
            self.excel_path = "/Users/lyp/Desktop/xxx.xlsx"
            self.sheet_id = 0

    # 获取excel

    def get_data(self):

        data = xlrd.open_workbook(self.excel_path)

        tables = data.sheet_by_index(self.sheet_id)

        return tables

    # 获取excel的行数

    def get_lines(self):
        line = self.get_data().nrows
        return line

    # 行内容
    def get_content_row(self, row):
        content_row = self.get_data().row_values(row)
        return content_row

    # 列内容
    def get_content_col(self, col):
        content_col = self.get_data().col_values(col)
        return content_col

    # 获取excle的内容
    # row 行数，col列

    def get_content(self, row, col):
        content = self.get_data().cell_value(row, col)
        return content

    def write_value(self, row, col, value):
        read_data = xlrd.open_workbook(self.excel_path)
        write_data = copy(read_data)
        sheet_data = write_data.get_sheet(0)
        sheet_data.write(row, col, value)
        write_data.save(self.excel_path)

    def Calculation(self):

        # 获取行
        line = self.get_lines()

        # 遍历一下取到excle中第2列所有的值并且生成一个content的列表

        for i in range(5, line):
            list_a = []
            content = self.get_content_col(i)  # 拿到了一个列表

            # 遍历一下content把"正常"的元素下标拿到

            for c in range(len(content)):
                if content[c] == "正常":
                    list_a.append(c)  # 拿到正常的下标

            return list_a

    def test1(self):

        # 获取行
        line = self.get_lines()

        for i in range(1, line):
            content = self.get_content_col(i)

            return content  # 拿到支出收入的列表

    def test2(self):
        list_b = []
        list_c = []
        i_want = self.Calculation()  # 正常的下标
        two = self.test1()  # 支出收入的列表

        '''
        通过正常的下标去循环遍历，取到了支出和收入的字符串
        然后判断如果包含，把下标放到一个列表
        不包含再放一个列表里，然后都返回出去，返回出去的是个元祖
        '''
        for i in i_want:
            if "支出" in two[i]:
                list_b.append(i)

            else:
                list_c.append(i)

        return list_b, list_c

    def test3(self):
        list4 = []
        yuanzu = self.test2()
        money = self.test4()
        for i in yuanzu[0]:
            list4.append(money[i])
        a = sum(list4)
        print("支出金额：" + str(sum(list4)))

        return a

    def test4(self):

        # 获取行
        line = self.get_lines()

        for i in range(4, line):
            content_money = self.get_content_col(i)

            return content_money  # 拿到支出收入的列表

    def test5(self):
        list5 = []
        yuanzu = self.test2()
        money = self.test4()
        for i in yuanzu[1]:
            list5.append(money[i])

        b = sum(list5)
        print("收入金额：" + str(sum(list5)))

        return b

    def xxx(self):
        a = self.test3()
        b = self.test5()
        c = b - a
        print("余额为" + (str(c)))


if __name__ == '__main__':
    oda = Calculate()
    a = oda.get_data()
    line = oda.get_lines()
    aa = oda.Calculation()
    # cc = oda.test1()
    # dd = oda.test3()
    # ee = oda.test5()
    ff = oda.xxx()
    # money=oda.test4()
    # print(money)

    # print(type(dd))
    # print(aa)
    # print(cc)
