# this is the first program
print("hello world")
a = 10
b = 11
print( a + b )

import io
import sys
import threading
import time
import pyautogui as pg
from pywinauto import application
import os
import datetime
import xlwings as xw
# import pandas as pd
import pymysql as pms


class Operate_Sap:
    # 提供全局参数
    def __init__(self, Sap_UserName, sap_password):
        self.Sap_UserName = Sap_UserName
        self.sap_password = sap_password

    def login_sap(self):
        app = application.Application()
        app.start(r"")  # exe图标地址
        time.sleep(3)
        pg.press("enter")
        time.sleep(2)
        pg.write(self.Sap_UserName, interval=0.10)
        pg.press("enter")
        time.sleep(1)
        pg.press("tab")
        time.sleep(1)
        pg.write(self.sap_password, interval=0.10)
        pg.press("enter")
        time.sleep(2)
        login_two = pg.locateOnScreen(r"D:\SAP_image\img1.png")
        pg.click(login_two)
        time.sleep(2)
        pg.press("enter")
        time.sleep(2)
        input_box = pg.locateOnScreen(r"D:\SAP_image\input_box.png")
        pg.click(input_box)
        time.sleep(1)
        pg.write("")
        time.sleep(1)
        pg.press("enter")
        time.sleep(1)
        enter_button = pg.locateOnScreen(r"D:\SAP_image\enter.png")
        pg.click(enter_button)
        time.sleep(2)
        continental = ["", "", "", "", "", ""]
        foreign_region = ["", "", ""]

        yesterday = (datetime.date.today() + datetime.timedelta(days=-1)).strftime("%Y%m%d")
        print(yesterday)
        # 所有主体都有订单数据需要导出，所以做个循环
        for index in range(len(continental)):
            pg.press("down")
            # 输入机构代码
            pg.write(continental[index])
            # 下移7个格子
            for index1 in range(7):
                pg.press("down")
            time.sleep(1)
            # 先清空数据
            pg.hotkey("ctrl", "a")
            pg.press("delete")
            # 输入日期
            pg.write(yesterday, interval=0.2)
            pg.press("enter")
            time.sleep(1)
            pg.press("tab")
            # 先清空数据
            pg.hotkey("ctrl", "a")
            pg.press("delete")
            # 输入时间
            pg.write(yesterday, interval=0.2)
            pg.press("enter")
            time.sleep(1)
            # 点击执行按钮
            execute_button = pg.locateOnScreen(r"D:\SAP_image\zhixing.png")
            pg.click(execute_button)
            if continental[index] == "":
                time.sleep(2400)
            else:
                time.sleep(600)
            # 使用热键清单->保存->文件
            pg.hotkey("alt", "l")
            time.sleep(1)
            pg.press("s")
            time.sleep(1)
            pg.press("f")
            time.sleep(1)
            # 向下移动一格，导出电子表格
            pg.press("down")
            # 按回车
            pg.press("enter")
            time.sleep(1)
            # 默认出现在“文件名称”处
            pg.hotkey("ctrl", "a")
            pg.press("Backspace")
            pg.write(continental[index] + ".xls")
            pg.press("enter")
            time.sleep(1)
            # 修改存储位置
            pg.moveTo(415, 334, duration=1)
            pg.click()
            time.sleep(1)
            pg.hotkey("ctrl", "a")
            pg.press("Backspace")
            pg.write(r"D:\SAP_Excel\continental")
            pg.press("enter")
            time.sleep(1)
            pg.press("enter")
            time.sleep(2)
            back_button = pg.locateOnScreen(r"D:\SAP_image\back.png")
            pg.click(back_button)
            time.sleep(1)

        for index_hw in range(len(foreign_region)):
            pg.press("down")
            # 输入机构代码
            pg.write(continental[index_hw])
            # 下移7个格子
            for index2 in range(7):
                pg.press("down")
            time.sleep(1)
            # 先清空数据
            pg.hotkey("ctrl", "a")
            pg.press("delete")
            # 输入日期
            pg.write(yesterday, interval=0.2)
            pg.press("enter")
            time.sleep(1)
            pg.press("tab")
            # 先清空数据
            pg.hotkey("ctrl", "a")
            pg.press("delete")
            # 输入时间
            pg.write(yesterday, interval=0.2)
            pg.press("enter")
            time.sleep(1)
            # 点击执行按钮
            execute_button = pg.locateOnScreen(r"D:\SAP_image\zhixing.png")
            pg.click(execute_button)
            time.sleep(600)
            # 使用热键清单->保存->文件
            pg.hotkey("alt", "l")
            time.sleep(1)
            pg.press("s")
            time.sleep(1)
            pg.press("f")
            time.sleep(1)
            # 向下移动一格，导出电子表格
            pg.press("down")
            # 按回车
            pg.press("enter")
            time.sleep(1)
            # 默认出现在“文件名称”处
            pg.hotkey("ctrl", "a")
            pg.press("Backspace")
            pg.write(continental[index] + ".xls")
            pg.press("enter")
            time.sleep(1)
            # 修改存储位置
            pg.moveTo(415, 334, duration=1)
            pg.click()
            time.sleep(1)
            pg.hotkey("ctrl", "a")
            pg.press("Backspace")
            pg.write(r"D:\SAP_Excel\foreign_region")
            pg.press("enter")
            time.sleep(1)
            pg.press("enter")
            time.sleep(2)
            # 返回上一页
            back_button = pg.locateOnScreen(r"D:\SAP_image\back.png")
            pg.click(back_button)
            time.sleep(1)
        time.sleep(2)
        # 关闭SAP
        close_app = pg.locateOnScreen(r"D:\SAP_image\img_close.png")
        pg.click(close_app)
        time.sleep(1)
        close_yes = pg.locateOnScreen(r"D:\SAP_image\yes.png")
        pg.click(close_yes)
        time.sleep(2)

    def Excel_Update(self):
        # 相关Excel的数据做操作
        # 做个数组，用于规定Excel名称
        continental = []
        foreign_region = []
        continental_num = []
        foreign_region_num = []

        # 大陆的表做操作
        for index_excel in range(len(continental)):
            # 读取文件
            app = xw.App(visible=True, add_book=False)
            # 打开新的工作簿
            # wb = app.books.add()
            # 打开想要打开的excel文件
            wb = app.books.open("D:\SAP_Excel\continental\{0}".format(continental[index_excel]))
            # 进入第一个报表页
            sheet = wb.sheets[0]
            # 获取所有行数
            rows = sheet.used_range.last_cell.row
            # 删除开始时间
            sheet['D:D'].delete()
            # 删除期初和期末中间的两列
            for index_del in range(2):
                sheet['F:F'].delete()
                time.sleep(0.1)
            # 删除后面6列数据
            for index_del1 in range(6):
                sheet['G:G'].delete()
                time.sleep(0.1)
            # 上下空白删除
            sheet['1:1'].delete()
            time.sleep(1)
            sheet['2:2'].delete()

            wb.sheets[continental_num[index_excel]].range('A1').value = "主体"
            # 输入数据
            wb.sheets[continental_num[index_excel]].range('G1').value = "流入方向"
            # li_nr = sheet.used_range.shape
            # 获取有数据的行数
            nrows = sheet.api.UsedRange.Rows.count
            # 循环遍历写入机构代码
            for index_row in range(nrows):
                if index_row == 0:
                    continue
                # 写入机构编码
                sheet.range(1 + index_row, 1).expand('down').value = continental_num[index_excel]
                time.sleep(0.1)

            # 循环遍历写入国内还是海外
            for index_row1 in range(nrows):
                if index_row1 == 0:
                    continue
                # 写入国内或者海外
                sheet.range(1 + index_row1, 7).expand('down').value = "国内"
                time.sleep(0.1)
            time.sleep(1)
            wb.save("D:\SAP_Excel\continental\{0}".format(continental[index_excel]))

        # 大陆以外的国家地区的表做操作
        for index_excel in range(len(foreign_region)):
            # 读取文件
            app = xw.App(visible=True, add_book=False)
            # 打开新的工作簿
            # wb = app.books.add()
            # 打开想要打开的excel文件
            wb = app.books.open(r"D:\SAP_Excel\foreign_region\{0}".format(foreign_region[index_excel]))
            # 进入第一个报表页
            sheet = wb.sheets[0]
            # 获取所有行数
            rows = sheet.used_range.last_cell.row
            # 删除开始时间
            sheet['D:D'].delete()
            # 删除期初和期末中间的两列
            for index_del in range(2):
                sheet['F:F'].delete()
                time.sleep(0.1)
            # 删除后面6列数据
            for index_del1 in range(6):
                sheet['G:G'].delete()
                time.sleep(0.1)
            # 上下空白删除
            sheet['1:1'].delete()
            time.sleep(1)
            sheet['2:2'].delete()
            wb.sheets[foreign_region_num[index_excel]].range('A1').value = "主体"
            # 输入数据
            wb.sheets[foreign_region_num[index_excel]].range('G1').value = "流入方向"
            # li_nr = sheet.used_range.shape
            nrows = sheet.api.UsedRange.Rows.count
            # 遍历写入机构编码
            for index_row in range(nrows):
                if index_row == 0:
                    continue
                sheet.range(1 + index_row, 1).expand('down').value = foreign_region_num[index_excel]
                time.sleep(0.1)

                # 循环遍历写入国内还是海外
            for index_row1 in range(nrows):
                if index_row1 == 0:
                    continue
                # 写入国内或者海外
                sheet.range(1 + index_row1, 7).expand('down').value = "国内"
                time.sleep(0.1)
            time.sleep(1)
            wb.save(r"D:\SAP_Excel\foreign_region\{0}".format(foreign_region[index_excel]))
            # 关掉，防止卡顿
        wb.close()
        app.kill()
        app.quit()
        time.sleep(2)

    def into_mysql(self):
        # 做个数组，用于规定Excel名称
        continental = []
        foreign_region = []
        continental_num = []
        foreign_region_num = []
        # 用于遍历大陆的Excel表格
        for index_continental in range(len(continental_num)):
            # 数据库操作
            connection = pms.Connection(
                host="",
                port=3306,
                user="",
                password="",
                db="",
                charset="utf8"
            )
            # 做个光标
            cur = connection.cursor()
            # 设置基础值
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(r"D:\SAP_Excel\continental\{0}".format(continental[index_continental]))
            sheet = wb.sheets[0]
            rows = sheet.used_range.last_cell.row
            # col = sheet.used_range.last_cell.column
            print(rows)
            for index_frame in range(rows):
                # 直接加数据
                index_frame += 2
                if index_frame > rows:
                    continue

                A_frame = str(int(sheet.range('A{0}:A{1}'.format(index_frame, index_frame)).value))
                B_frame = str(int(sheet.range('B{0}:B{1}'.format(index_frame, index_frame)).value))
                C_frame = str(int(sheet.range('C{0}:C{1}'.format(index_frame, index_frame)).value))
                E_frame = str(int(sheet.range('E{0}:E{1}'.format(index_frame, index_frame)).value))
                F_frame = str(int(sheet.range('F{0}:F{1}'.format(index_frame, index_frame)).value))
                D_frame = (datetime.date.today() + datetime.timedelta(days=-1)).strftime("%Y-%m-%d %H:%M:%S")
                G_frame = str(sheet.range('G{0}:G{1}'.format(index_frame, index_frame)).value)

                # 睡眠0.1秒
                time.sleep(0.1)

                # SQL语句
                sql = 'INSERT INTO dwd_sap_inventory_statistics(org_id, werks_no, material_id, opening_inventory, ending_inventory, count_date, inflow_direction) \
                       VALUE (' + A_frame + ',' + B_frame + ',' + C_frame + ',' + E_frame + ',' + F_frame + ',"' + D_frame + '","' + G_frame + '")'

                # 执行SQL
                cur.execute(sql)
                connection.commit()
                print(index_frame)

        for index_foreign in range(len(foreign_region_num)):

            # 数据库操作
            connection = pms.Connection(
                host="",
                port=3306,
                user="",
                password="",
                db="",
                charset="utf8"
            )
            # 做个光标
            cur = connection.cursor()
            # 设置基础值
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(r"D:\SAP_Excel\foreign_region\{0}".format(foreign_region[index_foreign]))
            sheet = wb.sheets[0]
            rows = sheet.used_range.last_cell.row
            # col = sheet.used_range.last_cell.column

            for index_frame in range(rows):
                # 判断，用于从第二行开始获取数据
                index_frame += 2
                # 没啥
                if index_frame > rows:
                    continue
                A_frame = str(int(sheet.range('A{0}:A{1}'.format(index_frame, index_frame)).value))
                B_frame = str(int(sheet.range('B{0}:B{1}'.format(index_frame, index_frame)).value))
                C_frame = str(int(sheet.range('C{0}:C{1}'.format(index_frame, index_frame)).value))
                E_frame = str(int(sheet.range('E{0}:E{1}'.format(index_frame, index_frame)).value))
                F_frame = str(int(sheet.range('F{0}:F{1}'.format(index_frame, index_frame)).value))
                D_frame = (datetime.date.today() + datetime.timedelta(days=-1)).strftime("%Y-%m-%d %H:%M:%S")
                G_frame = str(sheet.range('G{0}:G{1}'.format(index_frame, index_frame)).value)

                # 睡眠0.1秒
                time.sleep(0.1)
                
                time.sleep(0.1)


                # SQL语句
                sql = 'INSERT INTO dwd_sap_inventory_statistics(org_id, werks_no, material_id, opening_inventory, ending_inventory, count_date, inflow_direction) \
                                   VALUE (' + A_frame + ',' + B_frame + ',' + C_frame + ',' + E_frame + ',' + F_frame + ',"' + D_frame + '","' + G_frame + '")'

                # 执行SQL
                cur.execute(sql)
                # 提交语句
                connection.commit()
                print(index_frame)
        # 收尾
        connection.close()
        cur.close()
        wb.close()
        app.kill()
        app.quit()
        time.sleep(2)

    def delete_excel(self):
        # 做个数组，用于规定Excel名称
        continental = []
        foreign_region = []
        print("我要开始删除了")
        for index_delete in range(len(continental)):
            # 删除所有的文件，大陆主体
            os.remove(r"D:\SAP_Excel\continental\{0}".format(continental[index_delete]))
            time.sleep(0.1)
        for index_delete1 in range(len(foreign_region)):
            # 删除所有文件，海外主体
            os.remove(r"D:\SAP_Excel\foreign_region\{0}".format(foreign_region[index_delete1]))
            time.sleep(0.1)


Operate_Sap(Sap_UserName='', sap_password='').into_mysql()
