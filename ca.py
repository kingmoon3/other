#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Date    : 2021-09-11 16:52:57
# @Author  : Kingxujw (kingxujw@gmail.com)
# @Link    : https://www.github.com/kingmoon3/
# @Version : 1.0

import os
import pandas as pd
import sys
import re
import time
import datetime
from warnings import simplefilter
import xlsxwriter
import numpy as np


simplefilter(action='ignore', category=FutureWarning)


def run(filename):
    match_path = 'dy.xlsx'
    match_pd = pd.read_excel(
        match_path, sheet_name="Sheet1", engine='openpyxl')

    if 'bz' in filename:
        df = pd.read_excel(filename, sheet_name="Data", converters={
                           'ASN编号': str, '采购订单': str}, engine='openpyxl')
        sort = ["客户地点", "产品", "入库日期", "ASN编号",
                "采购订单", "收货数量", "净价", "价格单位", "入库数量"]
        sortno = []

        for i in sort:
            sortno.append(df.columns.get_loc(i))

        filter_merge = df.iloc[:, sortno]
        combine_pd = pd.merge(match_pd, filter_merge, how='outer', on="产品")
        match_customer = pd.read_excel(match_path, sheet_name="Sheet2")
        combine_customer = pd.merge(match_customer, combine_pd, on="客户地点")
        combine_customer["净价"] = combine_customer["净价"] / \
            combine_customer["价格单位"]
        combine_customer['收货数量'] = combine_customer['入库数量'].apply(
            np.sign) * combine_customer['收货数量'].abs()
        mask = combine_customer["净价"] < 0
        combine_customer.loc[mask, "收货数量"] *= -1
        combine_customer.loc[mask, "净价"] *= -1
        combine_customer.insert(
            9, "含税金额", combine_customer["收货数量"] * combine_customer["净价"] * 1.13)
        combine_customer.drop(columns=["入库数量"], inplace=True)
        combine_customer.rename(columns={'收货数量': '入库数量'}, inplace=True)

    elif 'js' in filename:
        df = pd.read_excel(filename, sheet_name="Data", converters={
                           'ASN编号': str, '采购订单': str}, engine='openpyxl')
        sort = ["客户", "物料号", "过账日期", "物料凭证", "CMPN", "基本数量", "开票单价", "价格单位"]
        sortno = []

        for i in sort:
            sortno.append(df.columns.get_loc(i))

        filter_merge = df.iloc[:, sortno]
        filter_merge.columns = ["客户地点", "产品",
                                "入库日期", "ASN编号", "采购订单", "入库数量", "净价", "价格单位"]
        combine_pd = pd.merge(match_pd, filter_merge, how='outer', on="产品")
        match_customer = pd.read_excel(
            match_path, sheet_name="Sheet2", engine='openpyxl')
        combine_customer = pd.merge(match_customer, combine_pd, on="客户地点")
        combine_customer["净价"] = combine_customer["净价"] / \
            combine_customer["价格单位"]
        mask = combine_customer["净价"] < 0
        combine_customer.loc[mask, "入库数量"] *= -1
        combine_customer.loc[mask, "净价"] *= -1
        combine_customer.insert(
            9, "含税金额", combine_customer["入库数量"] * combine_customer["净价"] * 1.13)

    else:
        pass

    if "hz" in filename:
        place = "湖州"
    elif "gd" in filename:
        place = "东莞"
    else:
        pass
    if "bz" in filename:
        datatype = "标准"
    elif "js" in filename:
        datatype = "寄售"
    else:
        pass

    try:
        combine_customer.insert(10, "发货工厂", place)
        combine_customer.insert(11, "订单类型", datatype)

        return combine_customer
    except:
        return


def file_name(file_dir):
    L = []
    for root, dirs, files in os.walk(file_dir):
        for file in files:
            if os.path.splitext(file)[1] == '.xlsx':
                L.append(os.path.join(file))
    return L


if __name__ == "__main__":
    run_path = os.path.split(os.path.abspath(sys.argv[0]))

    # 含子文件夹
    # allname = file_name(run_path[0])

    # 不含子文件夹
    allname = os.listdir()
    print('当前目录文件：', allname, '\n', '*' * 60)
    needname = []
    for name in allname:
        try:
            need = re.match('gd.*?xlsx|hz.*?xlsx', name).group()
            needname.append(need)
        except:
            pass
    print('待处理文件：', needname, '\n', '*' * 60)
    sort = ["客户地点", "客户名称", "产品", "型号", "入库日期",
            "ASN编号", "采购订单", "入库数量", "净价", "含税金额", "发货工厂", "订单类型"]
    dfall = pd.DataFrame(columns=sort)

    for name in needname:
        df = run(name)
        dfall = dfall._append(df, ignore_index=True)

    sortno = []

    for i in sort:
        sortno.append(dfall.columns.get_loc(i))

    dfall = dfall.iloc[:, sortno]
    # print(dfall)

    pivot = dfall.pivot_table(index='客户名称', values=[
                              '入库数量', '含税金额'], aggfunc='sum')
    pivot2 = dfall.pivot_table(index=['入库日期', '客户名称'], values=[
                               '入库数量', '含税金额'], aggfunc='sum')
    pivot3 = dfall.pivot_table(index='型号', values=['入库数量'], columns=[
                               '入库日期'], fill_value=0, aggfunc='sum')
    pivot4 = dfall.pivot_table(index=['客户名称', '型号'], values=['入库数量'], columns=[
        '入库日期'], fill_value=0, aggfunc='sum')
    # print(pivot3)
    pivot4.columns = pivot4.columns.droplevel(0)
    pivot5 = pivot4.rename_axis(None, axis=1).reset_index()
    # print(pivot5)
    td = datetime.date.today()
    od = datetime.timedelta(days=1)
    yd = td - od
    yesterday = str(yd)
    today = str(td)
    pivot2_yd = pivot2[yesterday:yesterday]
    pivot2_td = pivot2[today:today]
    pivot2_yd = pivot2_yd.reset_index()
    pivot2_td = pivot2_td.reset_index()
    pivot2_yd['入库日期'] = pivot2_yd['入库日期'].apply(
        lambda x: x.strftime('%Y-%m-%d'))
    pivot2_td['入库日期'] = pivot2_td['入库日期'].apply(
        lambda x: x.strftime('%Y-%m-%d'))
    # pivot2_yd.loc['合计']=pivot2_yd[['入库数量', '含税金额']].sum(axis=0)
    # pivot2_td.loc['合计']=pivot2_td[['入库数量', '含税金额']].sum(axis=0)

    pivot2_yd_sum = pivot2_yd[['入库数量', '含税金额']].sum()
    pivot2_yd_sum['客户名称'] = '合计'
    pivot2_yd_sum['入库日期'] = yesterday
    pivot2_yd = pivot2_yd._append(pivot2_yd_sum, ignore_index=True)
    pivot2_td_sum = pivot2_td[['入库数量', '含税金额']].sum()
    pivot2_td_sum['客户名称'] = '合计'
    pivot2_td_sum['入库日期'] = today
    pivot2_td = pivot2_td._append(pivot2_td_sum, ignore_index=True)

    print('输出透视表：', '\n', pivot, '\n', '*' * 60, '\n')
    # print('输出透视表：', '\n', pivot2, '\n', '*' * 60, '\n')
    for name in ["all-result.xlsx", "透视表.xlsx", "型号透视.xlsx"]:
        if os.path.exists(name):
            if os.path.exists(name.split('.')[0] + "_bak.xlsx"):
                try:
                    os.remove(name.split('.')[0] + "_bak.xlsx")
                    os.rename(name, name.split('.')[0] + "_bak.xlsx")
                except Exception as e:
                    print('错误：', str(e))
                    print('5秒后自动关闭窗口.....')
                    time.sleep(5)
                    sys.exit(0)

            else:
                try:
                    os.rename(name, name.split('.')[0] + "_bak.xlsx")
                except Exception as e:
                    print('错误：', str(e))
                    print('5秒后自动关闭窗口.....')
                    time.sleep(5)
                    sys.exit(0)

    try:
        print('正在输出结果到表格......')
        dfall.to_excel(excel_writer="all-result.xlsx", index=False)
        # pivot3.to_excel(excel_writer="型号透视.xlsx", index="型号")
        workbook = xlsxwriter.Workbook("型号透视.xlsx", options={
            'default_format_properties': {
                'font_name': '微软雅黑',
                'font_size': 9,
                'align': 'center',
                'valign': 'vcenter',
            },
        })
        worksheet = workbook.add_worksheet("Sheet1")
        workbook.add_format({})
        onlydate = pivot2.reset_index(inplace=False)[
            "入库日期"].drop_duplicates(keep='first', inplace=False)
        # pivot3_sum = pd.DataFrame(pivot3.sum()).T
        # print('pivot3_sum:', '\n', pivot3_sum)
        # pivot3_pivot_sum = pivot3.append(pivot3_sum)
        # print('pivot3_pivot_sum:', '\n', pivot3_pivot_sum)
        # pivot3_pivot_sum = pivot3_pivot_sum.rename(index={0: u'合计'})
        # print('pivot3_pivot_sum:', '\n', pivot3_pivot_sum)
        # pivot3_pivot_sum.to_excel('111.xlsx', encoding="utf-8")
        pivot3 = pd.DataFrame(pivot3)
        pivot3 = pivot3 / 1000
        pivot3.insert(0, "月合计", pivot3.apply(lambda x: x.sum(), axis=1))
        pivot3.insert(1, "日均", pivot3["月合计"] / len(onlydate))
        pivot4 = pd.DataFrame(pivot4)
        pivot4 = pivot4 / 1000
        pivot4.insert(0, "月合计", pivot4.apply(lambda x: x.sum(), axis=1))
        pivot4.insert(1, "日均", pivot4["月合计"] / len(onlydate))
        # print(pivot3)
        pivot3.to_excel('型号透视.xlsx')
        pivot4.to_excel('型号透视2.xlsx')
        # pivot3 = pivot3.reset_index()
        # # print(pivot3)
        # tuples = [tuple(x) for x in pivot3.values]
        # # print(tuples)
        # row = 1
        # col = 0

        # for item, count, amount in tuples:
        #     worksheet.write(row, col, item)
        #     worksheet.write(row, col + 1, count)
        #     worksheet.write(row, col + 2, amount)
        #     row += 1
        # # worksheet.write_formula(
        # #     'H2', '=IFERROR(VLOOKUP(INDIRECT("E"&ROW()),Sheet1!$A:$B,2,0)/1000,0)')
        # workbook.close()
        # # =IFERROR(VLOOKUP(INDIRECT("E"&ROW()),Sheet1!$A:$B,2,0)/1000,0)
        with pd.ExcelWriter("透视表.xlsx", engine='openpyxl') as writer:
            pivot.to_excel(excel_writer=writer, index='客户名称')
            # pivot2.to_excel(excel_writer=writer, index=['入库日期', '客户名称'], startcol=5)
            if pivot2_yd.empty != 1:
                pivot2_yd.to_excel(excel_writer=writer,
                                   index=False, startcol=4)
            if pivot2_td.empty != 1:
                pivot2_td.to_excel(excel_writer=writer,
                                   index=False, startcol=9)

        print('输出成功！')
        print('5秒后自动关闭窗口.....')
        time.sleep(5)
        sys.exit(0)
    except Exception as e:
        print('错误：', str(e))
        print('5秒后自动关闭窗口.....')
        time.sleep(5)
        sys.exit(0)
