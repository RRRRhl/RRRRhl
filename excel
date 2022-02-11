import openpyxl
from prettytable import *

def excel(ip='114.114.114.114'):
    #oracle环境
    workbook = openpyxl.load_workbook(r'/Users/rhl/other/134.96.253.56/134.96.263.56/svn4dba/文档资料/环境信息/TA_CTZJ_PROD_DB_Environment_Spec_v1.0.xlsx')
    sheet1 = workbook['Prod']
    env='ORRACLE_PROD'
    max_num_oracle = sheet1.max_row
    for i in range(max_num_oracle):
        ip_1 = sheet1.cell(row=i + 2, column=5).value
        ip_2 = sheet1.cell(row=i + 2, column=4).value
        if ip == ip_1 or ip == ip_2:
            #取出匹配的字段
            Systemname=str(sheet1.cell(row=i+2,column=2).value)
            Host=str(sheet1.cell(row=i+2,column=3).value)
            IP=str(sheet1.cell(row=i+2,column=4).value)
            Virtual_IP=str(sheet1.cell(row=i+2,column=5).value)
            SID=str(sheet1.cell(row=i+2,column=6).value)
            OSPwd=str(sheet1.cell(row=i+2,column=8).value)
            DBPwd=str(sheet1.cell(row=i+2,column=9).value)
            User=str(sheet1.cell(row=i+2,column=20).value)

            #把取到的数，插入PrettyTable，格式化输出
            table=PrettyTable(['Environment','Systemname','Host','IP','Virtual_IP','SID','OSPwd','DBPwd','User'])
            table.add_row([env,Systemname,Host,IP,Virtual_IP,SID,OSPwd,DBPwd,User])
            print(table)

    #bss crm环境
    workbook = openpyxl.load_workbook(r'/Users/rhl/other/134.96.253.56/134.96.263.56/svn4dba/文档资料/环境信息/生产环境服务器列表.xlsx')
    sheet_crm = workbook['BSS_CRM_PROD']
    env = 'BSS_CRM_PROD'
    max_num_crm = sheet_crm.max_row
    for i in range(max_num_crm):
        ip_1 = sheet_crm.cell(row=i + 2, column=5).value
        vip_1 = str(sheet_crm.cell(row=i + 2, column=6).value)
        if ip in vip_1 or ip == ip_1:
            Module = str(sheet_crm.cell(row=i + 2, column=2).value)
            Host = str(sheet_crm.cell(row=i + 2, column=4).value)
            IP = str(sheet_crm.cell(row=i + 2, column=5).value)
            Virtual_IP = str(sheet_crm.cell(row=i + 2, column=6).value)
            OSPwd = str(sheet_crm.cell(row=i + 2, column=7).value)
            Instance = str(sheet_crm.cell(row=i + 2, column=9).value)
            DBPwd = str(sheet_crm.cell(row=i + 2, column=10).value)
            Other = str(sheet_crm.cell(row=i + 2, column=12).value)

            # 把取到的数，插入PrettyTable，格式化输出
            table = PrettyTable(
                ['Environment', 'Module', 'Host', 'IP', 'Virtual_IP', 'OSPwd', 'Instance','DBPwd', 'Other'])
            table.add_row([env, Module, Host, IP, Virtual_IP,OSPwd,Instance,DBPwd, Other])
            print(table)

    #bss 计费环境
    #workbook = openpyxl.load_workbook(r'/Users/rhl/other/134.96.253.56/134.96.263.56/svn4dba/文档资料/环境信息/生产环境服务器列表.xlsx')
    sheet_jf = workbook['BSS_计费_PROD']
    env = 'BSS_JF_PROD'
    max_num_crm = sheet_jf.max_row
    for i in range(max_num_crm):
        ip_1 = sheet_jf.cell(row=i + 2, column=6).value
        vip_1 = str(sheet_jf.cell(row=i + 2, column=7).value)
        if ip in vip_1 or ip == ip_1:
            Module = str(sheet_jf.cell(row=i + 2, column=2).value)
            #Host = str(sheet_jf.cell(row=i + 2, column=5).value)
            IP = str(sheet_jf.cell(row=i + 2, column=6).value)
            Virtual_IP = str(sheet_jf.cell(row=i + 2, column=7).value)
            OSPwd = str(sheet_jf.cell(row=i + 2, column=9).value)
            Instance = str(sheet_jf.cell(row=i + 2, column=11).value)
            DBPwd = str(sheet_jf.cell(row=i + 2, column=12).value)
            Other = str(sheet_jf.cell(row=i + 2, column=13).value)
            port = str(sheet_jf.cell(row=i + 2, column=10).value)
            User1 = str(sheet_jf.cell(row=i + 2, column=14).value)
            User2 = str(sheet_jf.cell(row=i + 2, column=15).value)

            # 把取到的数，插入PrettyTable，格式化输出
            table = PrettyTable(['Environment', 'Module', 'IP', 'Virtual_IP', 'OSPwd', 'Instance','DBPwd', 'port'])
            table2=PrettyTable(['Other'])
            table3=PrettyTable(['User1','User2'])
            table.align["Environment"] = "l"
            table.align["Module"] = "l"
            table.align["IP"] = "l"
            table.align["Virtual_IP"] = "l"
            table.align["OSPwd"] = "l"
            table.align["Instance"] = "l"
            table.align["DBPwd"] = "l"
            table.align["port"] = "l"
            table2.align["Other"] = "l"
            table3.align["User1"] = "l"
            table3.align["User2"] = "l"
            table.add_row([env, Module, IP, Virtual_IP,OSPwd,Instance,DBPwd, port])
            table2.add_row([Other])
            table3.add_row([User1,User2])
            print(table)
            print(table2)
            print(table3)

    #oss环境
    #workbook = openpyxl.load_workbook(r'/Users/rhl/other/134.96.253.56/134.96.263.56/svn4dba/文档资料/环境信息/生产环境服务器列表.xlsx')
    sheet_jf = workbook['OSS_PROD']
    env = 'OSS3.0'
    max_num_crm = sheet_jf.max_row
    for i in range(max_num_crm):
        ip_1 = sheet_jf.cell(row=i + 2, column=5).value
        vip_1 = str(sheet_jf.cell(row=i + 2, column=6).value)
        if ip in vip_1 or ip == ip_1:
            Module = str(sheet_jf.cell(row=i + 2, column=2).value)
            Host = str(sheet_jf.cell(row=i + 2, column=4).value)
            IP = str(sheet_jf.cell(row=i + 2, column=5).value)
            Virtual_IP = str(sheet_jf.cell(row=i + 2, column=6).value)
            OSPwd = str(sheet_jf.cell(row=i + 2, column=7).value)
            Instance = str(sheet_jf.cell(row=i + 2, column=11).value)
            DBPwd = str(sheet_jf.cell(row=i + 2, column=9).value)
            User1=str(sheet_jf.cell(row=i + 2, column=13).value)
            User2=str(sheet_jf.cell(row=i + 2, column=14).value)
            Other = str(sheet_jf.cell(row=i + 2, column=15).value)

            # 把取到的数，插入PrettyTable，格式化输出
            table = PrettyTable(
                ['Environment', 'Module', 'Host', 'IP', 'Virtual_IP', 'OSPwd', 'Instance','DBPwd','User1','User2', 'Other'])
            table.add_row([env, Module, Host, IP, Virtual_IP,OSPwd,Instance,DBPwd,User1,User2, Other])
            print(table)

    #小系统上云的管理平台等机器
   #workbook = openpyxl.load_workbook(r'/Users/rhl/other/134.96.253.56/134.96.263.56/svn4dba/文档资料/环境信息/生产环境服务器列表.xlsx')
    sheet_xxt = workbook['SOTC_PROD_小系统上云']
    env = 'SOTC'
    max_num_xxt = sheet_xxt.max_row
    for i in range(max_num_xxt):
        ip_1 = sheet_xxt.cell(row=i + 2, column=5).value
        vip_1 = str(sheet_xxt.cell(row=i + 2, column=6).value)
        if ip in vip_1 or ip == ip_1:
            Module = str(sheet_xxt.cell(row=i + 2, column=2).value)
            # Host = str(sheet_xxt.cell(row=i + 2, column=4).value)
            IP = str(sheet_xxt.cell(row=i + 2, column=5).value)
            Virtual_IP = str(sheet_xxt.cell(row=i + 2, column=6).value)
            OSPwd = str(sheet_xxt.cell(row=i + 2, column=7).value)
            Instance = str(sheet_xxt.cell(row=i + 2, column=9).value)
            DBPwd = str(sheet_xxt.cell(row=i + 2, column=8).value)
            User1=str(sheet_xxt.cell(row=i + 2, column=13).value)
            Other = str(sheet_xxt.cell(row=i + 2, column=12).value)

            # 把取到的数，插入PrettyTable，格式化输出
            table = PrettyTable(
                ['Environment', 'Module', 'IP', 'Virtual_IP', 'OSPwd', 'Instance','DBPwd', 'Other','User1'])
            table.add_row([env, Module, IP, Virtual_IP,OSPwd,Instance,DBPwd, Other,User1])
            print(table)

    #大测试环境
    workbook = openpyxl.load_workbook(r'/Users/rhl/other/134.96.253.56/134.96.263.56/svn4dba/文档资料/环境信息/其他环境服务器列表.xlsx')
    sheet_dcs = workbook['大测试环境']
    env = '大测试环境'
    max_num_crm = sheet_dcs.max_row
    for i in range(max_num_crm):
        ip_1 = sheet_dcs.cell(row=i + 2, column=6).value
        vip_1 = str(sheet_dcs.cell(row=i + 2, column=7).value)
        if ip in vip_1 or ip == ip_1:
            Module = str(sheet_dcs.cell(row=i + 2, column=2).value)
            # Host = str(sheet_dcs.cell(row=i + 2, column=5).value)
            IP = str(sheet_dcs.cell(row=i + 2, column=6).value)
            Virtual_IP = str(sheet_dcs.cell(row=i + 2, column=7).value)
            OSPwd = str(sheet_dcs.cell(row=i + 2, column=12).value)
            Instance = str(sheet_dcs.cell(row=i + 2, column=15).value)
            DBPwd = str(sheet_dcs.cell(row=i + 2, column=14).value)
            Other = str(sheet_dcs.cell(row=i + 2, column=16).value)



            # 把取到的数，插入PrettyTable，格式化输出
            table = PrettyTable(
                ['Environment', 'Module', 'IP', 'Virtual_IP', 'OSPwd', 'Instance','DBPwd', 'Other'])
            table.add_row([env, Module, IP, Virtual_IP,OSPwd,Instance,DBPwd, Other])
            print(table)


if __name__ == '__main__':
    ip=input('please input ip:')
    excel(ip)
