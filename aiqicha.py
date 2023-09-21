import requests
# import json
import urllib3
from tabulate import tabulate
import argparse
import time
import datetime
import openpyxl
import os
# from lxml import etree

# 禁用证书验证的警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 获取企业页面信息 作废 接口不对
# def Pid_title(pid):
#     url = f"http://aiqicha.baidu.com/company_detail_{pid}"
#     headers = {
#         "Connection":"close",
#         "Cookie":"BAIDUID=21A6C0E62F9DD20A99E72C7B55CD2F4F:FG=1; __jdg_yd=lTM-TogKuTwn0mXPcGT6ZUJ8F15Yr0ewOCElznS0M70B4DA-u%2AQmKc1BX0VrDPm59X3daeht54aevqlP8nsnShOPSjsz-HQvXeZd7AkDRMY",
#         "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/117.0"
#         }
#     res = requests.get(url=url, headers=headers, verify=False)
#     html = res.text
#     print(html)
#     tree = etree.HTML(html)
#     print(tree)
#     title_element = tree.xpath('/html/head/title/text()')
#     print(title_element)
#     # result = html.xpath('/html/body/div[1]/div[2]/div/div[2]/div[1]/div[1]/div[2]/div[2]/h1/text()')
#     if title_element:
#         # title = title_element[0].text
#         # company_name = title.split(" - ")[0]
#         # print(company_name)
#         return title_element
#     else:
#         return pid

# 控股数据>20条数据
def Pid_KongGu_s(pid,page):
    url = "https://aiqicha.baidu.com/detail/holdsAjax?pid={}&p={}&size=20&confirm=".format(pid,page)
    # 控股数据请求头
    headers = {
            "Accept-Encoding":"gzip, deflate",
            "Accept-Language":"zh-CN,zh;q=0.9",
            "Connection":"close",
            "Cookie":"BAIDUID=F4BFD69380B19F9D14B978E3A37E3770:FG=1; BDUSS_BFESS=JlZ0pPZ353WXZjNHp1RE9kVDJCR1kycEJzQnVKVktDOTktemhVZ2FMdU0wVEJsRVFBQUFBJCQAAAAAAAAAAAEAAABAC~6n5ZaG4pyU56eR5oqAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIxECWWMRAllcH; ",
            "Referer":"https://aiqicha.baidu.com/company_detail_28806871089320",
            "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36"
            }
    res = requests.get(url=url, headers=headers, verify=False)
    json_data = res.json()
    table_data = []
    for item in json_data['data']['list']:
        ent_name = item['entName']
        p_id = item['pid']
        logo = item['logo']
        proportion = item['proportion']
        table_data.append([ent_name, p_id, logo, proportion])
    return table_data


# 控股数据<20条数据（首页数据）
def Pid_KongGu(pid):
    url = f"https://aiqicha.baidu.com/detail/holdsAjax?pid={pid}&p=1&size=20&confirm="
    # 控股数据请求头
    headers = {
            "Accept-Encoding":"gzip, deflate",
            "Accept-Language":"zh-CN,zh;q=0.9",
            "Connection":"close",
            "Cookie":"BAIDUID=F4BFD69380B19F9D14B978E3A37E3770:FG=1; BDUSS_BFESS=JlZ0pPZ353WXZjNHp1RE9kVDJCR1kycEJzQnVKVktDOTktemhVZ2FMdU0wVEJsRVFBQUFBJCQAAAAAAAAAAAEAAABAC~6n5ZaG4pyU56eR5oqAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIxECWWMRAllcH; ",
            "Referer":"https://aiqicha.baidu.com/company_detail_28806871089320",
            "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36"
            }
    res = requests.get(url=url, headers=headers, verify=False)
    json_data = res.json()

    # total = json_data['data']['total'] # 获取总数
    totalNum = json_data['data']['totalNum'] # 获取总数
    if totalNum <= 0:
        # 判断是否存在数据
        return False
    else:
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        #   
        file_name_path = "./results/{}_控股子公司_{}.xlsx".format(pid,current_time)

        pageCount = json_data['data']['pageCount'] # 总计页数
        table_data = []
        for item in json_data['data']['list']:
            ent_name = item['entName']
            p_id = item['pid']
            logo = item['logo']
            proportion = item['proportion']
            table_data.append([ent_name, p_id, logo, proportion])

        # 判断数据是否大于1页数据
        if pageCount == 1:
            print(table_data)
            if data_saver_excel(file_name_path,table_data):
                print("[*]第1页保存完成")
            else:
                print("[!]第1页保存失败，启用终端数据打印操作")
                data_tables(table_data)
        else:
            if data_saver_excel(file_name_path,table_data):
                print("[*]第1页保存完成")
                # 读取读取>=2页数据
                for page in range(2, pageCount + 1):
                    print("[*]防反爬机制，休眠3秒后开始下一页...")
                    time.sleep(3)
                    # 调用 >=2页数据 函数
                    table_data = Pid_KongGu_s(pid,page)
                    if data_saver_excel(file_name_path,table_data):
                        print(f"[*]第{page}页保存完成")
                    else:
                        print(f"[!]第{page}页保存失败，启用终端数据打印操作")
                        data_tables(table_data)
                return True
            else:
                print(f"[!]第1页保存失败，启用终端数据打印操作")
                data_tables(table_data)



# 执行终端数据打印 函数
def data_tables(table_data):
    headers = ['entName', 'pid', 'logo', 'proportion']
    table = tabulate(table_data, headers, tablefmt='pretty', stralign='left')
    print(table)

# 创建唯一表格文档
def data_saver_excel(file_name_path,table_data):
    try:
        # 判断文件是否存在
        if os.path.exists(file_name_path):
            workbook = openpyxl.load_workbook(file_name_path)
            sheet = workbook.active
        else:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            # 设置表头
            headers = ['公司', 'Pid', 'logo', '股份']
            sheet.append(headers)
        # 写入数据
        for row in table_data:
            sheet.append(row)
        # 保存Excel文件
        workbook.save(file_name_path)
        return True
    except Exception as e:
        print(f'{e}\n')
        return False


if __name__ == '__main__':

    parser = argparse.ArgumentParser(description='爱企查批量查询控股子公司信息')
    group = parser.add_mutually_exclusive_group()
    group.add_argument('-pid', '--pid', help='需要查询的企业pid')
    # group.add_argument('-pn', '--pidname', help='根据pid查询企业名称')
    args = parser.parse_args()

    # 获取文件路径参数的值
    pid = args.pid
    # pidname = args.pidname
    if pid:
        if Pid_KongGu(pid):
            print("[*]--完成--")
        else:
            print("[!]--失败--")
    # elif pidname:
    #     print(Pid_title(pidname))
    else:
        print('请提供文件路径参数。使用 -h 或 --help 选项获取帮助信息。')
