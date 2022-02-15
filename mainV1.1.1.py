import requests
import time
import json
import xlwt
import pandas as pd

# 常用参数
cookie = ""  # cookie手工更新
count = 0  # 过程输出屏幕计数
key_word = "eg"  # 查询关键字（业务号码框），暂时单次输入一个
data_start = "2022-02-01 00:00:00"  # 查询开始时间
data_end = "2022-02-12 23:59:59"  # 查询结束时间
page_size = 99999  # 单次查询最大条数
page = 1
time_delay = 0.5  # 查询延迟
file_name = "result.xlsx"


# cookie每次登录后都会变更
def header_get(cookie):
    """
    功能: 整合header信息，进行系统登录
    """
    head = {
        # UA和cookie设置，
        "Cookie": cookie,
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
                      " AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.82 Safari/537.36"
    }
    return head


def set_config():
    """
    功能: 输入Cookie、关键字、开始结束时间等相关查询参数
    """
    global key_word, data_start, data_end, page_size, page, time_delay, cookie, file_name, byniu
    while True:
        print("【查询信息输入，请按照提示格式进行输入，该版本无矫错机制，格式错误，程序退出】")
        temp1 = input("---- 请输入自己的Cookie(必需):")
        if temp1 == '':
            pass
        else:
            cookie = temp1

        temp2 = input(f"---- 请输入查询关键字eg（默认）/vpn...,回车保持默认:")
        if temp2 == '':
            key_word = "eg"
        else:
            key_word = temp2

        temp3 = input(f"---- 请输入查询起始时间,格式:2022-02-01(默认),回车保持默认:")
        if temp3 == '':
            data_start = "2022-02-01 00:00:00"
        else:
            data_start = temp3 + " 00:00:00"

        temp4 = input(f"---- 请输入查询结束时间,格式:2022-02-13(默认),回车保持默认:")
        if temp4 == '':
            data_end = "2022-02-12 23:59:59"
        else:
            data_end = temp4 + " 23:59:59"

        temp5 = input(f"---- 请输入存放数据的文件名(无需输扩展名),回车默认 result.xlsx:")
        if temp5 == '':
            file_name = "result.xlsx"
        else:
            file_name = temp5 + ".xlsx"

        temp6 = input(f"---- 请输入查询延迟0.01/0.1/2...(单位秒,防ban),回车保持0.5s:")
        if temp6 == '':
            time_delay = 0.5
        else:
            time_delay = float(temp6)

        print(f"**** 查询信息输入完毕：[关键字]{key_word}, [时间]{data_start}至{data_end},"
              f" [文件名]{file_name}, [查询延时]{time_delay}秒")
        confirm = input("**** 请确认是否变更查询信息(y/n),回车=n:")
        if confirm == "y":
            continue
        elif confirm == "n" or confirm == '':
            break
    print('-' * 99 + "\n【━(￣ー￣*|||━━开始定单提取】")


def login_sys():
    """
    功能：服开系统登录，默认使用niuhao账号，需收取验证码并输入后使用
    暂时不启用，暂时不启用，暂时不启用，使用cookie登录
    """
    # 收验证码前验证用户名密码网址
    check_url = "http://136.142.1.135:8080/iom-web/codeController4HB/checkIsNeedValidNum.inf"
    # 获得验证码网址
    code_url = "http://136.142.1.135:8080/iom-web/codeController4HB/getCode.inf"
    # 登录网址
    login_url = "http://136.142.1.135:8080/iom-web/login4HB/login.do"

    # 用户名+密码, 暂时使用hsniuhao账号
    user_data = {"username": "VvOt4KN1IBk2bvsMYX9RHg==", "password": "tmUIuCRx0y46mxgXLPPojg=="}
    # 验证码发送所需信息
    code_data = {"userName": "hsniuhao", "needNumFlag": 1}
    # 登录部门信息，验证码validateCode预留为空，后面添加
    login_data = {"username": "VvOt4KN1IBk2bvsMYX9RHg==", "password": "tmUIuCRx0y46mxgXLPPojg==", "validateCode": None}

    session = requests.session()
    # 获取收验证码前验证
    resp_check = session.post(check_url, data=user_data, headers=header)
    print(resp_check.json())
    time.sleep(2)
    # 获取验证码
    resp_code = session.post(code_url, data=code_data, headers=header)
    # print(resp_code.text)
    # 手机收到验证码后，手工输入验证码，并添加到code_num中
    code_num = input("请输入你收到的验证码：")
    login_data["validateCode"] = code_num
    time.sleep(2)
    # 尝试登录
    resp_login = session.post(login_url, data=login_data, headers=header)
    print(resp_login.text)

    return session


def export_excel(export):
    """
    功能: 将字典列表生成excel文件，内置固定字段，可设置变量进行字段调整）
    返回值: 无返回值，会生成一个excel文件
    """
    print("--︿(￣︶￣)︿-- 开始生成excel文件")
    # 将字典列表转换为DataFrame
    pf = pd.DataFrame(list(export))
    # 指定字段顺序
    order = ['定单编码', '接入号', '申请事项', '竣工时间', '端口速率', 'IP地址数量']
    pf = pf[order]
    # 将列名替换为中文（此处本身为中文，保留函数通用性）
    columns_map = {
        '定单编码': '定单编码',
        '接入号': '接入号',
        '申请事项': '申请事项',
        '竣工时间': '竣工时间',
        '端口速率': '端口速率',
        'IP地址数量': 'IP地址数量'
    }
    pf.rename(columns=columns_map, inplace=True)
    # 指定生成的Excel表格名称
    file_path = pd.ExcelWriter(file_name)
    # 替换空单元格
    pf.fillna(' ', inplace=True)
    # 输出
    pf.to_excel(file_path, encoding='utf-8', index=False)
    # 保存表格
    file_path.save()
    print(f"--︿(￣︶￣)︿-- 生成excel文件完毕，文件目录为 程序目录/{file_name}")
    print("-" * 99)


def get_order_list(head):
    """
    功能: 按条件调取固网监控工单列表
    返回: 返回工单order_list，字典格式
    """
    order_get_url = "http://136.142.1.135:8080/iom-web/orderManagerController/getOrders.qry"
    order_search_data = {"_search": "false", "nd": 1644799007974, "pageSize": page_size, "page": 1, "sidx": None,
                         "sord": "asc", "paramMap[endDate]": data_end, "paramMap[isHis]": "0",
                         "paramMap[accNbr]": key_word, "paramMap[orderState]": "10F",
                         "paramMap[startDate]": data_start, "paramMap[hisFlag]": "0"}
    session = requests.session()
    resp_order = session.post(order_get_url, data=order_search_data, headers=head)
    order_list_dic = resp_order.json()
    records = order_list_dic["records"]
    print(f"--︿(￣︶￣)︿-- 提取固网定单列表完毕，共提取 {records} 个定单")
    return order_list_dic


def get_order_info():
    """
    功能: 提取固网定单清单中的"定单编码"、"接入号"、"申请事项"、"竣工时间"四项并保存到提取列表中
    返回值: 返回包含有效信息提取列表，列表类型
    """
    # 准备一个列表存储目标信息
    order_need_list = []
    # print(order_list_dic["rows"])
    # print(type(order_list_dic["rows"]))
    for o in order_list_dic["rows"]:
        # 提取需要的信息
        order_content = {"定单编码": o["id"], "接入号": o["accNbr"], "申请事项": o["orderTitle"], "竣工时间": o["finishDate"]}
        # 有效信息写入新的列表中
        order_need_list.append(order_content)
    print("--︿(￣︶￣)︿-- 固网定单有效信息提取完毕")
    return order_need_list


def get_order_product_detail(head, order_need_list_count):
    """
    功能: 根据定单编号提取产品页详情信息
    返回值: 返回提取到的所有产品详情信息product_detail_dic(字典格式)
    """
    # 循环在对应字段中寻找端口速率和IP地址数量信息，比对成功则把对应的值写到字典中

    product_detail_dic = {}
    product_get_url = "http://136.142.1.135:8080/iom-web/orderManagerController/queryOrderProduct.qry"
    session = requests.session()

    order_detail_data = {"outeKey": int(order_need_list_count["定单编码"]),
                         "orderId": int(order_need_list_count["定单编码"]), "hisFlag": 0}
    resp_detail = session.post(product_get_url, data=order_detail_data, headers=head)
    product_detail_dic = resp_detail.json()
    print("提取产品详情完毕,", end=" ")

    # print("提取产品页详情信息product_detail_dic:", product_detail_dic)
    return product_detail_dic


def get_product_info(product_detail_dic):
    """
    功能：提取产品详情页的端口速率和IP地址数量
    返回值：包含端口速率和IP地址数量的一个字典 product_need_dic
    """
    # 字典用于存放目标信息
    product_need_dic = {}
    # print(type(p)) # 输出：<class 'dict'>
    for p in product_detail_dic["serviceOrderDto"]["indepProdOrderAttrDtos"]:
        if p["name"] == "端口速率":
            product_need_dic["端口速率"] = p["characterValue"]
        elif p["name"] == "IP地址数量":
            product_need_dic["IP地址数量"] = p["characterValue"]
        elif p["name"] == "云专线端口速率":
            product_need_dic["端口速率"] = p["characterValue"]
    if not "端口速率" in product_need_dic:
        product_need_dic["端口速率"] = "无"
    if not "IP地址数量" in product_need_dic:
        product_need_dic["IP地址数量"] = "无"
    # 返回提取到的信息，格式为字典
    print("提取'端口''IP地址数量'完毕,", end=" ")
    return product_need_dic


def need_info_comb():
    """
    功能：将定单提取信息和对应的端口速率以及IP地址数量组合
    返回值：返回一个列表，或字典
    """
    # 循环提取后的定单列表
    global count
    count = 0
    need_list = []
    for order_need_list_count in order_need_list:
        # 添加访问产品详情页的延迟，以免被系统ban掉
        time.sleep(time_delay)
        count += 1
        # 按照定单编号提取产品页详情
        print(f"--︿(￣︶￣)︿-- 处理第{count}个定单:", end=" ")
        product_detail_dic = get_order_product_detail(header, order_need_list_count)
        # 在产品详情页提取“端口速率”、“IP数量”，并添加到提取存储字典中
        product_need_dic = get_product_info(product_detail_dic)
        # 将获取的端口速率等字典与定单提取信息合并，得到最终数据
        for k, v in product_need_dic.items():
            order_need_list_count[k] = v
        print(f"第{count}个定单整理完毕")
        need_list.append(order_need_list_count)
    print("--︿(￣︶￣)︿-- 所有定单处理完毕")
    return need_list


def continue_confirm():
    """
    功能: 单次查询完毕后，确认是否继续查询
    返回值: 返回定单数量
    """
    input_str = input("**** 本次查询完毕，是否继续查询(y/n),回车=y:")
    if input_str == "y" or input_str == '':
        print("\n【Ψ(￣∀￣)Ψ程序初始化，重新开始】")
        time.sleep(1)
        return True
    elif input_str == "n":
        print("【ヾ(￣▽￣)Bye~Bye~】")
        time.sleep(2)
        return False


def order_num_pick(order_list_dic):
    """
    返回值: 返回定单数量
    """
    if order_list_dic["records"] == 0:
        print("-" * 99)
    return order_list_dic["records"]


def introduce():
    """
    功能: 打印标题和更新日志
    """
    title = '-' * 99 + "\n (￣▽￣)～■富强、民主、文明、和谐、自由、平等、公正、法治、爱国、敬业、诚信、友善■～(￣▽￣)\n" + '-' * 99
    update_log = "\n# 更新日志 V1.0 2022-02-12\n**** 仅支持cookie登录, " \
                 "需先在网页登录服开系统\n**** 一次性查询最多500条定单,建议暂时缩短时间解决\n" \
                 "**** 结果输出到程序同目录result文件夹下result.txt" \
                 "\n**** 添加了用定单编码查询访问产品详情页的延迟，以免被系统ban掉\n" \
                 "**** 查询时间、关键字等需在源码手工调整\n" + '-' * 99 + "\n" \
                 "# 更新日志 V1.1 2022-02-13\n**** 为了一点点通用性~! 增加了简陋的交互" \
                 "\n**** 进一步解放~! 结果输出调整为excel\n**** 突破~! 无视500条限制 \n**** " \
                 "bug处理：云专线提取信息异常; " \
                 "未查询到定单后程序异常\n**** 更新计划：用户名密码收取验证码登录，自动获取Cookie\n" + '-' * 99 + "\n"\
                 + "# 更新日志 V1.1.1 2022-02-15\n**** 优化了一点点交互和代码\n" + '-' * 99
    introduce_info = title + update_log

    if cookie == "":
        print(introduce_info)
    else:
        print(title)


if __name__ == '__main__':
    while True:
        # 1. 提示信息输出
        introduce()
        time.sleep(1)
        # 2. 重要参数手工输入, 获取登录信息
        set_config()
        header = header_get(cookie)
        # 3. 按条件获取固网定单清单
        order_list_dic = get_order_list(header)
        # 4. 判断是否无定单被获取，有- 程序继续，无- 结束本次查询
        if order_num_pick(order_list_dic) == 0:
            # 确认是否继续查询
            if continue_confirm():
                continue
            else:
                break
        else:
            pass
        # 5. 提取固网定单清单中的"定单编码"、"接入号"、"申请事项"、"竣工时间"四项并保存到提取列表中
        order_need_list = get_order_info()
        # 6. 通过提取列表中的定单编码查询对应的产品页详情，并提取产品详情数据中的"端口速率"、"IP地址数量"两项写入对应定单号的need_list列表中
        need_list = need_info_comb()
        # 7. 输出到excel
        export_excel(need_list)
        # 8. 单次查询完毕后确认是否继续查询
        if continue_confirm():
            continue
        else:
            break
