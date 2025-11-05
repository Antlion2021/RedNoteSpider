import sys
import time
import random
import pandas as pd
from pathlib import Path
from DrissionPage import ChromiumPage, ChromiumOptions

# 超参数
web_head = "https://pgy.xiaohongshu.com/solar/post-trade/data-center/deal/note/"
listen_list = ["/detail?orderid=", "/core_data"]
targets_dir = 'target'
output_dir = 'outputs'


def get_url(s):
    s = str(s)
    return web_head + s


def find_xlsx_files(directory_path):
    path = Path(directory_path)
    return list(path.glob('*.xlsx'))


# 设置自动化浏览器配置
# 设置Edge浏览器路径 -> 一般默认路径在C:\Program Files (x86)\Microsoft\Edge\Application文件夹中
options = ChromiumOptions().set_paths(browser_path="C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
try:
    page = ChromiumPage(options)
except Exception as e:
    print("程序没在系统里找到 Chrome 浏览器")
    time.sleep(30)
    sys.exit(1)

# 打开RedNote蒲公英网页 并验证登入
page.get('https://pgy.xiaohongshu.com/solar/pre-trade/home')
login_element = page.ele('@class=login-btn')  # 定位登入元素判断是否已经登入
try:
    if login_element:
        print('Cookie过期需要手动登入,登入超时时间为60s')
        login_element.click()
    # 等待登入,60秒后超时程序结束
    login_user = page.ele('@class=user-info-content', timeout=60)
    if not login_user:
        print('登入验证超时')
        time.sleep(30)
        sys.exit(1)
except Exception as e:
    print('登入验证时出错')
    time.sleep(30)
    sys.exit(1)
else:
    time.sleep(random.uniform(0.5, 1))
    print('验证登入成功')

# # # # # # # # # # # # # # # # #
# 获取多个表格中的链接准备开始采集信息 #
# # # # # # # # # # # # # # # # #

target_files = find_xlsx_files(targets_dir)
for target_file in target_files:
    # 通过订单号获取对应数据页面Url
    datas_list = []                            # 临时存储数据字典
    output_name = Path(target_file.name).stem  # 输出文件名对应输入文件名
    try:
        target_df = pd.read_excel(target_file)
        url_data = target_df['订单号']
        url_list = url_data.to_list()
        url_list = [get_url(url) for url in url_list]
    except RuntimeError as e:
        print(f'{target_file}缺少[订单号]字段或订单号不以文本形式存储')
        continue

    # # # # # # # # # # # # # # # # # # # # # #
    # 通过url采集数据字典并保存于字典列表datas_list #
    # # # # # # # # # # # # # # # # # # # # # #
    number = 1  # 序号
    for url in url_list:
        try:
            page.listen.start(listen_list)
            page.get(url)
            packet = page.listen.wait(count=2)
            core_data = packet[0].response.body['data']
            detail_data = packet[1].response.body['data']
            core_data = detail_data | core_data

            data_dict = {'序号': number,
                         '账号名': core_data['userName'],
                         '主页链接': 'https://www.xiaohongshu.com/user/profile/' + core_data['userId'],
                         '发布链接': core_data['noteLink'],
                         '任务名': core_data['taskName'],
                         '曝光': core_data['impNum']['data'],
                         '阅读': core_data['appReadNum']['data'],
                         '点赞': core_data['likeNum']['data'],
                         '收藏': core_data['favNum']['data'],
                         '评论': core_data['cmtNum']['data'],
                         '分享': core_data['shareNum']['data'],
                         '总互动': core_data['engageNum']['data']}
            datas_list.append(data_dict)
            number += 1
            time.sleep(random.uniform(1.5, 2))
        except Exception as e:
            print(f'url数据异常,跳过{url}')
            continue

    #  # # # # # # # # # # # # # # #
    #  创建表格文件并将数据字典写入表格  #
    #  # # # # # # # # # # # # # # #
    data_frame = pd.DataFrame(datas_list)
    print(f'{target_file}数据如下,开始写入并生成xlsx\n{data_frame}')
    try:
        data_frame.to_excel(excel_writer=f'{output_dir}/{output_name}.xlsx', index=False)
    except Exception as e:
        print('写入失败,请检查文件权限')
        time.sleep(2)
    else:
        print(f'写入成功,表格{output_name}完成')

print('程序正常完成')
time.sleep(20)
