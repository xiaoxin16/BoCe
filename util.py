#!/bin/python3
import json
import os
import platform
import re
from urllib.parse import urlparse

import dns
import requests
from requests.adapters import HTTPAdapter


# 选择源文件
from selenium import webdriver
from selenium.webdriver.common.by import By


def select_file(fp):
    import shutil
    files = os.listdir(fp)
    files_set = []
    print("文件列表：")
    for file in files:
        if os.path.splitext(file)[1] in [".xlsx", ".xls", '.docx', '.txt']:
            files_set.append(file)
            print('[', len(files_set), ']:', file)
    index = input("支持文件格式[.txt|.xlsx|.xls|.docx]，请输入对应文件的序号:")
    file_name = files_set[int(index) - 1]
    dst_file = file_name
    return dst_file


# 读取基本配置 in json format
def read_conf(conf_path):
    with open(conf_path, encoding='utf-8-sig') as f:
        data = json.load(f)
    return data


# 判断是否为IP
def isIP(str):
    p = re.compile('^((25[0-5]|2[0-4]\d|[01]?\d\d?)\.){3}(25[0-5]|2[0-4]\d|[01]?\d\d?)$')
    if p.match(str):
        return True
    else:
        return False


# 获取状态码和实际url（如发生跳转，返回跳转后的新url）
def get_status(url, quiet=False):
    import urllib3
    urllib3.disable_warnings()
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.141 Safari/537.36'
    }
    s = requests.Session()
    s.mount('http://', HTTPAdapter(max_retries=3))
    s.mount('https://', HTTPAdapter(max_retries=3))
    try:
        r = s.get(url=url, headers=headers, verify=False, timeout=(10, 5), allow_redirects=False)
        if not quiet:
            print(url, "\t:", r.status_code)
        s.close()
        if r.status_code in [301, 302]:
            return r.status_code, "SUCCESS", r.headers["Location"]
        else:
            return r.status_code, "SUCCESS", url
    except requests.exceptions.ConnectTimeout as e:
        if not quiet:
            print("ConnectTimeout:\t", url)
        return e.errno, "ConnectTimeout", url
    except requests.exceptions.ReadTimeout as e:
        if not quiet:
            print("ReadTimeout:\t", url)
        return e.errno, "ReadTimeout", url
    except requests.exceptions.ConnectionError as e:
        if not quiet:
            print("ConnectionError:\t", url)
        return e.errno, "ConnectionError", url
    except requests.exceptions.HTTPError as e:
        if not quiet:
            print("HTTPError:\t", url)
        return e.errno, "HTTPError", url
    except requests.exceptions.ConnectionResetError as e:
        if not quiet:
            print("ConnectionResetError:\t", url)
        return e.errno, "ConnectionResetError", url
    except Exception as e:
        if not quiet:
            print("未知错误:\t", url, e.__str__())
        s.close()
        return 500, "未知错误", url


# 标准化网址，return: status_code, new_url
def get_url_normalize(url, short_urls=None):
    url_new = url.replace(" ", "")
    default_scheme = "http"
    if urlparse(url).scheme == '':
        url_new = default_scheme + "://" + url
    if urlparse(url_new).hostname is None:
        url_new = "异常"
    elif "." not in urlparse(url_new).hostname:
        url_new = "异常"
    elif urlparse(url_new).hostname[0] == ".":
        url_new = "异常"
    elif len(urlparse(url_new).hostname) < 4:
        url_new = "异常"
    else:
        {}
    if url_new != "异常":
        # return 200, url_new
        domain = urlparse(url_new).hostname
        if domain in short_urls:
            # print("短域名 ", url_new)
            (code, status, r_url) = get_status(url_new, quiet=True)
            # print((code, status, r_url))
            return code, r_url
            # print("\t跳转网址：", url_new)
        elif len(domain) < 6:
            print("疑似短域名 ", domain)
            (code, status, r_url) = get_status(url_new, quiet=True)
            # print((code, status, r_url))
            return code, r_url
        else:
            return 200, url_new
    else:
        return 500, url_new


# 解析IP地址，返回状态码，以及对应的IP列表
def get_ips(domain, default_servers=None, quiet=False):
    ips = []
    myResolver = dns.resolver.Resolver()
    if default_servers:
        myResolver.nameservers = default_servers
    try:
        A = myResolver.resolve(domain, "A")
        for i in A.response.answer:
            for j in i.items:
                if j.rdtype == 1:
                    ips.append(j.address)
        if not quiet:
            print("IPS: ", domain, ips)
        return 1, list(set(ips))
    except dns.resolver.NXDOMAIN:
        if not quiet:
            print("[.] Resolved but no entry for " + str(domain))
        return 2, None
    except dns.resolver.NoNameservers:
        if not quiet:
            print("[-] Answer refused for " + str(domain))
        return 3, None
    except dns.resolver.NoAnswer:
        if not quiet:
            print("[-] No answer section for " + str(domain))
        return 4, None
    except dns.exception.Timeout:
        if not quiet:
            print("[-] Timeout")
        return 5, None


# 批量解析domain
def do_dns(data, fp, headless, timeout, PC):
    # print("\tDNS解析...")
    if data is None:
        return
    if self.conf["ip_fix"]:
        driver = init_driver(fp=fp, headless=headless, timeout=timeout, PC=PC)
        for value in data:
            if value[3] == "异常":
                continue
            if value[5] != '' and value[5] != "异常":
                continue
            domain = value[4]
            myaddr = []
            if type(domain) != str:
                print("value = ", value)
            if isIP(domain):
                myaddr.append(domain)
                value[5] = myaddr
                value[12] = "无"
            else:
                (retn, retv) = get_ips(domain, quiet=True)
                if retn == 1:
                    value[5] = retv
                else:
                    value[5] = "异常"  # IP
                    value[6] = "境外"
                    value[7] = "境外"
            if value[5] == "异常":
                continue
            dbpath = self.conf["conf"] + "/ipipfree.ipdb"
            dbc = ipdb.City(dbpath)
            city_str = dbc.find(value[5][0], "CN")
            jing_nei = "中国,香港,澳门,台湾"
            if city_str[0] in jing_nei:
                if city_str[1] in jing_nei:
                    value[6] = "境外"
                else:
                    value[6] = "境内"
            else:
                value[6] = "境外"
            value[7] = (city_str[0] + "·" + city_str[1])
            if self.conf["ip_fix"]:
                if value[6] == "境内":
                    (a, b) = ip_fix(driver, value[5][0], locat=value[6], detail=value[7])
                    value[6] = a
                    value[7] = b
                    time.sleep(1)
                    # if a == "境外":
                    #     print("IP %s %s 修订 OK..." % (value[6], value[7]))
                time.sleep(random.uniform(1, 2))
            time.sleep(0.1)
            # print(value)
            # print("processing domain %s" % value)
        if self.conf["ip_fix"]:
            driver.quit()
        return data


# 初始化diver chrome
def init_driver(fp, headless=False, timeout=None, PC=True):
    if fp is None:
        fp = "./conf"
    if platform.system() == "Windows":
        chrome_fp = os.path.join(fp, ".chromedriver.exe")
    else:
        chrome_fp = os.path.join(fp, ".chromedriver")
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    chrome_options = Options()
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    if not PC:
        mobileEmulation = {'deviceName': 'iPhone X'}
        chrome_options.add_experimental_option('mobileEmulation', mobileEmulation)
    chrome_options.headless = headless
    chrome_options.add_argument(
        'user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Safari/537.36')
    service = Service(chrome_fp)
    driver = webdriver.Chrome(options=chrome_options, service=service)
    with open('./conf/stealth.min.js') as f:
        js = f.read()
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": js
    })
    df_timeout = 20
    if timeout:
        df_timeout = timeout
    driver.set_page_load_timeout(df_timeout)
    driver.set_script_timeout(df_timeout)
    return driver


# 修订IP地址，ip138，返回：境内外，详情
def ip_fix(driver, ipstr, locat=None, detail=None, ip138=None, ipipnet=None):
    if driver is None:
        driver = init_driver()
    try:
        driver.get("https://ip138.com/iplookup.asp?ip=%s&action=2" % ipstr)
        driver.find_element(By.ID, "ip").clear()
        driver.find_element(By.ID, "ip").send_keys(ipstr)
        driver.find_element(By.CLASS_NAME, "input-button").click()
        matchObj = re.search(r'[{]"ASN归属地"(.?)*[}]', driver.page_source, re.M | re.I)
        if matchObj:
            strinfo = json.loads(matchObj.group())
            if strinfo["ip_c_list"][0]['ct'] == "中国":
                if strinfo["ip_c_list"][0]['prov'].replace('特别行政区', '') in "中国,香港,澳门,台湾":
                    locat = "境外"
                    detail = strinfo["ip_c_list"][0]['ct'] + "·" + strinfo["ip_c_list"][0]['prov']
                    # print("\t更新IP位置:", strinfo['ASN归属地'])
            else:
                locat = "境外"
                detail = strinfo["ip_c_list"][0]['ct'] + "·" + strinfo["ip_c_list"][0]['prov']
        # else:
        # print("\t%s fix 失败, %s %s" % (ipstr, locat, detail))
    except Exception as e:
        print("异常，当前函数：%s，%s" % ("ip_fix", e.__str__()))
    return locat, detail

