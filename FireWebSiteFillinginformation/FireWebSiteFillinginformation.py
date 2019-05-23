# encoding=utf-8
import time
import click
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities


# 前台开启浏览器模式
def openChrome():
    # 加启动配置
    option = webdriver.ChromeOptions()
    option.add_argument('disable-infobars')
    # 打开chrome浏览器
    driver = webdriver.Chrome(chrome_options=option)
    return driver


def open360Chrome():
    # 360浏览器的地址
    __browser_url = r'C:\Users\Administrator\AppData\Roaming\360se6\Application\360se.exe'
    # 加启动配置
    option = Options()
    option.binary_location = __browser_url
    option.add_argument('disable-infobars')
    # 打开chrome浏览器
    driver = webdriver.Chrome(chrome_options=option)
    return driver


def openFirefox():
    # 加启动配置
    option = webdriver.FirefoxOptions()
    option.add_argument('disable-infobars')
    # 打开chrome浏览器
    driver = webdriver.Firefox(firefox_options=option)
    return driver


def openIe():
    # 加启动配置
    option = webdriver.IeOptions()
    option.add_argument('disable-infobars')
    # 打开chrome浏览器
    driver = webdriver.Ie(ie_options=option)
    return driver


# 授权操作
def operation(driver):
    url = "https://www.baidu.com/"
    driver.get(url)
    # 找到输入框并输入查询内容
    elem = driver.find_element_by_xpath('//*[@id="kw"]')
    elem.send_keys("python")
    # 提交表单
    driver.find_element_by_xpath('//*[@id="su"]').click()

    # time.sleep(3)
    # handles =driver.window_handles
    url = 'https://www.so.com/'
    driver.get(url)
    # driver.switch_to_window(handles[0])

    print('查询操作完毕！')


# 消防安全户籍化管理系统_山东省公安消防总队(HJGL_1.0.2_r_b20170323登陆函数
def firewebsitefillinginformlogin(driver):
    # 登陆页面
    urldl = "http://124.128.8.246:85/FrameSet/Login.aspx"
    driver.get(urldl)
    username = driver.find_element_by_xpath('//*[@id="txtUserName"]')
    username.send_keys('37029827')
    password = driver.find_element_by_xpath('//*[@id="txtPwd"]')
    password.send_keys('729')
    yanzhengma = driver.find_element_by_xpath('//*[@id="txtValidateCode"]')
    yanzhengma.send_keys('')
    time.sleep(6)
    dengluanniu = driver.find_element_by_xpath('//*[@id="ImageButton1"]').click()
    time.sleep(3)


# 消防安全户籍化管理系统_山东省公安消防总队(HJGL_1.0.2_r_b20170323)填日报表函数
def firewebsitefillinginformforday(driver, year, month, bday, eday):
    day = bday
    while day >= bday and day <= eday:
        # 构造日期和编号字符串
        if day > 0 and day < 10:
            datestr = "{0}年{1:0>2d}月0{2:d}日".format(year, month, day)
        else:
            datestr = "{0}年{1:0>2d}月{2:>2d}日".format(year, month, day)
        datestrbh = '{0} {1:0>2d}'.format(year, month)

        # 新增与提交页面
        url = 'http://124.128.8.246:85/JCDAPage/XFGZJLPage/FHXC_SimpleAddPage.aspx'
        driver.get(url)
        # 点击“新增”按钮
        # elemxinzeng = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_Button1"]').click()
        # elem.send_keys("python")

        # 巡查日期
        elemrq = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_txtRQ"]')
        elemrq.send_keys(datestr)

        # 编号
        elembh = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_txtBH"]')
        elembh.send_keys(datestrbh)

        # 巡查员
        elemxcy = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_txtXCY"]')
        if day % 2 == 1:
            elemxcy.send_keys('李在峰')
        else:
            elemxcy.send_keys('尹保全')

        # 巡查次数
        elemxccs = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_txtXCCS"]')
        elemxccs.send_keys('2')

        # 巡查总体情况
        elemfxwt = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_txtFXWT"]')
        elemfxwt.send_keys('正常')

        # 核查人
        elemjcr = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_txtJCR"]')
        if day % 2 == 1:
            elemjcr.send_keys('李在峰')
        else:
            elemjcr.send_keys('尹保全')

        # 检查日期
        elemjcrq = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_txtJCRQ"]')
        elemjcrq.send_keys(datestr)

        # 主管人
        elemzgr = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_txtZGR"]')
        elemzgr.send_keys('张志霞')

        # 提交表单
        driver.find_element_by_xpath('//*[@id="ctl00_MainContent_btnAdd"]').click()

        time.sleep(6)

        # 处理弹窗
        try:
            driver.switch_to.alert.accept()
        except Exception:
            pass

        # handles =driver.window_handles
        # url = 'https://www.so.com/'
        # driver.get(url)
        # driver.switch_to_window(handles[0])

        print('{0}年{1}月{2}日填写完毕！'.format(year, month, day))
        day += 1


# 消防安全户籍化管理系统_山东省公安消防总队(HJGL_1.0.2_r_b20170323)填月报表函数
def firewebsitefillinginformformonth(driver, year, bmonth, emonth):
    month = bmonth
    while month >= bmonth and month <= emonth:
        # 构造日期字符串
        strdate = '{0}年{1:0>2}月{2:0>2}日'.format(year, month, 1)
        # 新增与提交页面
        url = 'http://124.128.8.246:85//WBBAPage/SHDW_XFSSWBBAPage.aspx'
        driver.get(url)

        # 月工作——设施维保
        # time.sleep(3)
        # elemsswb = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_lblWB"]').click()

        # 首页——月工作——设施维护——“新增”按钮
        elemxz = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_btn_XKSCYRY"]').click()

        # 维保日期
        elemwbrq = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_txt_BAYF"]')
        elemwbrq.click()
        driver._switch_to.frame(driver.find_element_by_xpath('//*[@id="_my97DP"]/iframe'))
        driver.find_element_by_xpath('//*[@id="dpOkInput"]').click()
        driver._switch_to.default_content()
        js = 'document.getElementById("ctl00_MainContent_txt_BAYF").removeAttribute("readonly");'
        driver.execute_script(js)
        elemwbrq.clear()
        elemwbrq.send_keys(strdate)

        # 维保企业
        driver._switch_to.frame(driver.find_element_by_xpath('//*[@id="ym-ml"]/div/div/div/iframe'))
        driver.find_element_by_xpath('//*[@id="ctl00_MainContent_grid"]/tbody/tr[2]/td[4]/a').click()
        driver._switch_to.default_content()

        # 室内消防系统
        elemsnxf = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_rpt_XFSS_ctl08_txt_BZ"]')
        elemsnxf.send_keys("无水")

        # 室外消防系统
        elemswxf = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_rpt_XFSS_ctl09_txt_BZ"]')
        elemswxf.send_keys("无水")

        # 疏散指示标志和应急照明
        elemyjzm = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_rpt_XFSS_ctl10_radio_ZCYX"]').click()

        # 灭火器
        elemmhq = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_rpt_XFSS_ctl16_radio_ZCYX"]').click()

        # 消防控制室值班人员持证上岗及熟练操作情况
        elemczsg = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_txt_XKSZBRYQK"]')
        elemczsg.send_keys('无消防控制室。')

        # 维护保养人员(签名)
        elemwbry = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_txt_WBRY"]')
        if month % 2 == 1:
            elemwbry.send_keys('李在峰')
        else:
            elemwbry.send_keys('尹保全')
        # 维护保养人员(签名)日期
        elemwbrq1 = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_txt_WBRQ1"]')
        elemwbrq1.click()
        driver._switch_to.frame(driver.find_element_by_xpath('//*[@id="_my97DP"]/iframe'))
        driver.find_element_by_xpath('//*[@id="dpOkInput"]').click()
        driver._switch_to.default_content()
        js1 = 'document.getElementById("ctl00_MainContent_txt_WBRQ1").removeAttribute("readonly");'
        driver.execute_script(js1)
        elemwbrq1.clear()
        elemwbrq1.send_keys(strdate)
        elemwbk = driver.find_element_by_xpath('//*[@id="aspnetForm"]/div[3]')
        elemwbk.click()

        # 消防安全责任人或消防安全管理人(签名)
        elemaqzrr = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_txt_XFAQZRR"]')
        elemaqzrr.send_keys('张志霞')
        # time.sleep(3)

        # 维护保养人员(签名)日期
        elemwbrq2 = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_txt_WBRQ2"]')
        elemwbrq2.send_keys(strdate)
        elemwbk.click()
        # 责任人或管理人(签名)
        elemzrr = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_txt_ZRR"]')
        elemzrr.send_keys('高云兵')
        # time.sleep(3)

        # 维护保养人员(签名)日期
        elemwbrq3 = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_txt_WBRQ3"]')
        elemwbrq3.send_keys(strdate)
        elemwbk.click()

        # 我单位报告备案的信息及数据真实有效，并对此负责。
        driver.find_element_by_xpath('//*[@id="ctl00_MainContent_cb_ZRSM"]').click()

        # 提交表单
        driver.find_element_by_xpath('//*[@id="ctl00_MainContent_btnTY"]').click()

        time.sleep(6)

        # 处理弹窗
        try:
            driver.switch_to.alert.accept()
        except Exception:
            pass
        time.sleep(6)
        # 处理弹窗
        try:
            driver.switch_to.alert.accept()
        except Exception:
            pass

        print('{0}年{1}月{2}日填写完毕！'.format(year, month, month))
        month += 1


# 方法主入口
if __name__ == '__main__':
    # 加启动配置
    driver = openChrome()
    firewebsitefillinginformlogin(driver)
    # 日报表
    firewebsitefillinginformforday(driver, 2019, 5, 23, 31)
    # 月报表
    # firewebsitefillinginformformonth(driver, 2019, 3, 5)

