# encoding=utf-8
from action.PageAction import *
import time


def TestSendMailWithAttachment():
    print("启动Firefox浏览器")
    open_browser()
    maximize_browser()
    print("访问126邮箱登录页...")
    visit_url("http://mail.126.com")
    sleep(5)
    click("xpath", '//*[@id="lbNormal"]')
    # 断言页面出现的关键内容
    assert_string_in_pagesource("126网易免费邮")
    print("访问126邮箱登录页成功")
    waitFrameToBeAvailableAndSwitchToIt("xpath", "//iframe[contains(@id,'x-URS-iframe')]")
    print("输入登录用户名")
    input_string("xpath", "//input[@name='email']", "AgilityToSR")
    print("输入登录密码")
    input_string("xpath", "//input[@name='password']", "AutoTest123")
    # 点击登录按钮
    click("id", "dologin")
    sleep(5)
    assert_title("网易邮箱")
    print("登录成功")
    switch_to_default_content()
    print("添加联系人")
    # 显示等待通讯录链接在页面上可见
    waitVisibilityOfElementLocated("xpath", "//div[text()='通讯录']")
    # 点击通讯录链接
    click("xpath", "//div[text()='通讯录']")
    # 点击新建联系人按钮
    click("xpath", "//span[text()='新建联系人']")
    # 输入联系人姓名
    input_string("xpath",
                 "//a[@title='编辑详细姓名']/preceding-sibling::div/input",
                 "lily")
    # 输入联系人邮箱
    input_string("xpath",
                 "//*[@id='iaddress_MAIL_wrap']//input",
                 "lily@qq.com")
    # 点击星标联系人复选框
    click("xpath",
          "//span[text()='设为星标联系人']/preceding-sibling::span/b")
    # 输入联系人手机号
    input_string("xpath", "//*[@id='iaddress_TEL_wrap']//dd//input",
                 "185xxxxx")
    # 输入联系人备注
    input_string("xpath", "//textarea", "朋友")
    # 点击确定按钮，保存新联系人
    click("xpath", "//span[text()='确 定']")
    time.sleep(1)
    # 断言页面是否出现关键内容
    assert_string_in_pagesource("lily@qq.com")
    print("添加联系人成功")
    # 点击首页链接，进入首页界面
    click("xpath", "//div[.='首页']")
    # 显示等待写信链接出现在页面上
    waitVisibilityOfElementLocated("xpath", "//span[text()='写 信']")
    # 点击写信链接按钮，进入写信页面
    click("xpath", "//span[text()='写 信']")
    print("开始写信")
    print("输入收件人地址")
    input_string("xpath",
                 "//div[contains(@id,'_mail_emailinput')]/input",
                 "38617389@qq.com")
    print("输入邮件主题")
    input_string("xpath",
                 "//div[@aria-label='邮件主题输入框，请输入邮件主题']/input",
                 "新邮件")
    # print("点击上传附件按钮")
    # click("xpath", "//div[contains(@title, '600首MP3')]")
    # print("上传附件")
    # paste_string("/Users/wangzhao/Desktop/自动化测试-吴老师/KeyWordAndDataDrivenFrameWork/testData/a.txt")
    # sleep(2)
    # press_enter_key()
    # sleep(2)
    # waitVisibilityOfElementLocated("xpath", '//span[.="上传完成"]')
    # print("上传附件成功")
    # 进入邮件正文的frame框体中
    waitFrameToBeAvailableAndSwitchToIt("xpath", "//iframe[@tabindex=1]")
    print("写入邮件正文")
    input_string("xpath", "/html/body", "发给光荣之路的一封信")
    # 退出邮件正文的frame框体，进入默认回话框体
    switch_to_default_content()
    print("写信完成")
    print("开始发送邮件...")
    click("xpath", "//header//span[text()='发送']")
    time.sleep(3)
    assert_string_in_pagesource("发送成功")
    print("邮件发送成功")
    close_browser()


if __name__ == '__main__':
    TestSendMailWithAttachment()
