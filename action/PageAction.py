#encoding=utf-8
from selenium import webdriver
from config.VarConfig import firefoxDriverFilePath
from util.ObjectMap import getElement
# from util.ClipboardUtil import Clipboard
# from util.KeyBoardUtil import KeyboardKeys
from util.DirAndTime import *
from util.WaitUtil import WaitUtil
import time

# 定义全局driver变量
driver = None
# 全局的等待类实例对象
waitUtil = None

def open_browser(*arg):
    # 打开浏览器
    global driver, waitUtil
    try:
        driver = webdriver.Chrome(executable_path = firefoxDriverFilePath)
        # driver对象创建成果后，创建等待类实例对象
        waitUtil = WaitUtil(driver)
    except Exception as err:
        raise err

def visit_url(url, *arg):
    # 访问某个网址
    global driver
    try:
        driver.get(url)
    except Exception as err:
        raise err

def close_browser(*arg):
    # 关闭浏览器
    global driver
    try:
        driver.quit()
    except Exception as err:
        raise err

def sleep(sleepSeconds, *arg):
    # 强制等待
    try:
        time.sleep(int(sleepSeconds))
    except Exception as err:
        raise err

def clear(locationType, locatorExpression, *arg):
    # 清楚输入框默认内容
    global driver
    try:
        getElement(driver, locationType, locatorExpression).clear()
    except Exception as err:
        raise err

def input_string(locationType, locatorExpression, inputContent):
    # 在页面输入框中输入数据
    global driver
    try:
        getElement(driver, locationType,
            locatorExpression).send_keys(inputContent)
    except Exception as err:
        raise err

def click(locationType, locatorExpression, *arg):
    # 点击页面元素
    global driver
    try:
        getElement(driver, locationType, locatorExpression).click()
    except Exception as err:
        raise err

def assert_string_in_pagesource(assertString, *arg):
    # 断言页面源码是否存在某关键字或关键字符串
    global driver
    try:
        assert assertString in driver.page_source, \
            "%s not found in page source!" %assertString
    except AssertionError as err:
        raise AssertionError(err)
    except Exception as e:
        raise e

def assert_title(titleStr, *args):
    # 断言页面标题是否存在给定的关键字符串
    global driver
    try:
        assert titleStr in driver.title, \
            "%s not found in title!" %titleStr
    except AssertionError as err:
        raise AssertionError(err)
    except Exception as e:
        raise e

def getTitle(*arg):
    # 获取页面标题
    global driver
    try:
        return driver.title
    except Exception as err:
        raise err

def getPageSource(*arg):
    # 获取页面源码
    global driver
    try:
        return driver.page_source
    except Exception as err:
        raise err

def switch_to_frame(locationType, frameLocatorExpression, *arg):
    # 切换进入frame
    global driver
    try:
        driver.switch_to.frame(getElement
            (driver, locationType, frameLocatorExpression))
    except Exception as err:
        raise err

def switch_to_default_content(*arg):
    # 切出frame
    global driver
    try:
        driver.switch_to.default_content()
    except Exception as err:
        raise err

def paste_string(pasteString, *arg):
    # 模拟ctrl + v操作
    try:
        Clipboard.setText(pasteString)
        # 等待2秒，防止代码执行的太快，而未成功粘贴内容
        time.sleep(2)
        KeyboardKeys.twoKeys("ctrl", "v")
    except Exception as err:
        raise err

def press_tab_key(*arg):
    # 模拟 tab 键
    try:
        KeyboardKeys.oneKey("tab")
    except Exception as err:
        raise err

def press_enter_key(*arg):
    # 模拟 enter 键
    try:
        KeyboardKeys.oneKey("enter")
    except Exception as err:
        raise err

def maximize_browser():
    # 窗口最大化
    global driver
    try:
        driver.maximize_window()
    except Exception as err:
        raise err

def capture_screen(*args):
    # 截取屏幕图片
    global driver
    currTime = getCurrentTime()
    picNameAndPath = str(createCurrentDateDir()) + "/" + str(currTime) + ".png"
    try:
        driver.get_screenshot_as_file(picNameAndPath)
    except Exception as err:
        raise err
    else:
        return picNameAndPath

def waitPresenceOfElementLocated(locationType, locatorExpression, *arg):
    '''显示等待页面元素出现在DOM中，但并一定可以见，
            存在则返回该页面元素对象'''
    global waitUtil
    try:
        element = waitUtil.presenceOfElementLocated(locationType, locatorExpression)
        return element
    except Exception as err:
        raise err

def waitFrameToBeAvailableAndSwitchToIt(locationType, locatorExpression, *args):
    '''检查frame是否存在，存在则切换进frame控件中'''
    global waitUtil
    try:
        waitUtil.frameToBeAvailableAndSwitchToIt(locationType, locatorExpression)
    except Exception as err:
        raise err

def waitVisibilityOfElementLocated(locationType, locatorExpression, *args):
    '''显示等待页面元素出现在DOM中，并且可见，存在返回该页面元素对象'''
    global waitUtil
    try:
        element = waitUtil.visibilityOfElementLocated(locationType, locatorExpression)
        return element
    except Exception as err:
        raise err
