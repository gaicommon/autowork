# # encoding=utf-8
# from selenium import webdriver
# from ProjVar.var import *
# from Util.find import *
# from Util.KeyBoardUtil import KeyBoardKeys
# from Util.ClipboardUtil import Clipboard
# from Util.waitutil import WaitUtil
# from selenium.webdriver.firefox.options import Options
# from selenium.webdriver.chrome.options import Options
# import time
# from Util.fromattime import *
#
# # 定义全局变量driver
# driver = None
# # 定义全局的等待类实例对象
# waitUtil = None
#
#
# def open(browserName):
#     # 打开浏览器
#     global driver, waitUtil
#     try:
#         if browserName.lower() == "ie":
#             driver = webdriver.Ie(executable_path=ieDriverFilePath)
#         elif browserName.lower == "chrome":
#             # 创建Chrome浏览器的一个Options实例对象
#             chrome_options = Options()
#             # 添加屏蔽--ignore--certificate--errors提示信息的设置参数项
#             chrome_options.add_experimental_option("excludeSwitches", ["ignore-certificate-errors"])
#             driver = webdriver.Chrome(executable_path=chromeDriverFilePath, chrome_options=chrome_options)
#         else:
#             driver = webdriver.Firefox(executable_path=firefoxDriverFilePath)
#         # driver对象创建成功后，创建等待类实例对象
#         waitUtil = WaitUtil(driver)
#     except Exception as e:
#         raise e
#
#
# def visit(url):
#     # 访问某个网站
#     global driver
#     try:
#         driver.get(url)
#     except Exception as e:
#         raise e
#
#
# def close_browser():
#     # 关闭浏览器
#     global driver
#     try:
#         driver.quit()
#     except Exception as e:
#         raise e
#
#
# def sleep(sleepSeconds):
#     # 强制等待
#     try:
#         time.sleep(int(sleepSeconds))
#     except Exception as e:
#         raise e
#
#
# def clear(locationType, locatorExpression):
#     # 清空输入框默认内容
#     global driver
#     try:
#         getElement(driver, locationType, locatorExpression).clear()
#     except Exception as e:
#         raise e
#
#
# def input_string(locationType, locatorExpression, inputContent):
#     # 在页面输入框中输入数据
#     global driver
#     try:
#         getElement(driver, locationType, locatorExpression).send_keys(inputContent)
#     except Exception as e:
#         raise e
#
#
# def click(locationType, locatorExpression, *args):
#     # 点击页面元素
#     global driver
#     try:
#         getElement(driver, locationType, locatorExpression).click()
#     except Exception as e:
#         raise e
#
#
# def assert_string_in_pagesource(assertString, *args):
#     # 断言页面源码是否存在某个关键字或关键字符串
#     global driver
#     try:
#         assert assertString in driver.page_source, u"%s not found in page source!" % assertString
#     except AssertionError as e:
#         raise AssertionError(e)
#     except Exception as e:
#         raise e
#
#
# def assert_title(titleStr, *args):
#     # 断言页面标题是否存在给定的关键字符串
#     global driver
#     try:
#         assert titleStr in driver.title, u"%s not found in page title!" % titleStr
#     except AssertionError as e:
#         raise AssertionError(e)
#     except Exception as e:
#         raise e
#
#
# def getTitle(*args):
#     # 获取页面标题
#     global driver
#     try:
#         return driver.title
#     except Exception as e:
#         raise e
#
#
# def getPageSource(*args):
#     # 获取页面源码
#     global driver
#     try:
#         return driver.page_source
#     except Exception as e:
#         raise e
#
#
# def switch_to_frame(locationType, frameLocatorExpressoin, *args):
#     # 切换进frame
#     global driver
#     try:
#         driver.switch_to.frame(getElement(driver, locationType, frameLocatorExpressoin))
#     except Exception as e:
#         print("frame error!")
#         raise e
#
#
# def switch_to_default_content(*args):
#     # 切换妯frame
#     global driver
#     try:
#         driver.switch_to.default_content()
#     except Exception as e:
#         raise e
#
#
# def paste_string(pasteString, *args):
#     # 模拟Ctrl+V操作
#     try:
#         Clipboard.setText(pasteString)
#         # 等待2秒，防止代码执行过快，而未成功粘贴内容
#         time.sleep(2)
#         KeyBoardKeys.twoKeys("ctrl", "v")
#     except Exception as e:
#         raise e
#
#
# def press_tab_key(*args):
#     # 模拟tab键
#     try:
#         KeyBoardKeys.oneKey("tab")
#     except Exception as e:
#         raise e
#
#
# def press_enter_key(*args):
#     # 模拟enter键
#     try:
#         KeyBoardKeys.oneKey("enter")
#     except Exception as e:
#         raise e
#
#
# def maximize(*args):
#     # 窗口最大化
#     global driver
#     try:
#         driver.maximize_window()
#     except Exception as e:
#         raise e
#
#
# def capture(file_path):
#     try:
#         driver.save_screenshot(file_path)
#     except Exception as e:
#         raise e
#
#
# def waitPresenceOfElementLocated(locationType, locatorExpression, *args):
#     """显式等待页面元素出现在DOM中，但不一定可见，存在则返回该页面元素对象"""
#     global waitUtil
#     try:
#         waitUtil.presenceOfElementLocated(locationType, locatorExpression)
#     except Exception as e:
#         raise e
#
#
# def waitFrameToBeAvailableAndSwitchToIt(locationType, locatorExprssion, *args):
#     """检查frame是否存在，存在则切换进frame控件中"""
#     global waitUtil
#     try:
#         waitUtil.frameToBeAvailableAndSwitchToIt(locationType, locatorExprssion)
#     except Exception as e:
#         raise e
#
#
# def waitVisibilityOfElementLocated(locationType, locatorExpression, *args):
#     """显式等待页面元素出现在Dom中，并且可见，存在返回该页面元素对象"""
#     global waitUtil
#     try:
#         waitUtil.visibilityOfElementLocated(locationType, locatorExpression)
#     except Exception as e:
#         raise e
