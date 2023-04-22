# encoding=utf-8
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC


class WaitUtil(object):
    def __init__(self, driver):
        self.locationTypeDict = {
            "xpath": By.XPATH,
            "id": By.ID,
            "name": By.NAME,
            "css_selector": By.CSS_SELECTOR,
            "class_name": By.CLASS_NAME,
            "tag_name": By.TAG_NAME,
            "link_text": By.LINK_TEXT,
            "partial_link_text": By.PARTIAL_LINK_TEXT,
        }
        self.driver = driver
        self.wait = WebDriverWait(driver, 60)

    def presenceOfElementLocated(self, locatorMethod, locatorExpression, *args):
        """显式等待页面元素出现在DOM中，但并不一定可见，存在则返回该页面元素"""
        try:
            # if self.locationTypeDict.has_key(locatorMethod.lower()):
            if locatorMethod.lower() in self.locationTypeDict:
                self.wait.until(
                    EC.presence_of_element_located((self.locationTypeDict[locatorMethod.lower()], locatorExpression)))
            else:
                raise Exception(u"未找到定位方式，请确认定位方式是否书写正确")
        except Exception as e:
            raise e

    def frameToBeAvailableAndSwitchToIt(self, locationType, locatorExpression, *args):
        """检查frame是否存在，存在则切换进frame控件中"""
        try:
            self.wait.until(EC.frame_to_be_available_and_switch_to_it(
                (self.locationTypeDict[locationType.lower()], locatorExpression)))
        except Exception as e:
            raise e

    def visibilityOfElementLocated(self, locationMethod, locatorExperssion, *args):
        """显式等待页面元素出现在DOM中，并且可见，存在则返回该页面元素对象"""
        try:
            self.wait.until(
                EC.visibility_of_element_located((self.locationTypeDict[locationMethod.lower()], locatorExperssion)))
        except Exception as e:
            raise e


if __name__ == "__main__":
    from selenium import webdriver

    driver = webdriver.Firefox(executable_path="e:\\geckodriver")
    driver.get("https://mail.126.com/")
    waitUtil = WaitUtil(driver)
    driver.find_element_by_id("lbNormal").click()
    # waitUtil.presenceOfElementLocated("id","lbNormal")
    waitUtil.frameToBeAvailableAndSwitchToIt("xpath", '//iframe[contains(@id,"x-URS-iframe")]')
    waitUtil.visibilityOfElementLocated("xpath", "//input[@name='email']")
    waitUtil.presenceOfElementLocated("xpath", "//input[@name='email']")
    driver.quit()
