from selenium.webdriver.support.ui import WebDriverWait

from ProjVar.var import IE64DriverFilePath


# 获取单个页面元素对象
def getElement(driver, localtorType, localtorExpression):
    try:
        element = WebDriverWait(driver, 60).until(lambda x: x.find_element(by=localtorType, value=localtorExpression))
        return element
    except Exception as e:
        raise e


# 获取多个页面元素对象
def getElements(driver, localtorType, localtorExpression):
    try:
        elements = WebDriverWait(driver, 60).until(lambda x: x.find_elements(by=localtorType, value=localtorExpression))
        return elements
    except Exception as e:
        raise e


if __name__ == "__main__":
    from selenium import webdriver

    # 进行单元测试
    driver = webdriver.Ie(executable_path=IE64DriverFilePath)
    driver.maximize_window()
    driver.get("https://mail.126.com/")
    lb = getElement(driver, "id", "lbNormal")
    print(lb)
    driver.quit()
