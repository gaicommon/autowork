from selenium import webdriver


if __name__ =="__main__":
    driver = webdriver.Ie(r'./driver/IEDriverServer.exe')
    driver.get(r'http://www.baidu.com')

    print('打开成功')