from dataclasses import dataclass

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait


TimeOut = 180

Error = {

    "Wait_Token_Error": {
        'message': '停在等待授权页面，请授权AED后重新启动创建DRCS程序'
    },
    "Internet_Error": {
        'message': '网络无法链接，请核实网络，外网用户请核实是否已经登陆VPN'
    },
    "Page_load_Error": {
        'message': f'页面加载超过了程序设定的时间{TimeOut}秒,请检查网络.'
    }

}

# 选择frame查询列表中的值
@dataclass
class Frame_select_attr:
    btn_click:list=None
    wait_load_check_id:str=None
    iframe_id:str=None
    text_key:list=None
    el_list_xpath:str=None
    btn_lookup:list=None
    btn_confirm:list=None


# 选择Select中的选项
def select_by_visible_text(driver,el_str,vs_text,by=By.ID):
    select_el = Select(driver.find_element(by,el_str))
    #选择规定的选项
    select_el.select_by_visible_text(vs_text)



# 等待加载
def wait_el_load(driver,el_str:str,by=By.ID):
    '''
    用于检查指定元素是否已加载完成，即加载至所需要的页面，若加载完成，返回制定的页面，若未加载完成，则返回False
    '''
    try:
        # 等待
        return WebDriverWait(driver, timeout=TimeOut,poll_frequency=0.2).until(lambda x: x.find_element(by,el_str))
    except:
        return  False

# 通过标签文本内容查找列表的所需要的元素
def find_el_click(driver,xpath_str,text_value,attr_name="textContent",is_only=False,wait_time=30):
    '''在元素清单中查找属性名为attr_name的制与所提供制need_value匹配的元素，并点击选择,成功返回True，不成功返回False'''
    wait_times = 0
    while(wait_times<wait_time*10):
        el_list = driver.find_elements(By.XPATH,xpath_str)
        if el_list and (not is_only or len(el_list)==1 ):
            for i in el_list:
                try:
                    name = i.get_attribute(attr_name)
                    if name ==text_value:
                        i.click()
                        return True
                except:
                    wait_times+=1
                    continue
        wait_times+=1
        driver.implicitly_wait(0.1)
    return False
# 核查点击后页面是否关闭
def confirm_click_close(driver,el_str,by=By.ID,weit_time=30):
    driver.find_element(by,el_str).click()
    i=0
    while(i<weit_time*10):
        try:
            driver.find_element(by,el_str).click()
            driver.implicitly_wait(0.1)
            # if not driver.find_elements(by,el_str):
            #     return True
            # i+=1
        except:
            return True
    return False

def get_frame_select(driver,frame_select:Frame_select_attr):
    '''打开选择列表frame，并选择相关选项'''
    # 点击注页面的添加按钮
    driver.find_element(frame_select.btn_click[0],frame_select.btn_click[1]).click()
    # 等待 选择项目的iframe加载
    load_result = wait_el_load(driver,frame_select.wait_load_check_id)
    if load_result:
        # 将driver转到frame
        driver.switch_to.frame(frame_select.iframe_id)
        # 将查询值，写入查询input框
        driver.find_element(frame_select.text_key[0],frame_select.text_key[1]).send_keys(frame_select.text_key[2])
        # 点击查询按钮
        driver.find_element(frame_select.btn_lookup[0],frame_select.btn_lookup[1]).click()
        # 查找值并选择
        find_el_click(driver,frame_select.el_list_xpath,frame_select.text_key[2])
        # 转会主页面
        driver.switch_to.parent_frame()
        # 点击确认按钮
        return confirm_click_close(driver,frame_select.btn_confirm[1],frame_select.btn_confirm[0])
    else:
        return load_result


def get_frame_input(driver,frame_select:Frame_select_attr):
    '''打开选择列表frame，并选择相关选项'''
    # 点击注页面的添加按钮
    driver.find_element(frame_select.btn_click[0],frame_select.btn_click[1]).click()
    # 等待 选择项目的iframe加载
    wait_el_load(driver,frame_select.wait_load_check_id)
    # 将driver转到frame
    driver.switch_to.frame(frame_select.iframe_id)
    # 将值写入input框
    driver.find_element(frame_select.text_key[0],frame_select.text_key[1]).send_keys(frame_select.text_key[2])
    # 转会主页面
    driver.switch_to.parent_frame()
    # 点击确认按钮
    return confirm_click_close(driver,frame_select.btn_confirm[1],frame_select.btn_confirm[0])

# 等待元素加载并点击
def wait_el_and_click(driver,el_str:str,by=By.ID,is_close=True):
    load_result=wait_el_load(driver,el_str,by)
    if load_result:
        if is_close:
            cloase_result= confirm_click_close(driver,el_str,by)
            return cloase_result
        el = driver.find_element(by,el_str).click()
        return True
    return False

def get_value_list_by_dickey(dic_list:list,key:str):
    if dic_list:
        result_list = []
        for data in dic_list:
            result_list.append(data[key].strip())
        return result_list
