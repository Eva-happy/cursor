# 导入必要的库
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
from bs4 import BeautifulSoup

def get_region_options(driver, level="province"):
    """获取地区选项"""
    try:
        # 等待页面加载
        time.sleep(5)
        
        options = []
        if level == "province":
            # 获取省份列表 - 使用class='f66 fsize14'
            elements = driver.find_elements(By.CSS_SELECTOR, 'a.f66.fsize14')
            for element in elements:
                text = element.text.strip()
                if text and text != '省份':
                    options.append(text)
        
        elif level == "city":
            # 获取城市列表
            try:
                # 点击"信息公开"
                print("点击信息公开...")
                info_btn = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), '信息公开')]"))
                )
                info_btn.click()
                time.sleep(2)
                
                # 点击"电价标准"
                print("点击电价标准...")
                price_btn = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), '电价标准')]"))
                )
                price_btn.click()
                time.sleep(2)
                
                # 等待城市选择弹出框
                print("等待城市选择弹出框...")
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div.el-dialog__body'))
                )
                time.sleep(2)
                
                # 获取城市列表
                city_elements = driver.find_elements(By.CSS_SELECTOR, 'ul.tab-con-box-ul li.tab-con-box-li')
                print(f"找到 {len(city_elements)} 个城市元素")
                
                for element in city_elements:
                    text = element.text.strip()
                    if text and text != '请选择':
                        options.append(text)
                        print(f"添加城市: {text}")
                
            except Exception as e:
                print(f"获取城市列表失败: {str(e)}")
                print("\n调试信息 - 当前页面HTML:")
                print(driver.page_source[:2000])
        
        else:  # district
            # 获取区县列表
            try:
                # 等待区县列表出现
                print("等待区县列表...")
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div.tab-con-box'))
                )
                time.sleep(2)
                
                # 获取区县列表
                district_elements = driver.find_elements(By.CSS_SELECTOR, 'div.tab-con-box li.tab-con-box-li')
                for element in district_elements:
                    text = element.text.strip()
                    if text and text != '请选择':
                        options.append(text)
                        print(f"添加区县: {text}")
                
            except Exception as e:
                print(f"获取区县列表失败: {str(e)}")
        
        if not options:
            print(f"警告: 未找到任何{level}选项")
        
        return options
    except Exception as e:
        print(f"获取{level}选项失败: {str(e)}")
        return []

def extract_region_data(driver, province_name=None):
    """提取地区数据的递归函数"""
    try:
        html_content = driver.page_source
        soup = BeautifulSoup(html_content, 'html.parser')
        
        if not province_name:
            # 获取所有省份
            provinces = {}
            province_li = soup.find('li', {'data-v-07831be2': True})
            if province_li:
                for element in province_li.find_all('i', {'data-v-07831be2': True, 'class': 'io'}):
                    name = element.text.strip()
                    if name and name != '省份':
                        provinces[name] = []
            return provinces
        else:
            # 获取指定省份的城市和区县
            province_data = {}
            cities = []
            
            # 获取城市列表
            city_li = soup.find_all('li', {'data-v-07831be2': True})[1] if len(soup.find_all('li', {'data-v-07831be2': True})) > 1 else None
            if city_li:
                for city_element in city_li.find_all('i', {'data-v-07831be2': True, 'class': 'io'}):
                    city_name = city_element.text.strip()
                    if city_name and city_name != '市':
                        city_data = {'city': city_name, 'districts': []}
                        
                        # 获取区县列表
                        district_li = soup.find_all('li', {'data-v-07831be2': True})[2] if len(soup.find_all('li', {'data-v-07831be2': True})) > 2 else None
                        if district_li:
                            for district_element in district_li.find_all('i', {'data-v-07831be2': True, 'class': 'io'}):
                                district_name = district_element.text.strip()
                                if district_name and district_name != '县/区':
                                    city_data['districts'].append(district_name)
                        
                        cities.append(city_data)
            
            province_data[province_name] = cities
            return province_data
    except Exception as e:
        print(f"提取地区数据失败: {str(e)}")
        return {}

def click_element(driver, element):
    """点击元素的通用函数"""
    try:
        # 确保元素在视图中
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
        time.sleep(0.5)
        
        # 尝试点击
        try:
            if element.is_displayed() and element.is_enabled():
                element.click()
            else:
                driver.execute_script("arguments[0].click();", element)
        except:
            driver.execute_script("arguments[0].click();", element)
        return True
    except Exception as e:
        print(f"点击元素失败: {str(e)}")
        return False

def select_region_element(driver, text, level):
    """选择地区元素"""
    try:
        print(f"\n尝试选择{level}: {text}")
        
        # 等待页面更新
        time.sleep(2)
        
        # 根据不同级别使用不同的选择器
        if level == "province":
            # 选择省份
            selectors = [
                f"//a[contains(@class, 'f66 fsize14') and contains(text(), '{text}')]",
                f"//a[contains(@class, 'f66 fsize14') and .//text()='{text}']"
            ]
        else:  # city or district
            # 选择城市或区县
            selectors = [
                f"//li[@data-v-40fe627e]//i[contains(@class, 'io') and contains(text(), '{text}')]",
                f"//li[@data-v-40fe627e]//i[contains(@class, 'io') and .//text()='{text}']"
            ]
        
        for selector in selectors:
            try:
                print(f"尝试选择器: {selector}")
                element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, selector))
                )
                if element.is_displayed():
                    print(f"找到可见元素: {selector}")
                    # 如果是城市或区县，需要点击整个li元素
                    if level != "province":
                        element = element.find_element(By.XPATH, "./ancestor::li[@data-v-40fe627e]")
                    return element
                else:
                    print(f"元素不可见: {selector}")
            except Exception as e:
                print(f"选择器失败: {selector}")
                print(f"错误: {str(e)}")
                continue
        
        print(f"\n警告: 未找到{level}元素 {text}")
        return None
    except Exception as e:
        print(f"选择{level}元素失败: {str(e)}")
        return None

def navigate_to_region(driver, text, level):
    """导航到指定地区"""
    try:
        # 选择并点击地区元素
        element = select_region_element(driver, text, level)
        if element:
            # 尝试点击元素
            try:
                element.click()
                print(f"成功点击: {text}")
                # 等待元素状态更新（从io变为it）
                if level != "district":  # 区县不需要等待状态更新
                    try:
                        WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.XPATH, f"//i[contains(@class, 'it') and contains(text(), '{text}')]"))
                        )
                    except:
                        pass
                time.sleep(3)  # 等待页面加载
                return True
            except:
                # 如果直接点击失败，尝试使用JavaScript点击
                try:
                    driver.execute_script("arguments[0].click();", element)
                    print(f"使用JavaScript点击: {text}")
                    time.sleep(3)  # 等待页面加载
                    return True
                except Exception as e:
                    print(f"JavaScript点击失败: {str(e)}")
                    return False
        return False
    except Exception as e:
        print(f"导航失败: {str(e)}")
        return False

def wait_for_page_update(driver, province):
    """等待页面更新到指定省份"""
    try:
        max_retries = 3
        for retry in range(max_retries):
            try:
                # 等待页面标题或URL更新
                WebDriverWait(driver, 10).until(
                    lambda d: province in d.title or province in d.current_url
                )
                
                # 等待加载动画消失（如果有的话）
                try:
                    WebDriverWait(driver, 5).until_not(
                        EC.presence_of_element_located((By.CSS_SELECTOR, ".loading, .spinner"))
                    )
                except:
                    pass
                
                # 等待新内容加载
                time.sleep(2)
                
                # 验证页面是否包含省份信息
                if province in driver.page_source:
                    print(f"页面已包含 {province} 的信息")
                    return True
                else:
                    print(f"页面未包含 {province} 的信息，重试中...")
                    time.sleep(2)
            except Exception as e:
                print(f"等待页面更新失败，重试 {retry + 1}/{max_retries}: {str(e)}")
                time.sleep(2)
        
        print("页面更新重试次数已达上限")
        return False
    except Exception as e:
        print(f"等待页面更新失败: {str(e)}")
        return False

def is_direct_municipality(province):
    """判断是否为直辖市"""
    return province in ["北京", "上海", "天津", "重庆"]

def main():
    # 获取脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    driver_path = os.path.join(script_dir, "msedgedriver.exe")
    
    # 检查驱动程序
    if not os.path.exists(driver_path):
        print(f"错误: msedgedriver.exe 不存在于 {driver_path}")
        return
    
    # 设置浏览器选项
    edge_options = Options()
    edge_options.add_argument("--start-maximized")
    edge_options.add_argument('--disable-gpu')
    edge_options.add_argument('--no-sandbox')
    edge_options.add_argument('--disable-dev-shm-usage')
    edge_options.add_argument('--ignore-certificate-errors')
    
    try:
        # 初始化浏览器
        service = Service(driver_path)
        driver = webdriver.Edge(service=service, options=edge_options)
        
        # 访问网站
        print("正在访问网站...")
        driver.get("https://www.95598.cn/osgweb/index")
        print("等待页面加载...")
        time.sleep(5)
        
        # 点击地区选择器
        print("\n尝试点击地区选择器...")
        region_selectors = [
            "//div[@id='city_select']//a[contains(@class, 'current')]",
            "//div[contains(@class, 'region')]//a[contains(@class, 'current')]",
            "//a[contains(@class, 'current fsize16')]",
            "//div[@data-v-07831be2]//a[contains(@class, 'current')]"
        ]
        
        region_element = None
        for selector in region_selectors:
            try:
                print(f"尝试选择器: {selector}")
                region_element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, selector))
                )
                if region_element.is_displayed():
                    print(f"找到可见元素: {selector}")
                    break
                else:
                    print(f"元素不可见: {selector}")
            except Exception as e:
                print(f"选择器失败: {selector}")
                continue
        
        if region_element:
            try:
                region_element.click()
                print("成功点击地区选择器")
            except:
                try:
                    driver.execute_script("arguments[0].click();", region_element)
                    print("使用JavaScript点击地区选择器")
                except Exception as e:
                    print(f"点击地区选择器失败: {str(e)}")
                    return
        else:
            print("未找到地区选择器")
            return
        
        time.sleep(2)
        
        # 获取省份列表
        print("\n正在获取省份列表...")
        provinces = get_region_options(driver, "province")
        
        if not provinces:
            print("错误: 未找到任何省份")
            print("\n当前页面源码:")
            print(driver.page_source[:1000])
            return
        
        # 显示省份列表
        print("\n请选择一个省份：")
        for i, province in enumerate(provinces, 1):
            print(f"{i}. {province}")
        
        # 用户选择省份
        while True:
            try:
                choice = int(input("请输入省份的数字："))
                if 1 <= choice <= len(provinces):
                    selected_province = provinces[choice-1]
                    print(f"您选择的省份是：{selected_province}")
                    
                    # 导航到选中的省份
                    if not navigate_to_region(driver, selected_province, "province"):
                        print("导航到省份失败")
                        return
                    
                    # 等待省份页面加载
                    print("等待省份页面加载...")
                    time.sleep(5)
                    
                    # 点击"信息公开"链接
                    print("点击信息公开链接...")
                    try:
                        # 使用更精确的选择器
                        info_selectors = [
                            "//li[@data-v-4781518a and @class='li_lis']//span[@id='right-class' and contains(text(), '信息公开')]",
                            "//span[@id='right-class' and @class='right-class span' and contains(text(), '信息公开')]",
                            "//li[@data-v-4781518a]//span[contains(text(), '信息公开')]",
                            "//li[@class='li_lis']//span[contains(text(), '信息公开')]"
                        ]
                        
                        info_btn = None
                        for selector in info_selectors:
                            try:
                                print(f"尝试信息公开选择器: {selector}")
                                info_btn = WebDriverWait(driver, 5).until(
                                    EC.element_to_be_clickable((By.XPATH, selector))
                                )
                                if info_btn.is_displayed():
                                    print(f"找到可见的信息公开链接: {selector}")
                                    break
                            except:
                                continue
                        
                        if info_btn:
                            try:
                                # 滚动到元素位置
                                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", info_btn)
                                time.sleep(1)
                                
                                # 尝试直接点击
                                info_btn.click()
                                print("成功点击信息公开链接")
                            except:
                                # 如果直接点击失败，尝试使用JavaScript点击
                                try:
                                    driver.execute_script("arguments[0].click();", info_btn)
                                    print("使用JavaScript点击信息公开链接")
                                except Exception as e:
                                    print(f"点击信息公开链接失败: {str(e)}")
                                    return
                        else:
                            print("未找到信息公开链接")
                            # 打印当前页面源码以便调试
                            print("\n当前页面HTML:")
                            print(driver.page_source[:2000])
                            return
                        
                        time.sleep(3)
                        
                        # 点击"电价标准"链接
                        print("点击电价标准链接...")
                        price_selectors = [
                            "//a[text()='电价标准']",
                            "//div[contains(@class, 'submenu')]//a[text()='电价标准']",
                            "//ul[contains(@class, 'submenu')]//a[text()='电价标准']",
                            "//div[contains(@class, 'nav')]//a[text()='电价标准']"
                        ]
                        
                        price_btn = None
                        for selector in price_selectors:
                            try:
                                print(f"尝试电价标准选择器: {selector}")
                                price_btn = WebDriverWait(driver, 5).until(
                                    EC.element_to_be_clickable((By.XPATH, selector))
                                )
                                if price_btn.is_displayed():
                                    print(f"找到可见的电价标准链接: {selector}")
                                    break
                            except:
                                continue
                        
                        if price_btn:
                            try:
                                # 滚动到元素位置
                                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", price_btn)
                                time.sleep(1)
                                
                                # 尝试直接点击
                                price_btn.click()
                                print("成功点击电价标准链接")
                            except:
                                # 如果直接点击失败，尝试使用JavaScript点击
                                try:
                                    driver.execute_script("arguments[0].click();", price_btn)
                                    print("使用JavaScript点击电价标准链接")
                                except Exception as e:
                                    print(f"点击电价标准链接失败: {str(e)}")
                                    return
                        else:
                            print("未找到电价标准链接")
                            return
                        
                        time.sleep(3)
                        
                        # 等待城市选择弹出框
                        print("等待城市选择弹出框...")
                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, 'div.el-dialog__body'))
                        )
                        
                        # 获取城市列表
                        city_elements = driver.find_elements(By.CSS_SELECTOR, 'ul.tab-con-box-ul li.tab-con-box-li')
                        cities = []
                        for element in city_elements:
                            text = element.text.strip()
                            if text and text != '请选择':
                                cities.append(text)
                                print(f"找到城市: {text}")
                        
                        if not cities:
                            print("未找到任何城市")
                            return
                        
                        # 显示城市列表供用户选择
                        print("\n请选择城市：")
                        for i, city in enumerate(cities, 1):
                            print(f"{i}. {city}")
                        
                        # 用户选择城市
                        while True:
                            try:
                                choice = int(input("\n请输入城市编号: "))
                                if 1 <= choice <= len(cities):
                                    selected_city = cities[choice-1]
                                    print(f"您选择的城市是：{selected_city}")
                                    
                                    # 点击选中的城市
                                    try:
                                        city_element = driver.find_element(By.XPATH, f"//li[contains(@class, 'tab-con-box-li') and contains(text(), '{selected_city}')]")
                                        city_element.click()
                                        print(f"成功点击城市: {selected_city}")
                                        time.sleep(2)
                                        
                                        # 等待区县列表加载
                                        print("等待区县列表加载...")
                                        WebDriverWait(driver, 10).until(
                                            EC.presence_of_element_located((By.CSS_SELECTOR, 'ul.tab-con-box-ul li.tab-con-box-li'))
                                        )
                                        
                                        # 获取区县列表
                                        district_elements = driver.find_elements(By.CSS_SELECTOR, 'ul.tab-con-box-ul li.tab-con-box-li')
                                        districts = []
                                        for element in district_elements:
                                            text = element.text.strip()
                                            if text and text != '请选择':
                                                districts.append(text)
                                                print(f"找到区县: {text}")
                                        
                                        if not districts:
                                            print("未找到任何区县")
                                            return
                                        
                                        # 显示区县列表供用户选择
                                        print(f"\n{selected_city}的区县列表：")
                                        for i, district in enumerate(districts, 1):
                                            print(f"{i}. {district}")
                                        
                                        # 用户选择区县
                                        while True:
                                            try:
                                                choice = int(input("\n请输入区县编号: "))
                                                if 1 <= choice <= len(districts):
                                                    selected_district = districts[choice-1]
                                                    print(f"您选择的区县是：{selected_district}")
                                                    
                                                    # 点击选中的区县
                                                    district_element = driver.find_element(By.XPATH, f"//li[contains(@class, 'tab-con-box-li') and contains(text(), '{selected_district}')]")
                                                    district_element.click()
                                                    print(f"成功点击区县: {selected_district}")
                                                    time.sleep(2)
                                                    
                                                    # 点击代理购电选项展开月份列表
                                                    print("点击代理购电选项...")
                                                    proxy_selectors = [
                                                        # 尝试点击展开箭头
                                                        "//div[text()='代理购电']/parent::div//i[contains(@class, 'el-icon-arrow-right')]",
                                                        "//div[text()='代理购电']/parent::div//div[contains(@class, 'el-table__expand-icon')]",
                                                        "//div[text()='代理购电']/parent::div",
                                                        "//div[text()='代理购电']"
                                                    ]
                                                    
                                                    # 打印当前页面源码以便调试
                                                    print("\n当前页面HTML:")
                                                    page_source = driver.page_source
                                                    print(page_source[:2000])
                                                    
                                                    # 检查页面中是否包含"代理购电"文本
                                                    if "代理购电" in page_source:
                                                        print("页面中包含'代理购电'文本")
                                                    else:
                                                        print("警告: 页面中未找到'代理购电'文本")
                                                    
                                                    proxy_btn = None
                                                    for selector in proxy_selectors:
                                                        try:
                                                            print(f"尝试代理购电选择器: {selector}")
                                                            elements = driver.find_elements(By.XPATH, selector)
                                                            print(f"找到 {len(elements)} 个匹配元素")
                                                            
                                                            for element in elements:
                                                                try:
                                                                    if element.is_displayed():
                                                                        text = element.text.strip()
                                                                        html = element.get_attribute('outerHTML')
                                                                        print(f"找到可见元素:")
                                                                        print(f"- 文本: {text}")
                                                                        print(f"- HTML: {html}")
                                                                        proxy_btn = element
                                                                        break
                                                                except:
                                                                    continue
                                                            
                                                            if proxy_btn:
                                                                break
                                                                
                                                        except Exception as e:
                                                            print(f"选择器 {selector} 失败: {str(e)}")
                                                            continue
                                                    
                                                    if proxy_btn:
                                                        try:
                                                            # 等待元素可点击
                                                            WebDriverWait(driver, 10).until(
                                                                EC.element_to_be_clickable((By.XPATH, "//div[text()='代理购电']/parent::div//i[contains(@class, 'el-icon-arrow-right')]"))
                                                            )
                                                            
                                                            # 滚动到元素位置
                                                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", proxy_btn)
                                                            time.sleep(2)  # 等待滚动完成
                                                            
                                                            # 尝试使用JavaScript点击
                                                            driver.execute_script("arguments[0].click();", proxy_btn)
                                                            print("成功点击代理购电选项")
                                                            
                                                            # 等待展开动画完成
                                                            time.sleep(3)
                                                            print("\n获取月份列表...")
                                                            
                                                            # 获取展开后的月份列表
                                                            month_elements = driver.find_elements(By.XPATH, "//div[contains(@class, 'el-table__expanded-cell')]//div[contains(text(), '年') or contains(text(), '月')]")
                                                            if month_elements:
                                                                print("\n找到以下月份：")
                                                                for i, element in enumerate(month_elements, 1):
                                                                    print(f"{i}. {element.text}")
                                                            else:
                                                                print("未找到月份信息")
                                                                
                                                        except Exception as e:
                                                            print(f"点击代理购电选项失败: {str(e)}")
                                                            return
                                                    else:
                                                        print("未找到代理购电选项")
                                                        return
                                                    
                                                    break
                                                else:
                                                    print("无效的选择，请重试")
                                            except ValueError:
                                                print("请输入有效的数字")
                                        
                                    except Exception as e:
                                        print(f"操作失败: {str(e)}")
                                        return
                                    
                                    break
                                else:
                                    print("无效的选择，请重试")
                            except ValueError:
                                print("请输入有效的数字")
                        
                    except Exception as e:
                        print(f"操作失败: {str(e)}")
                        # 打印当前页面源码以便调试
                        print("\n当前页面HTML:")
                        print(driver.page_source[:2000])
                        return
                    
                    break
                else:
                    print("无效的选择，请重试")
            except ValueError:
                print("请输入有效的数字")
        
    except Exception as e:
        print(f"发生错误: {str(e)}")
        import traceback
        print(traceback.format_exc())
    finally:
        try:
            driver.quit()
        except:
            pass

if __name__ == "__main__":
    print("=== 开始运行自动化程序 ===")
    main() 