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
from selenium.webdriver.common.keys import Keys

def setup_driver():
    """设置并返回WebDriver"""
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        driver_path = os.path.join(script_dir, "msedgedriver.exe")
        
        if not os.path.exists(driver_path):
            print(f"错误: msedgedriver.exe 不存在于 {driver_path}")
            return None
        
        edge_options = Options()
        edge_options.add_argument("--start-maximized")
        edge_options.add_argument('--disable-gpu')
        edge_options.add_argument('--no-sandbox')
        edge_options.add_argument('--disable-dev-shm-usage')
        edge_options.add_argument('--ignore-certificate-errors')
        
        service = Service(driver_path)
        driver = webdriver.Edge(service=service, options=edge_options)
        return driver
    except Exception as e:
        print(f"设置驱动程序失败: {str(e)}")
        return None

def wait_and_click_element(driver, element, use_js=False):
    """等待并点击元素"""
    try:
        element_text = "选择地区" if element.text.strip() == "上海" else element.text.strip()
        print(f"\n正在点击: {element_text}")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
        time.sleep(0.5)  # 减少等待时间
        
        if not use_js:
            element.click()
            print(f"成功点击: {element_text}")
        else:
            driver.execute_script("arguments[0].click();", element)
            print(f"使用JS点击: {element_text}")
        return True
    except Exception as e:
        print(f"点击元素失败: {str(e)}")
        return False

def find_element_with_retry(driver, selectors, timeout=5):
    """使用多个选择器尝试查找元素"""
    for selector in selectors:
        try:
            element = WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable((By.XPATH, selector))
            )
            if element.is_displayed():
                return element
        except:
            continue
    return None

def get_visible_elements(driver, selector, timeout=3):  # 减少默认等待时间
    """获取所有可见的元素"""
    try:
        elements = WebDriverWait(driver, timeout).until(
            EC.presence_of_all_elements_located((By.XPATH, selector))
        )
        return [e for e in elements if e.is_displayed()]
    except:
        return []

def extract_element_text(element):
    """提取元素文本"""
    try:
        text = element.text.strip()
        if element.tag_name == 'tr':
            cells = element.find_elements(By.CSS_SELECTOR, 'div.cell')
            cell_texts = [cell.text.strip() for cell in cells if cell.text.strip()]
            if cell_texts:
                text = ' | '.join(cell_texts)
        return text
    except:
        return ""

def get_clickable_element(element):
    """获取可点击的元素"""
    try:
        return element.find_element(By.TAG_NAME, "a")
    except:
        return element

def navigate_to_price_page(driver):
    """导航到电价标准页面"""
    try:
        # 点击信息公开
        info_selectors = [
            "//li[@data-v-4781518a and @class='li_lis']//span[@id='right-class' and contains(text(), '信息公开')]",
            "//span[@id='right-class' and @class='right-class span' and contains(text(), '信息公开')]",
            "//li[@data-v-4781518a]//span[contains(text(), '信息公开')]"
        ]
        info_btn = find_element_with_retry(driver, info_selectors)
        if not info_btn or not wait_and_click_element(driver, info_btn):
            return False
        
        time.sleep(1)
        
        # 点击电价标准
        price_selectors = [
            "//a[text()='电价标准']",
            "//div[contains(@class, 'submenu')]//a[text()='电价标准']",
            "//ul[contains(@class, 'submenu')]//a[text()='电价标准']"
        ]
        price_btn = find_element_with_retry(driver, price_selectors)
        if not price_btn or not wait_and_click_element(driver, price_btn):
            return False
        
        return True
    except Exception as e:
        print(f"导航到电价标准页面失败: {str(e)}")
        return False

def wait_for_city_dialog(driver):
    """等待城市选择对话框"""
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div.el-dialog__body'))
        )
        time.sleep(1)  # 减少等待时间
        return True
    except Exception as e:
        print(f"等待城市选择框超时: {str(e)}")
        return False

def get_cities(driver):
    """获取城市列表"""
    city_selectors = [
        "//div[contains(@class, 'tab-con-box')]//li[contains(@class, 'tab-con-box-li')]",
        "//ul[contains(@class, 'tab-con-box-ul')]//li",
        "//div[contains(@class, 'el-dialog__body')]//li[contains(@class, 'tab-con-box-li')]"
    ]
    
    cities = []
    for selector in city_selectors:
        elements = get_visible_elements(driver, selector)
        for element in elements:
            text = extract_element_text(element)
            if text and text != '请选择' and text not in [c['text'] for c in cities]:
                cities.append({"text": text, "element": element})
            
    return cities

def get_districts(driver):
    """获取区县列表"""
    try:
        print("\n正在加载区县列表，请稍候...")
        
        # 使用更短的等待时间
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'ul.tab-con-box-ul li.tab-con-box-li'))
        )
        
        districts = []
        elements = driver.find_elements(By.CSS_SELECTOR, 'ul.tab-con-box-ul li.tab-con-box-li')
        for element in elements:
            text = extract_element_text(element)
            if text and text != '请选择':
                districts.append({"text": text, "element": element})
        
        if districts:
            print(f"\n成功获取到 {len(districts)} 个区县")
        else:
            print("\n未找到任何区县信息")
        
        return districts
    except Exception as e:
        print(f"\n获取区县列表失败: {str(e)}")
        return []

def get_projects(driver):
    """获取项目列表"""
    try:
        print("\n正在获取电价标准项目列表...")
        
        # 等待表格加载
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '.el-table__body-wrapper'))
        )
        time.sleep(1)  # 等待表格内容完全加载
        
        # 获取所有表格行
        rows = driver.find_elements(By.CSS_SELECTOR, '.el-table__body-wrapper .el-table__row')
        
        projects = []
        seen_texts = set()
        
        for row in rows:
            try:
                # 获取行中的所有单元格
                cells = row.find_elements(By.CSS_SELECTOR, '.cell')
                
                # 获取第一个单元格的文本（通常是项目名称）
                if cells:
                    project_text = cells[0].text.strip()
                    if project_text and project_text not in seen_texts:
                        seen_texts.add(project_text)
                        # 使用整个行作为元素，因为展开箭头在行级别
                        projects.append({"text": project_text, "element": row})
            except:
                continue
        
        if projects:
            print(f"\n成功获取到 {len(projects)} 个项目")
            return projects
        
        print("\n未找到任何项目，尝试刷新页面...")
        driver.refresh()
        time.sleep(2)
        
        # 再次尝试获取项目
        rows = driver.find_elements(By.CSS_SELECTOR, '.el-table__body-wrapper .el-table__row')
        for row in rows:
            try:
                cells = row.find_elements(By.CSS_SELECTOR, '.cell')
                if cells:
                    project_text = cells[0].text.strip()
                    if project_text and project_text not in seen_texts:
                        seen_texts.add(project_text)
                        projects.append({"text": project_text, "element": row})
            except:
                continue
        
        if projects:
            print(f"\n刷新后成功获取到 {len(projects)} 个项目")
        else:
            print("\n刷新后仍未找到任何项目")
        
        return projects
        
    except Exception as e:
        print(f"\n获取项目列表失败: {str(e)}")
        return []

def get_subprojects(driver, project_text):
    """获取子项目列表"""
    subproject_selectors = [
        "//tr[contains(@class, 'el-table__expanded-row')]//div[contains(@class, 'cell')]",
        "//div[contains(@class, 'el-table__expanded-cell')]//div[contains(@class, 'cell')]",
        f"//tr[contains(@class, 'el-table__row') and contains(., '{project_text}')]/following-sibling::tr[1]//div[contains(@class, 'cell')]"
    ]
    
    subprojects = []
    for selector in subproject_selectors:
        elements = get_visible_elements(driver, selector)
        for element in elements:
            text = extract_element_text(element)
            if text and text not in [p['text'] for p in subprojects]:
                clickable_element = get_clickable_element(element)
                subprojects.append({"text": text, "element": clickable_element})
    
    return subprojects

def display_menu(items, title="请选择："):
    """显示菜单并获取用户选择"""
    while True:
        print(f"\n{title}")
        for i, item in enumerate(items, 1):
            print(f"{i}. {item['text']}")
        
        try:
            choice = int(input("\n请输入编号: "))
            if 1 <= choice <= len(items):
                selected_item = items[choice-1]
                print("\n" + "="*50)
                print(f"您选择了: {selected_item['text']}")
                print("="*50 + "\n")
                return selected_item
            print("无效的选择，请重试")
        except ValueError:
            print("请输入有效的数字")

def handle_project_navigation(driver, project):
    """处理项目导航"""
    try:
        # 滚动到元素位置
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", project['element'])
        time.sleep(0.5)
        
        # 尝试找到并点击展开箭头
        try:
            expand_icon = project['element'].find_element(By.CSS_SELECTOR, ".el-table__expand-icon")
            if "expanded" not in expand_icon.get_attribute("class"):
                expand_icon.click()
                print("成功点击展开箭头")
                time.sleep(1)
            return True
        except:
            # 如果找不到展开箭头，尝试点击整行
            try:
                project['element'].click()
                print("直接点击项目行")
                time.sleep(1)
                return True
            except:
                print("点击项目失败")
                return False
        
    except Exception as e:
        print(f"项目导航失败: {str(e)}")
        return False

def handle_subproject_navigation(driver, subproject):
    """处理子项目导航"""
    try:
        if not wait_and_click_element(driver, subproject['element']):
            return False
        
        print("\n已进入选择的页面...")
        time.sleep(1)
        
        while True:
            print("\n当前操作选项：")
            print("1. 继续浏览当前页面")
            print("2. 返回选择城市")
            print("3. 退出程序")
            print("4. 保存当前页面内容")
            
            action = input("\n请选择操作 (1/2/3/4): ").strip()
            
            if action == "1":
                print("\n当前页面浏览选项：")
                print("1. 向下滚动")
                print("2. 向上滚动")
                print("3. 回到顶部")
                print("4. 回到底部")
                print("5. 返回上一级菜单")
                
                browse_action = input("\n请选择浏览操作 (1/2/3/4/5): ").strip()
                
                if browse_action == "1":
                    driver.execute_script("window.scrollBy(0, 500);")
                elif browse_action == "2":
                    driver.execute_script("window.scrollBy(0, -500);")
                elif browse_action == "3":
                    driver.execute_script("window.scrollTo(0, 0);")
                elif browse_action == "4":
                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                elif browse_action == "5":
                    continue
                else:
                    print("无效的选择，请重试")
                continue
                
            elif action == "2":
                print("\n正在返回选择城市...")
                driver.get("https://www.95598.cn/osgweb/index")
                time.sleep(2)
                
                # 重新执行到选择城市的步骤
                if not handle_province_selection(driver):
                    return False
                if not handle_price_page_navigation(driver):
                    return False
                return "back_to_city"
                
            elif action == "3":
                confirm = input("\n确定要退出程序吗？(y/n): ").strip().lower()
                if confirm == 'y':
                    print("\n正在退出程序...")
                    return "exit"
                continue
            
            elif action == "4":
                try:
                    page_content = driver.page_source
                    timestamp = time.strftime("%Y%m%d_%H%M%S")
                    filename = f"电价标准_{timestamp}.html"
                    with open(filename, "w", encoding="utf-8") as f:
                        f.write(page_content)
                    print(f"\n页面内容已保存到文件: {filename}")
                except Exception as e:
                    print(f"\n保存页面内容失败: {str(e)}")
                continue
            
            else:
                print("无效的选择，请重试")
        
    except Exception as e:
        print(f"子项目导航失败: {str(e)}")
        return False

def get_provinces(driver):
    """获取省份列表"""
    try:
        # 点击地区选择器
        region_selectors = [
            "//div[@id='city_select']//a[contains(@class, 'current')]",
            "//div[contains(@class, 'region')]//a[contains(@class, 'current')]",
            "//a[contains(@class, 'current fsize16')]",
            "//div[@data-v-07831be2]//a[contains(@class, 'current')]"
        ]
        
        region_element = find_element_with_retry(driver, region_selectors)
        if not region_element or not wait_and_click_element(driver, region_element):
            print("点击地区选择器失败")
            return []
        
        time.sleep(2)
        
        # 获取省份列表
        province_selectors = [
            "//a[contains(@class, 'f66 fsize14')]",
            "//div[contains(@class, 'province-list')]//a",
            "//div[contains(@class, 'region-list')]//a"
        ]
        
        provinces = []
        for selector in province_selectors:
            elements = get_visible_elements(driver, selector)
            for element in elements:
                text = extract_element_text(element)
                if text and text != '省份' and text not in [p['text'] for p in provinces]:
                    provinces.append({"text": text, "element": element})
        
        return provinces
    except Exception as e:
        print(f"获取省份列表失败: {str(e)}")
        return []

def handle_district_selection(driver):
    """处理区县选择"""
    print("\n第四步：选择区县")
    time.sleep(0.5)  # 保留短暂等待
    
    districts = get_districts(driver)
    if not districts:
        print("未找到区县列表")
        return False
    
    selected_district = display_menu(districts, "请选择区县：")
    if not selected_district:
        return False
    
    return wait_and_click_element(driver, selected_district['element'])

def handle_return_navigation(driver, step_name):
    """处理返回导航"""
    print(f"\n正在返回{step_name}...")
    driver.back()
    time.sleep(1)  # 等待基本页面加载
    
    try:
        # 等待页面加载完成
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        
        # 根据不同步骤等待不同的元素
        if "区县" in step_name:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'ul.tab-con-box-ul li.tab-con-box-li'))
            )
        elif "城市" in step_name:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div.el-dialog__body'))
            )
        elif "项目" in step_name:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '.el-table__body-wrapper'))
            )
        
        time.sleep(1)  # 额外等待确保页面元素完全加载
        return True
    except Exception as e:
        print(f"等待页面加载失败: {str(e)}")
        try:
            print("尝试刷新页面...")
            driver.refresh()
            time.sleep(2)
            return True
        except:
            return False

def main():
    """主函数"""
    print("=== 开始运行自动化程序 ===")
    
    driver = None
    try:
        # 初始化驱动
        driver = setup_driver()
        if not driver:
            return
        
        # 访问网站
        print("正在访问网站...")
        driver.get("https://www.95598.cn/osgweb/index")
        time.sleep(3)
        
        while True:  # 主循环
            # 第一步：选择省份
            if not handle_province_selection(driver):
                continue
            
            # 第二步：导航到电价标准页面
            if not handle_price_page_navigation(driver):
                driver.get("https://www.95598.cn/osgweb/index")
                time.sleep(2)
                continue
            
            while True:  # 城市选择循环
                # 第三步：选择城市
                if not handle_city_selection(driver):
                    driver.get("https://www.95598.cn/osgweb/index")
                    time.sleep(2)
                    break
                
                # 第四步：选择区县
                if not handle_district_selection(driver):
                    driver.get("https://www.95598.cn/osgweb/index")
                    time.sleep(2)
                    break
                
                # 第五步：选择电价标准项目
                if not handle_project_selection(driver):
                    driver.get("https://www.95598.cn/osgweb/index")
                    time.sleep(2)
                    break
                
                # 第六步：选择具体项目
                result = handle_subproject_selection(driver)
                if result == "exit":
                    return
                elif not result:
                    driver.get("https://www.95598.cn/osgweb/index")
                    time.sleep(2)
                    break
                
                # 处理子项目导航
                result = handle_subproject_navigation(driver, result)
                if result == "exit":
                    return
                elif result == "back_to_city":
                    continue  # 继续城市选择循环
                elif not result:
                    driver.get("https://www.95598.cn/osgweb/index")
                    time.sleep(2)
                    break
    
    except KeyboardInterrupt:
        print("\n程序被用户中断")
    except Exception as e:
        print(f"\n程序运行出错: {str(e)}")
        import traceback
        print(traceback.format_exc())
    finally:
        if driver:
            print("\n正在关闭浏览器...")
            try:
                driver.quit()
            except Exception as e:
                print(f"关闭浏览器时出错: {str(e)}")
            finally:
                print("程序已结束运行")

def handle_province_selection(driver):
    """处理省份选择"""
    print("\n第一步：选择省份")
    
    # 点击地区选择器
    region_selectors = [
        "//div[@id='city_select']//a[contains(@class, 'current')]",
        "//div[contains(@class, 'region')]//a[contains(@class, 'current')]",
        "//a[contains(@class, 'current fsize16')]"
    ]
    
    region_element = find_element_with_retry(driver, region_selectors, timeout=3)
    if not region_element or not wait_and_click_element(driver, region_element):
        print("点击地区选择器失败")
        return False
    
    time.sleep(0.5)
    
    # 获取省份列表
    province_selectors = ["//a[contains(@class, 'f66 fsize14')]"]  # 简化选择器
    
    provinces = []
    for selector in province_selectors:
        elements = get_visible_elements(driver, selector)
        for element in elements:
            text = extract_element_text(element)
            if text and text != '省份' and text not in [p['text'] for p in provinces]:
                provinces.append({"text": text, "element": element})
    
    if not provinces:
        print("未找到任何省份")
        return False
    
    # 用户选择省份
    selected_province = display_menu(provinces, "请选择省份：")
    if not selected_province:
        return False
    
    # 点击选中的省份
    return wait_and_click_element(driver, selected_province['element'])

def handle_price_page_navigation(driver):
    """处理电价标准页面导航"""
    print("\n第二步：导航到电价标准页面")
    time.sleep(0.5)
    
    # 点击信息公开
    info_selectors = ["//span[@id='right-class' and contains(text(), '信息公开')]"]
    info_btn = find_element_with_retry(driver, info_selectors, timeout=3)
    if not info_btn or not wait_and_click_element(driver, info_btn):
        print("点击信息公开失败")
        return False
    
    time.sleep(0.5)
    
    # 点击电价标准
    price_selectors = ["//a[text()='电价标准']"]
    price_btn = find_element_with_retry(driver, price_selectors, timeout=3)
    if not price_btn or not wait_and_click_element(driver, price_btn):
        print("点击电价标准失败")
        return False
    
    return True

def handle_city_selection(driver):
    """处理城市选择"""
    print("\n第三步：选择城市")
    time.sleep(0.5)  # 保留短暂等待
    
    cities = get_cities(driver)
    if not cities:
        print("未找到城市列表")
        return False
    
    selected_city = display_menu(cities, "请选择城市：")
    if not selected_city:
        return False
    
    return wait_and_click_element(driver, selected_city['element'])

def handle_project_selection(driver):
    """处理项目选择"""
    print("\n第五步：选择电价标准项目")
    
    # 获取项目列表
    projects = get_projects(driver)
    if not projects:
        print("未找到电价标准项目，尝试刷新页面...")
        driver.refresh()
        time.sleep(1)
        projects = get_projects(driver)
        if not projects:
            return False
    
    selected_project = display_menu(projects, "请选择要查看的电价标准项目：")
    if not selected_project:
        return False
    
    return handle_project_navigation(driver, selected_project)

def handle_subproject_selection(driver):
    """处理子项目选择"""
    print("\n第六步：选择具体项目")
    time.sleep(0.5)
    
    subprojects = get_subprojects(driver, "")  # 获取所有子项目
    if not subprojects:
        print("未找到任何子项目")
        return False
    
    return display_menu(subprojects, "请选择要查看的具体项目：")

if __name__ == "__main__":
    main() 
