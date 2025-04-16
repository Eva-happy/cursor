from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
from selenium.webdriver.common.action_chains import ActionChains

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

def wait_and_click_element(driver, element):
    """等待并点击元素"""
    try:
        # 确保元素在视图中
        driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", element)
        time.sleep(0.5)
        
        # 尝试多种点击方法
        try_methods = [
            lambda: element.click(),
            lambda: driver.execute_script("arguments[0].click();", element),
            lambda: ActionChains(driver).move_to_element(element).click().perform()
        ]
        
        for method in try_methods:
            try:
                method()
                return True
            except:
                continue
        
        return False
    except Exception as e:
        print(f"点击元素失败: {str(e)}")
        return False

def get_provinces(driver):
    """获取省份列表"""
    try:
        print("等待切换地区按钮加载...")
        # 等待并点击切换地区按钮
        region_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(@class, 'ant-dropdown-link')]"))
        )
        print("找到切换地区按钮，正在点击...")
        wait_and_click_element(driver, region_button)
        time.sleep(2)  # 增加等待时间
        
        # 获取省份列表
        provinces = []
        expected_provinces = ['广东省', '广西壮族自治区', '云南省', '贵州省', '海南省']
        expected_icons = ['guangdong', 'guangxi', 'yunnan', 'guizhou', 'hainan']
        
        print("等待省份列表加载...")
        # 等待省份列表容器加载
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//ul[@data-v-4c44d656]"))
        )
        
        # 使用更精确的XPath选择器，基于完整的HTML结构
        for i, (province, icon) in enumerate(zip(expected_provinces, expected_icons)):
            try:
                # 使用省份名称和图标类名来精确定位每个省份
                selector = f"//li[@data-v-4c44d656][.//i[contains(@class, '{icon}')] and .//h5[text()='{province}']]"
                print(f"尝试查找省份: {province}")
                
                element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, selector))
                )
                
                if element:
                    provinces.append({"text": province, "element": element})
                    print(f"成功添加省份: {province}")
            except Exception as e:
                print(f"查找省份 {province} 失败: {str(e)}")
                continue
        
        if len(provinces) != 5:
            print(f"警告：只找到 {len(provinces)} 个省份，应该有5个")
            print("找到的省份：")
            for province in provinces:
                print(f"- {province['text']}")
        else:
            print("成功找到全部5个省份")
        
        return provinces
    except Exception as e:
        print(f"获取省份列表失败: {str(e)}")
        import traceback
        print(traceback.format_exc())
        return []

def get_cities(driver):
    """获取城市列表"""
    try:
        time.sleep(1)  # 等待城市列表加载
        
        # 获取城市列表
        cities = []
        city_elements = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, "//ul[@data-v-4c44d656 and @class='cityList']/li[@data-v-4c44d656]/h5[@data-v-4c44d656]"))
        )
        
        for element in city_elements:
            text = element.text.strip()
            if text:
                cities.append({"text": text, "element": element.find_element(By.XPATH, "..")})  # 使用父元素li作为点击目标
        
        return cities
    except Exception as e:
        print(f"获取城市列表失败: {str(e)}")
        return []

def get_page_info(driver):
    """获取页码信息"""
    try:
        print("等待资讯公告列表加载...")
        # 等待资讯公告列表容器加载
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'infoList')]"))
        )
        time.sleep(1)  # 等待内容加载
        
        # 获取总条数和计算总页数
        total_text = driver.find_element(By.XPATH, "//li[@class='ant-pagination-total-text']").text
        total_count = int(total_text.split()[1])
        total_pages = (total_count + 4) // 5  # 每页5条，向上取整
        print(f"\n总共有 {total_pages} 页")
        
        return total_pages
    except Exception as e:
        print(f"获取页码信息失败: {str(e)}")
        import traceback
        print(traceback.format_exc())
        return 0

def get_page_announcements(driver, page_number):
    """获取指定页面的公告列表"""
    try:
        # 如果不是第一页，需要点击到指定页
        if page_number > 1:
            # 查找并点击指定页码
            page_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[contains(@class, 'ant-pagination-item') and @title='{page_number}']"))
            )
            wait_and_click_element(driver, page_button)
            time.sleep(1)  # 等待页面加载
        
        # 获取当前页的公告
        announcements = []
        announcement_elements = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, "//div[@data-v-7707c85a and @class='list-item']"))
        )
        
        for element in announcement_elements:
            try:
                # 获取标题和日期
                title_element = element.find_element(By.XPATH, ".//div[@data-v-7707c85a and @class='esp']")
                date_element = element.find_element(By.XPATH, ".//div[@data-v-7707c85a and @class='timeLine']")
                
                title = title_element.text.strip()
                date = date_element.text.strip()
                
                if title:
                    # 使用包含链接的div作为点击目标
                    link_element = element.find_element(By.XPATH, ".//div[@data-v-7707c85a and @class='link']")
                    announcements.append({
                        "text": f"{title} ({date})",
                        "element": link_element,
                        "page": page_number
                    })
            except Exception as e:
                print(f"处理公告元素失败: {str(e)}")
                continue
        
        return announcements
    except Exception as e:
        print(f"获取第 {page_number} 页公告失败: {str(e)}")
        return []

def display_menu(items, title="请选择：", allow_return=True):
    """显示菜单并获取用户选择"""
    while True:
        if any('page' in item for item in items):  # 检查是否包含页码信息
            # 获取所有可用页码
            pages = sorted(list(set(item['page'] for item in items if 'page' in item)))
            
            print("\n=== 可选页码 ===")
            for page in pages:
                print(f"{page}. 第 {page} 页")
            if allow_return:
                print("0. 返回上一步")
            
            try:
                page_choice = int(input("\n请输入页码: "))
                if page_choice == 0 and allow_return:
                    return None
                if page_choice not in pages:
                    print(f"无效的页码，请输入: {', '.join(map(str, pages))}")
                    continue
                
                # 筛选选中页码的项目
                page_items = [item for item in items if item.get('page') == page_choice]
                
                print(f"\n=== 第 {page_choice} 页的公告列表 ===")
                for i, item in enumerate(page_items, 1):
                    print(f"{i}. {item['text']}")
                if allow_return:
                    print("0. 返回上一步")
                
                try:
                    item_choice = int(input(f"\n请输入编号(0-{len(page_items)}): "))
                    if item_choice == 0 and allow_return:
                        continue
                    if 1 <= item_choice <= len(page_items):
                        selected_item = page_items[item_choice-1]
                        print(f"\n您选择了: {selected_item['text']}")
                        return selected_item
                    print(f"无效的选择，请输入0-{len(page_items)}之间的数字")
                except ValueError:
                    print("请输入有效的数字")
            except ValueError:
                print("请输入有效的页码")
        else:
            # 原有的非分页显示逻辑
            print(f"\n{title}")
            print("可选项：")
            for i, item in enumerate(items, 1):
                print(f"{i}. {item['text']}")
            if allow_return:
                print("0. 返回上一步")
            
            try:
                choice = int(input(f"\n请输入编号(0-{len(items)}): "))
                if choice == 0 and allow_return:
                    return None
                if 1 <= choice <= len(items):
                    selected_item = items[choice-1]
                    print(f"\n您选择了: {selected_item['text']}")
                    return selected_item
                print(f"无效的选择，请输入0-{len(items)}之间的数字")
            except ValueError:
                print("请输入有效的数字")

def main():
    """主函数"""
    print("=== 南方电网电价信息提取程序 ===")
    
    driver = None
    try:
        # 初始化驱动
        driver = setup_driver()
        if not driver:
            return
        
        while True:  # 主循环
            # 访问网站
            print("\n正在访问网站...")
            driver.get("https://95598.csg.cn/#/gd/serviceInquire/information/list")
            time.sleep(3)
            
            # 获取省份列表
            print("\n正在获取省份列表...")
            provinces = get_provinces(driver)
            if not provinces:
                print("未找到省份列表")
                return
            
            # 选择省份
            selected_province = display_menu(provinces, "请选择省份：")
            if selected_province is None:
                continue  # 返回主循环开始
            if not wait_and_click_element(driver, selected_province['element']):
                print("点击省份失败")
                continue
                
            # 获取城市列表
            print("\n正在获取城市列表...")
            cities = get_cities(driver)
            if not cities:
                print("未找到城市列表")
                continue
            
            # 选择城市
            selected_city = display_menu(cities, "请选择城市：")
            if selected_city is None:
                continue  # 返回主循环开始
            if not wait_and_click_element(driver, selected_city['element']):
                print("点击城市失败")
                continue
            
            # 获取总页数
            total_pages = get_page_info(driver)
            if total_pages == 0:
                print("获取页码信息失败")
                continue
                
            # 显示页码选择菜单
            while True:  # 页码选择循环
                print("\n=== 可选页码 ===")
                for page in range(1, total_pages + 1):
                    print(f"{page}. 第 {page} 页")
                print("0. 返回上一步")
                
                try:
                    page_choice = int(input("\n请输入页码: "))
                    if page_choice == 0:
                        break  # 退出页码选择循环，返回主循环
                    
                    if 1 <= page_choice <= total_pages:
                        # 获取选中页面的公告
                        announcements = get_page_announcements(driver, page_choice)
                        if not announcements:
                            print(f"第 {page_choice} 页没有找到公告")
                            continue
                        
                        # 显示该页公告列表
                        print(f"\n=== 第 {page_choice} 页的公告列表 ===")
                        for i, item in enumerate(announcements, 1):
                            print(f"{i}. {item['text']}")
                        print("0. 返回上一步")
                        
                        try:
                            item_choice = int(input(f"\n请输入编号(0-{len(announcements)}): "))
                            if item_choice == 0:
                                continue  # 返回页码选择
                            
                            if 1 <= item_choice <= len(announcements):
                                selected_announcement = announcements[item_choice-1]
                                print(f"\n您选择了: {selected_announcement['text']}")
                                
                                # 点击选中的公告
                                print(f"\n正在打开公告: {selected_announcement['text']}")
                                if not wait_and_click_element(driver, selected_announcement['element']):
                                    print("点击公告失败")
                                    continue
                                    
                                print("\n已打开选中的公告页面")
                                
                                while True:  # 操作选择循环
                                    print("\n请选择操作：")
                                    print("1. 继续浏览")
                                    print("2. 返回上一步")
                                    print("3. 退出程序")
                                    
                                    try:
                                        choice = int(input("\n请输入编号(1-3): "))
                                        if choice == 1:
                                            print("\n继续浏览...")
                                            time.sleep(1)
                                        elif choice == 2:
                                            break  # 返回页码选择
                                        elif choice == 3:
                                            return  # 退出程序
                                        else:
                                            print("无效的选择，请输入1-3之间的数字")
                                    except ValueError:
                                        print("请输入有效的数字")
                            else:
                                print(f"无效的选择，请输入0-{len(announcements)}之间的数字")
                        except ValueError:
                            print("请输入有效的数字")
                    else:
                        print(f"无效的页码，请输入0-{total_pages}之间的数字")
                except ValueError:
                    print("请输入有效的页码")
        
    except Exception as e:
        print(f"\n程序运行出错: {str(e)}")
        import traceback
        print(traceback.format_exc())
    finally:
        if driver and not keep_browser_open:
            print("\n正在关闭浏览器...")
            try:
                driver.quit()
            except:
                pass
            print("程序已结束运行")

if __name__ == "__main__":
    keep_browser_open = True  # 设置为True以保持浏览器打开
    main()
