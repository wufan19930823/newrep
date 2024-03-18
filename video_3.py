import time
import logging
import pyautogui
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC



# 配置日志记录器  
logging.basicConfig(level=logging.INFO,  
                    format='%(asctime)s - %(levelname)s - %(message)s',  
                    filename='selenium_log.log',  # 日志文件名  
                    filemode='w')  # 写入模式，覆盖之前的日志  
  
console = logging.StreamHandler()  # 创建一个handler，用于写入日志文件  
console.setLevel(logging.INFO)  # 再创建一个handler，用于输出到控制台  
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')  # 定义handler的输出格式  
console.setFormatter(formatter)  # 给handler添加formatter  
logging.getLogger('').addHandler(console)  # 给logger添加handler









###########################################################
# 正式用
#读取Excel文件
excel_path = 'D:\\领克中技汽修听课表.xlsx'  # 修改为你的Excel文件实际路径
df = pd.read_excel(excel_path)
print("可用的列名：", df.columns)
if '身份证号' not in df.columns:
    raise KeyError(f"列名 '身份证号' 在Excel文件中不存在，请检查文件内容。")
ids = df['身份证号'].dropna().tolist()  # 假设Excel中身份证号的列名是'身份证号'，并去掉NaN值
###########################################################
# 测试用
#ids = ['130731199712243426', '150922200003210510']
###########################################################

# 初始化WebDriver，指定geckodriver路径（如果已添加到PATH则不需要）
# 你可以根据需要设置firefox_options，比如无头模式等
driver = webdriver.Firefox()

# 最大化浏览器
driver.maximize_window()

# 访问登录页面
driver.get('http://www.hbace.cn/index.html')

# 循环遍历身份证号，并尝试登录
for id_number in ids:
    try:
        logging.info(f"登陆：{id_number}")

        # 拉起浏览器
        wait = WebDriverWait(driver, 10)

        # 输入用户名密码
        username_element = wait.until(EC.presence_of_element_located((By.ID, 'username')))
        password_element = wait.until(EC.presence_of_element_located((By.ID, 'password')))
        logging.info("已找到用户名和密码输入框，正在输入...")
        username_element.send_keys(str(id_number))
        password_element.send_keys('123456')
        logging.info('用户名和密码已输入')
        time.sleep(1)

        # 点击登陆
        logging.info("正在点击登录按钮...") 
        login_button = driver.find_element(By.ID, 'login')  # 确保ID是正确的
        login_button.click()
        time.sleep(1)

        # 关闭手机号绑定
        try:
            logging.info("正在查找手机号绑定弹窗...") 
            phone_bind_close = driver.find_element(By.XPATH, '//*[@id="layui-layer1"]/span[1]')
            logging.info("已找到手机号绑定弹窗，正在关闭...")
            phone_bind_close.click()
        except Exception as e:
            logging.info("未找到手机号绑定弹窗，无需处理。")

        # 找进入学习按钮
        for index in range(1, 10):
            logging.info(f"正在查找第{index}个可能的进入学习按钮...")
            enter_learn = wait.until(
                EC.presence_of_element_located((By.XPATH, f'//*[@id="planlist"]/div/span/a[{index}]')))
            if enter_learn.text == "进入学习":
                logging.info("已找到进入学习按钮，正在点击...")
                enter_learn.click()
                break
        time.sleep(1)

        # 等待课程列表加载完毕
        logging.info("正在等待课程列表加载完毕...")
        
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="course"]')))
        logging.info(f"身份证号 {id_number} 登录成功并进入学习页面。")
        time.sleep(1)

        # 遍历课程列表
        for index in range(1, 50):
            logging.info(f"正在查找第{index}个课程的标题和进度...")
            try:
                logging.info(f"正在等待标题元素出现：index={index}, id_number={id_number}")
                title = wait.until(
                    EC.presence_of_element_located((By.XPATH, f'//*[@id="course"]/div[{index}]/div[2]/div/h4')))
                logging.info(f"标题元素已找到：{title.text}")
                logging.info(f"正在等待进程元素出现：index={index}, id_number={id_number}")
                process = wait.until(
                    EC.presence_of_element_located((By.XPATH, f'//*[@id="course"]/div[{index}]/div[2]/div/p[2]/span')))
                logging.info(f"进程元素已找到：{process.text}")
                logging.info(f"{id_number} {title.text} {process.text}")
                if title.text not in ['领克汽修电工基础(57分)', '领克汽修机械制图(80分)','领克汽修新能源汽车技术概论(30分)', '领克汽修金属材料及热处理常识(30分)', '领克汽修汽车文化教学材料(30分)']:
                    logging.info(f"跳过课程：{title.text}，因为标题不匹配")
                    continue
                if process.text == "已学习完毕":
                    logging.info(f"课程{title.text}已学习完毕，跳过...")
                    continue

                logging.info(f"开始点击开始学习按钮：index={index}, id_number={id_number}")
                start_learn = wait.until(
                    EC.presence_of_element_located((By.XPATH, f'//*[@id="course"]/div[{index}]/div[2]/div/p[4]/a[1]')))
                start_learn.click()
                time.sleep(1)
                menu_learn = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="continuestudy"]')))
                logging.info(f"课程{title}点击继续学习按钮...") 
                menu_learn.click()
                time.sleep(1)
                
                # 开始播放
                logging.info(f"开始点击播放按钮：index={index}, id_number={id_number}")
                play_button = wait.until(EC.presence_of_element_located(
                    (By.XPATH, '/html/body/div[2]/div/div/div[2]/div[2]/div/div[3]/div/div[1]/div')))
                play_button.click()
                last_time = ""

                driver.execute_script("document.querySelector('video').playbackRate = 5;")
                time.sleep(3)

                
                # 等待视频播完
                while True:
                    time.sleep(5)
                    try:
                        logging.info("检查进度条元素...")
                        process_box = driver.find_element(By.XPATH,
                                                          '/html/body/div[2]/div/div/div[2]/div[2]/div/div[3]/div/div[4]')
                        logging.info("确保进度条元素可见...")
                        driver.execute_script("arguments[0].style.display = 'block';", process_box)

                        logging.info("等待当前时间元素...")
                        time_now = wait.until(EC.presence_of_element_located(
                            (By.XPATH, '/html/body/div[2]/div/div/div[2]/div[2]/div/div[3]/div/div[4]/div[4]/span[1]')))

                        logging.info("等待总时间元素...") 
                        time_total = wait.until(EC.presence_of_element_located(
                            (By.XPATH, '/html/body/div[2]/div/div/div[2]/div[2]/div/div[3]/div/div[4]/div[4]/span[3]')))
                        if time_now.text == time_total.text:
                            logging.info("进度完成，退出循环...")
                            break
                        if last_time == "" or last_time != time_now.text:
                            last_time = time_now.text
                        elif last_time == time_now.text:
                            logging.info("进度没有变化，尝试点击播放按钮...") 
                            new_play_button = driver.find_element(By.XPATH,
                                                              '/html/body/div[2]/div/div/div[2]/div[2]/div/div[3]/div/div[1]/div')
                            
                            
                            
                            if new_play_button.is_displayed():
                                logging.info("播放按钮可见，点击播放按钮...") 
                                driver.execute_script("arguments[0].style.display = 'block';", new_play_button)
                                new_play_button.click()

                                #如果出现进度条弹窗，按下回车键
                                pyautogui.press('enter')
                                pyautogui.press('enter')

                                
                                #关闭弹窗
                                try:  
                                    # 等待弹窗出现（如果需要）  
                                    WebDriverWait(driver, 10).until(EC.alert_is_present())  
  
                                    # 获取弹窗对象  
                                    alert = Alert(driver)  
  
                                    # 关闭弹窗  
                                    alert.dismiss()  # 或者使用 alert.accept() 来接受弹窗
                                    

                                    '''
                                    # 假设弹窗有一个特定的ID或类名，你可以通过JavaScript来关闭它  
                                    driver.execute_script("""  
                                        var dialog = document.querySelector('#dialog-id'); // 替换为实际的ID或选择器  
                                        if (dialog) {  
                                            dialog.style.display = 'none'; // 或者其他关闭弹窗的方法  
                                        }  
                                    """)
                                    '''


                                #不管用的话手动点击播放按钮
                                #pyautogui.moveTo(686,758,duration=1)
                                #pyautogui.click()

                                except Exception as e:  
                                    logging.info(f"没有检测到弹窗: {e}")

                                
                    except Exception as e:  
                        logging.error(f"发生错误: {e}")  
                        #break  # 如果发生错误，退出循环

                #点击返回菜单按钮
                try:
                    logging.info("查找返回菜单按钮...")
                    menu_return = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="mycourse"]')))

                    logging.info("点击返回菜单按钮...") 
                    menu_return.click()
                    time.sleep(1)
                except Exception as e:  
                    logging.error(f"在点击返回菜单按钮时发生错误: {e}")
            except Exception as e:
                logging.error(f"标题元素未找到：{e}")
                #break

        time.sleep(3)
        #continue
        # 退出登陆
        #logout = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[2]/a[2]')))
        #logout.click()
    except Exception as e:
        logging.error(f"发生错误：{e}")
        #logging.error(f"登录失败：身份证号 {id_number}，错误信息：{e}")
        #continue  # 继续尝试下一个身份证号登录
        pass

#driver.quit()
