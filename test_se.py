from selenium import webdriver
import time

driver = webdriver.Chrome()  # 如果使用其他浏览器，如 Firefox，需要相应修改
# 打开一个网站
driver.get("https://www.baidu.com")

# 获取页面标题
print(driver.title)

# 关闭浏览器
driver.quit()