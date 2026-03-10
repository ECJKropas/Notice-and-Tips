import re
import threading
import time
import tkinter as tk

from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By


def slow_scroll(driver, step=500):
    # 找到滚动容器
    container = driver.find_element(By.ID, "workspace")
    # last_height = driver.execute_script("return arguments[0].scrollHeight", container)
    cnt = 0
    while cnt <= 10:
        driver.execute_script(f"arguments[0].scrollBy(0, {step});", container)
        time.sleep(0.5)  # 等待新内容加载
        # new_height = driver.execute_script("return arguments[0].scrollHeight", container)
        cnt += 1


class FloatingTipsApp:
    def __init__(self, root):
        self.root = root
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.configure(bg="#2C3E50")  # 深蓝色背景

        # 初始位置和大小
        self.root.geometry("400x60+500+100")

        # UI 标签
        self.label = tk.Label(
            root,
            text="系统启动中，正在初始化浏览器...",
            fg="#ECF0F1",
            bg="#2C3E50",
            font=("Microsoft YaHei", 10, "bold"),
            wraplength=380,
            justify="left",
        )
        self.label.pack(expand=True, fill="both", padx=10, pady=5)

        # 鼠标拖动绑定
        self.label.bind("<Button-1>", self.start_move)
        self.label.bind("<B1-Motion>", self.do_move)

        # 右键退出
        self.root.bind("<Button-3>", self.on_closing)

        # 启动 Selenium 线程
        self.running = True
        self.update_thread = threading.Thread(
            target=self.selenium_loader_task, daemon=True
        )
        self.update_thread.start()

    def start_move(self, event):
        self.x = event.x
        self.y = event.y

    def do_move(self, event):
        deltax = event.x - self.x
        deltay = event.y - self.y
        x = self.root.winfo_x() + deltax
        y = self.root.winfo_y() + deltay
        self.root.geometry(f"+{x}+{y}")

    # --- Selenium 核心逻辑 ---
    def selenium_loader_task(self):
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--disable-gpu")

        driver = None
        try:
            driver = webdriver.Chrome(options=chrome_options)
            # 建议增加隐式等待
            driver.implicitly_wait(10)

            while self.running:
                try:
                    driver.get("https://www.kdocs.cn/l/colfFw2Piprw")
                    # 等待文档内容加载（根据 KDocs 特性，可能需要一点缓冲）
                    time.sleep(10)
                    # 逻辑：查找类名中包含“关闭图标”的按钮
                    driver.find_element(
                        By.CSS_SELECTOR, "button.is-icon .kd-icon-symbol_cross_two"
                    ).click()
                    slow_scroll(driver)

                    # 1. 抓取页面上所有的 SVG text 标签
                    text_elements = driver.find_elements(By.TAG_NAME, "text")
                    print(text_elements)
                    all_text = "".join([el.text for el in text_elements])
                    print(all_text)

                    # 2. 使用正则提取 [StartNotice] 和 [EndNotice] 之间的内容
                    pattern = r"\[StartNotice\](.*?)\[EndNotice\]"
                    match = re.search(pattern, all_text, re.S)  # re.S 让 . 匹配换行符

                    if match:
                        raw_content = match.group(1)

                        # 3. 自动分条逻辑：识别 "数字." 并换行
                        # 匹配诸如 1. 2. 10. 这种格式，并在前面加上换行符
                        formatted_content = re.sub(
                            r"(\d+\.)", r"\n\1", raw_content
                        ).strip()

                        # 如果需要循环滚动显示，也可以存成列表
                        # tips_list = [line.strip() for line in formatted_content.split('\n') if line.strip()]

                        display_text = f"📢 最新公告：\n{formatted_content}"
                    else:
                        display_text = "未检测到公告内容或标记..."

                    # 更新 UI
                    self.root.after(0, lambda: self.label.config(text=display_text))

                except Exception as e:
                    # 提前将 e 转换为字符串，避免作用域问题
                    error_msg = f"提取失败: {str(e)[:50]}"
                    self.root.after(0, lambda m=error_msg: self.label.config(text=m))

                time.sleep(300)  # 建议拉长抓取间隔（如5分钟），避免触发验证码

        except Exception as e:
            error_msg = f"浏览器故障: {e}"
            self.root.after(0, lambda m=error_msg: self.label.config(text=m))
        finally:
            if driver:
                driver.quit()

    def on_closing(self, event=None):
        self.running = False
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = FloatingTipsApp(root)
    root.mainloop()
