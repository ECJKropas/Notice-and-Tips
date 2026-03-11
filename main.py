#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Floating tips app with carousel rotation.

Changes:
- Extract tips into a list and rotate the displayed tip once per minute.
- Show a live countdown in seconds until the next rotation.
- Selenium fetching runs in a background thread and updates the tips list periodically.
- UI updates the displayed tip and countdown every second via tkinter's `after`.
"""

import re
import threading
import time
import tkinter as tk
from typing import List

from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService


def slow_scroll(driver, step=500):
    """
    Slowly scroll a container to ensure content loads into the DOM (KDocs specific).
    """
    try:
        container = driver.find_element(By.ID, "workspace")
    except Exception:
        return
    cnt = 0
    while cnt <= 10:
        try:
            driver.execute_script(f"arguments[0].scrollBy(0, {step});", container)
        except Exception:
            break
        time.sleep(0.5)
        cnt += 1


class FloatingTipsApp:
    def __init__(self, root):
        self.root = root
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.configure(bg="#2C3E50")  # 深蓝色背景
        self.root.attributes("-alpha", 0.7)  # 默认半透明

        # 初始位置和大小
        self.root.geometry("460x80+500+100")

        # Data for carousel
        self.tips_lock = threading.Lock()
        self.tips: List[str] = []  # list of tip strings
        self.current_index = 0
        self.tip_shown_at = time.time()  # timestamp when current tip started showing
        self.rotation_seconds = 60  # 每条显示时长（秒）

        # Running flag
        self.running = True
        self.fetch_interval_seconds = 300  # selenium fetch interval (默认300秒)

        # UI 标签（两行：提示内容 + 倒计时）
        self.tip_label = tk.Label(
            root,
            text="系统启动中，正在初始化浏览器...",
            fg="#ECF0F1",
            bg="#2C3E50",
            font=("Microsoft YaHei", 10, "bold"),
            wraplength=440,
            justify="left",
        )
        self.tip_label.pack(side="top", expand=True, fill="both", padx=10, pady=(8, 0))

        self.countdown_label = tk.Label(
            root,
            text="下次更新: -- 秒",
            fg="#F1C40F",
            bg="#2C3E50",
            font=("Microsoft YaHei", 9),
            anchor="e",
            justify="right",
        )
        self.countdown_label.pack(side="bottom", fill="x", padx=10, pady=(0, 8))

        # 鼠标拖动绑定
        self.tip_label.bind("<Button-1>", self.start_move)
        self.tip_label.bind("<B1-Motion>", self.do_move)
        self.tip_label.bind("<ButtonRelease-1>", self.stop_move)
        self.countdown_label.bind("<Button-1>", self.start_move)
        self.countdown_label.bind("<B1-Motion>", self.do_move)
        self.countdown_label.bind("<ButtonRelease-1>", self.stop_move)

        # 右键退出
        self.root.bind("<Button-3>", self.on_closing)

        # 启动 Selenium 线程（守护线程）
        self.update_thread = threading.Thread(
            target=self.selenium_loader_task, daemon=True
        )
        self.update_thread.start()

        # 启动 UI 每秒更新（倒计时 + 轮播）
        self._schedule_ui_update()

    # ---------------------
    # 窗口拖动相关
    # ---------------------
    def start_move(self, event):
        self.x = event.x
        self.y = event.y
        self.root.attributes("-alpha", 1.0)  # 拖动时不透明

    def do_move(self, event):
        deltax = event.x - self.x
        deltay = event.y - self.y
        x = self.root.winfo_x() + deltax
        y = self.root.winfo_y() + deltay
        self.root.geometry(f"+{x}+{y}")

    def stop_move(self, event):
        self.root.attributes("-alpha", 0.7)  # 拖动结束后恢复半透明

    # ---------------------
    # Selenium fetch thread
    # ---------------------
    def _parse_tips_from_raw(self, raw_content: str) -> List[str]:
        """
        Parse raw content between [StartNotice] and [EndNotice] into a list of tips.

        Strategy:
        - Normalize lines.
        - Insert newlines before enumerations like '1.' '2.' etc (if not already).
        - Split by blank lines or enumeration markers to produce individual tips.
        """
        if not raw_content:
            return []

        # Ensure enumerations start on new lines (e.g., "1. text")
        text = re.sub(r"\s*(\d+\.)\s*", r"\n\1 ", raw_content)

        # Normalize different newline styles and strip
        text = text.replace("\r\n", "\n").replace("\r", "\n").strip()

        # Split into candidate lines by newline; group lines that belong together
        lines = [ln.strip() for ln in text.split("\n") if ln.strip()]

        tips: List[str] = []
        current_tip_lines: List[str] = []

        # Heuristic: if line starts with enumerator like "1." treat as new tip
        for ln in lines:
            if re.match(r"^\d+\.\s+", ln):
                # finish previous
                if current_tip_lines:
                    tips.append(" ".join(current_tip_lines).strip())
                # start new
                current_tip_lines = [re.sub(r"^\d+\.\s*", "", ln).strip()]
            else:
                # Continue current tip
                current_tip_lines.append(ln)

        if current_tip_lines:
            tips.append(" ".join(current_tip_lines).strip())

        # Final cleanup: remove empty and very short fragments
        clean = [t for t in [t.strip() for t in tips] if t and len(t) > 1]
        return clean

    def selenium_loader_task(self):
        """
        Background thread that periodically fetches the page and updates self.tips.
        """
        edge_options = EdgeOptions()
        edge_service = EdgeService(executable_path="edge/msedgedriver.exe")
        edge_options.add_argument("--headless")
        edge_options.add_argument("--disable-gpu")

        driver = None
        try:
            driver = webdriver.Edge(options=edge_options, service=edge_service)
            driver.implicitly_wait(10)

            while self.running:
                try:
                    # Update temporary UI state
                    self.root.after(
                        0, lambda: self.tip_label.config(text="正在抓取公告...")
                    )
                    driver.get("https://www.kdocs.cn/l/colfFw2Piprw")
                    # 等待页面加载
                    time.sleep(8)
                    # 关闭可能的弹窗（选择性）
                    try:
                        driver.find_element(
                            By.CSS_SELECTOR, "button.is-icon .kd-icon-symbol_cross_two"
                        ).click()
                        time.sleep(0.5)
                    except Exception:
                        pass

                    self.root.after(
                        0, lambda: self.tip_label.config(text="抓取并解析中...")
                    )
                    slow_scroll(driver)

                    # 1. 抓取页面上所有的 SVG text 标签（此前实现）
                    text_elements = driver.find_elements(By.TAG_NAME, "text")
                    all_text = "".join([el.text for el in text_elements])

                    # 2. 使用正则提取 [StartNotice] 和 [EndNotice] 之间的内容
                    pattern = r"\[StartNotice\](.*?)\[EndNotice\]"
                    match = re.search(pattern, all_text, re.S)

                    if match:
                        raw_content = match.group(1)
                        parsed = self._parse_tips_from_raw(raw_content)

                        if parsed:
                            # Replace tips atomically
                            with self.tips_lock:
                                self.tips = parsed
                                self.current_index = 0
                                self.tip_shown_at = time.time()
                            # Optional: immediately update UI to first tip
                            self._refresh_ui_now()
                        else:
                            # No parsed tips
                            with self.tips_lock:
                                self.tips = []
                            self.root.after(
                                0,
                                lambda: self.tip_label.config(
                                    text="未检测到有效公告内容"
                                ),
                            )
                    else:
                        with self.tips_lock:
                            self.tips = []
                        self.root.after(
                            0,
                            lambda: self.tip_label.config(
                                text="未检测到公告标记 [StartNotice]/[EndNotice]"
                            ),
                        )

                except Exception as e:
                    # Keep UI informed
                    msg = f"抓取失败: {str(e)[:80]}"
                    self.root.after(0, lambda m=msg: self.tip_label.config(text=m))

                # 等待下次抓取（线程睡眠）
                # 使用循环睡眠检测运行状态以便快速退出
                sleep_left = self.fetch_interval_seconds
                while sleep_left > 0 and self.running:
                    time.sleep(1)
                    sleep_left -= 1

        except Exception as e:
            error_msg = f"浏览器启动失败: {e}"
            self.root.after(0, lambda m=error_msg: self.tip_label.config(text=m))
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass

    # ---------------------
    # UI 更新逻辑（每秒）
    # ---------------------
    def _schedule_ui_update(self):
        """
        Schedule the next UI update in 1 second.
        """
        if not self.running:
            return
        self._update_display_and_countdown()
        # schedule again in 1 second
        self.root.after(1000, self._schedule_ui_update)

    def _refresh_ui_now(self):
        """
        Force immediate UI refresh from the main thread.
        """
        self.root.after(0, self._update_display_and_countdown)

    def _update_display_and_countdown(self):
        """
        Update the displayed tip and countdown label. This runs in the main thread (tkinter).
        """
        now = time.time()
        with self.tips_lock:
            tips_copy = list(self.tips)

        if tips_copy:
            # Ensure current_index is within range
            if self.current_index >= len(tips_copy):
                self.current_index = 0
                self.tip_shown_at = now

            elapsed = now - self.tip_shown_at
            # If elapsed exceeds rotation_seconds, advance index
            if elapsed >= self.rotation_seconds:
                advance_by = int(elapsed // self.rotation_seconds)
                self.current_index = (self.current_index + advance_by) % len(tips_copy)
                # Recompute shown_at to align with the rotation period
                self.tip_shown_at += advance_by * self.rotation_seconds
                elapsed = now - self.tip_shown_at

            remaining = max(0, int(self.rotation_seconds - elapsed))

            # Build display text: headline + tip
            display_text = f"📢 最新公告 ({self.current_index + 1}/{len(tips_copy)}):\n{tips_copy[self.current_index]}"
            countdown_text = f"下次更新: {remaining} 秒"

        else:
            # No tips available: show placeholder and keep a short refresh cadence
            display_text = "未检测到公告内容或标记，正在等待抓取..."
            countdown_text = "下次更新: -- 秒"

        # Update labels (only if text changed to reduce flicker)
        # Use .after(0, ...) is not needed because we are already on main thread, but safe to ensure thread-safety.
        self.tip_label.config(text=display_text)
        self.countdown_label.config(text=countdown_text)

    # ---------------------
    # 关闭与退出
    # ---------------------
    def on_closing(self, event=None):
        self.running = False
        # Give background threads a moment to stop gracefully
        # (They are daemon threads, but we attempt a polite stop)
        try:
            # destroy root in the main thread
            self.root.destroy()
        except Exception:
            pass


if __name__ == "__main__":
    root = tk.Tk()
    app = FloatingTipsApp(root)
    root.mainloop()
