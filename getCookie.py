import json
import time
import os
import tkinter as tk
from tkinter import filedialog
from selenium import webdriver
from selenium.common.exceptions import WebDriverException, NoSuchElementException
from selenium.webdriver.common.by import By


def save_cookies():
    # 创建 WebDriver 对象
    driver = webdriver.Chrome()

    try:
        # 打开网页并最大化窗口
        driver.maximize_window()
        driver.get('https://www.tianyancha.com/nadvance')

        print("请手动登录账户...")

        # 轮询检查登录状态
        while True:
            try:
                # 检查某个只有在成功登录后才会出现的元素
                driver.find_element(By.XPATH, "//*[@id=\"page-header\"]/div/div[3]/div/div[6]/div/a/span")  # 假设这是一个成功登录后出现的元素
                print("登录成功，正在保存 cookies...")
                break  # 登录成功，退出循环
            except NoSuchElementException:
                time.sleep(1)  # 等待一秒后继续检查

        # 获取 cookies
        cookies = driver.get_cookies()
        print(f"获取到的 cookies: {cookies}")  # 打印获取的 cookies

        # 默认保存路径和文件名
        default_file_path = os.path.join(os.getcwd(), "cookies.txt")

        # 选择保存 cookies 的文件，默认文件名为 cookies.txt
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        file_path = filedialog.asksaveasfilename(title="选择保存 cookies 的文件名",
                                                   initialfile="cookies.txt",
                                                   defaultextension=".txt",
                                                   filetypes=[("文本文件", "*.txt")],
                                                   initialdir=os.getcwd())  # 默认目录为当前路径
        if not file_path:
            print("未选择文件，程序终止。")
            return

        # 将 cookies 保存为 JSON 格式
        with open(file_path, 'w') as f:
            json.dump(cookies, f)  # 使用 json.dump 直接保存
        print(f"Cookies 已成功保存到 {file_path}")

    except WebDriverException as e:
        print(f"发生错误：{e}")
    finally:
        driver.quit()


if __name__ == "__main__":
    save_cookies()
