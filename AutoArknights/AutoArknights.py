# Copyright (c) 2025 jw <amada.minami@protonmail.com>, All rights reserved.

# 只支持MuMu模拟器，默认2560 * 1440分辨率，DPI 360

import pytesseract
import subprocess
import numpy
import cv2
import os
import time
import logging
import sys
import chardet
import socket
import configparser
import yagmail
import base64

from PIL import Image
from logging.handlers import RotatingFileHandler

def get_config():
    config = configparser.ConfigParser()
    config.read("AutoArknights.ini", encoding="utf-8")
    return config

def is_network_available():
    try:
        # 尝试连接 Google 的公共 DNS 服务器
        socket.create_connection(("8.8.8.8", 53), timeout=5)
        return True
    except OSError:
        return False

def close_application(process_name):
    try:
        logger = logging.getLogger()
        subprocess.run(['taskkill', '/f', '/im', process_name], check=True)
        print(f"{process_name} killed.")
    except subprocess.CalledProcessError:
        print(f"Can't Kill {process_name}")
        logger.info("Can't Kill " + process_name)

# 重启adb server
def reboot_adb_server():
    config = get_config()
    adb_path = config.get("adb", "path", fallback="adb")

    logger = logging.getLogger()
    result = subprocess.run([adb_path, 'kill-server'], capture_output=True, text=True)
    print(result.stdout)
    logger.info('Kill adb server, ' + result.stdout)
    result = subprocess.run([adb_path, 'start-server'], capture_output=True, text=True)
    print(result.stdout)
    if (("daemon started successfully" in result.stdout) or ("daemon started successfully" in result.stderr)):
        time.sleep(3)
        logger.info('Startup adb server successfully.')
        return True
    logger.info('Startup adb server failed, ' + result.stdout + ',' + result.stderr)
    go_to_exit(1)
    exit(1)
    return False

# 关闭模拟器
def close_mumu_simulator():
    logger = logging.getLogger()
    close_application("MuMuPlayer.exe")
    close_application("MuMuVMMSVC.exe")
    close_application("MuMuVMMHeadless.exe")
    logger.info('MuMu simulator closed.')

# 启动模拟器，异步启动
def open_mumu_simulator(player_path):
    logger = logging.getLogger()
    config = get_config()
    adb_path = config.get("adb", "path", fallback="adb")
    simulator_adb_addr = config.get("adb", "addr", fallback="127.0.0.1:16384")

    subprocess.Popen([player_path])
    time.sleep(30)
    result = subprocess.run([adb_path, 'connect', simulator_adb_addr], capture_output=True, text=True)
    print(result.stdout)
    logger.info('Open mumu simulator, ' + result.stdout)
    if (("connected to " + simulator_adb_addr) in result.stdout):
        logger.info('Open mumu simulator successfully.')
        return True
    logger.info('Open mumu simulator failed.')
    return False
    
# 重启模拟器
def reboot_mumu_simulator():
    close_mumu_simulator()
    time.sleep(3)

    config = get_config()
    simulator_path = config.get("simulator", "path", fallback=r"C:\Program Files\Netease\MuMu Player 12\shell\MuMuPlayer.exe")
    result = open_mumu_simulator(simulator_path)
    if not result:
        print("Failed to starup mumu simulator.")
        go_to_exit(1)
        exit(1)
        return
    print("mumu simulator started.")


# 关闭MAA
def close_MAA():
    logger = logging.getLogger()
    close_application('MAA.exe')
    logger.info('MAA Closed.')

# 启动MAA， 启动MAA会自动执行明日方舟
# 如果此时明日方舟已经打开，MAA也不会再开明日方舟了
def open_MAA(MAA_path):
    logger = logging.getLogger()
    subprocess.Popen([MAA_path])
    time.sleep(5)
    logger.info('MAA started.')

# 重启MAA
def reboot_MAA():
    close_MAA()
    time.sleep(3)
    config = get_config()
    maa_path = config.get("maa", "path", fallback=r"D:\software\MAA-v5.3.1-win-x64\MAA.exe")
    
    open_MAA(maa_path)

def adb_get_gray_screenshot(x, y, width, height):
    logger = logging.getLogger()
    config = get_config()
    adb_path = config.get("adb", "path", fallback="adb")
    # 运行adb进行截图
    result = subprocess.run([adb_path, 'exec-out', 'screencap', '-p'], stdout=subprocess.PIPE)
    screenshot = numpy.frombuffer(result.stdout, dtype=numpy.uint8)
    # 用opencv读取截图
    screenshot = cv2.imdecode(screenshot, cv2.IMREAD_COLOR)
    # 提取感兴趣的区域并转换为灰度图
    logger.info('adb get gray screenshot ' + str(x) + ',' + str(y) + ',' + str(width) + ',' + str(height))
    roi = screenshot[int(y):int(y)+int(height), int(x):int(x)+int(width)]
    gray_roi = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
    return gray_roi

# 点击屏幕上某个区域
def adb_tap(x, y):
    logger = logging.getLogger()
    config = get_config()
    adb_path = config.get("adb", "path", fallback="adb")
    subprocess.run([adb_path, 'shell', 'input', 'tap', str(x), str(y)])
    logger.info('adb tap ' + str(x) + ',' + str(y))

# 按下，移动，再抬起
def adb_swipe(start_x, start_y, end_x, end_y, duration):
    logger = logging.getLogger()
    config = get_config()
    adb_path = config.get("adb", "path", fallback="adb")
    subprocess.run([adb_path, 'shell', 'input', 'swipe', 
        str(start_x), str(start_y), 
        str(end_x), str(end_y), 
        str(duration)])
    logger.info('adb swipe ' + 
        str(start_x) + ',' + str(start_y) + ',' + 
        str(end_x) + ',' + str(end_y) + ',' + str(duration))

def setup_logging():
    # 清空之前的日志文件
    with open('AutoArknights.log', 'w') as f:
        f.truncate(0)  # 清空文件内容

    # 创建一个 RotatingFileHandler 对象
    handler = RotatingFileHandler('AutoArknights.log', maxBytes=5*1024*1024, backupCount=3, mode='w')
    handler.setLevel(logging.DEBUG)
    # 设置日志格式
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    # 将 handler 添加到日志器
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    logger.addHandler(handler)
    logger.info('AutoArkninghts logger registed.')

def startup_arknights():
    logger = logging.getLogger()

    config = get_config()
    x = config.get("arknights", "text_x", fallback=2055)
    y = config.get("arknights", "text_y", fallback=795)
    w = config.get("arknights", "text_w", fallback=218)
    h = config.get("arknights", "text_h", fallback=69)
    # 获取在模拟器中，明日方舟图标下的文字，以确认已经启动到Launcher界面了。
    app_name_image = adb_get_gray_screenshot(x, y, w, h)
    cv2.imwrite('app_name_image.png', app_name_image)
    text = pytesseract.image_to_string(app_name_image, lang='chi_sim')
    if ('明日方舟' not in text):
        logger.error('Can\'t find out arknights in MuMu simulator.')
        go_to_exit(1)
        exit(1)
        return
    logger.info('The Arknights icon has been acquired, let\'s click to start.')
    x = config.get("arknights", "icon_x", fallback=2133)
    y = config.get("arknights", "icon_y", fallback=725)
    adb_tap(x, y)      #   明日方舟，启动！
    time.sleep(30)
    # 启动后识别辅助的“最小化”按钮
    x = config.get("helper", "min_btn_x", fallback=519)
    y = config.get("helper", "min_btn_y", fallback=105)
    w = config.get("helper", "min_btn_w", fallback=131)
    h = config.get("helper", "min_btn_h", fallback=54)
    min_image = adb_get_gray_screenshot(x, y, w, h)
    min_image = cv2.adaptiveThreshold(min_image, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
    cv2.imwrite('min_image.png', min_image)
    text = pytesseract.image_to_string(min_image, lang='chi_sim')
    print('text of min image: ' + text)
    if ('最小化' not in text):
        logger.error('Can\'t find the minimize button for the helper.')
        go_to_exit(1)
        exit(1)
        return
    logger.info('The Helper\'s min button has been acquired.')
    x = config.get("helper", "min_btn_tap_x", fallback=578)
    y = config.get("helper", "min_btn_tap_y", fallback=126)
    adb_tap(x, y)      # 把辅助器最小化
    time.sleep(1)
    x = config.get("helper", "min_btn_start_x", fallback=59)
    y = config.get("helper", "min_btn_start_y", fallback=181)
    x1 = config.get("helper", "min_btn_end_x", fallback=184)
    y1 = config.get("helper", "min_btn_end_y", fallback=1382)
    adb_swipe(x, y, x1, y1, 3000) # 把辅助移动到一个合适的位置

def detect_encoding(file_path):
    with open(file_path, 'rb') as file:
        raw_data = file.read()
        result = chardet.detect(raw_data)
        return result['encoding']

def wait_MAA_finished():
    logger = logging.getLogger()
    config = get_config()
    maa_guilog_path = config.get("maa", "gui_log_path", fallback=r'D:\software\MAA-v5.3.1-win-x64\debug\gui.log')

    # 自动检测编码
    encoding = detect_encoding(maa_guilog_path)
    logger.info(f'Detected encoding: {encoding}')
    with open(maa_guilog_path, 'r', encoding=encoding) as file:
        file.seek(0, 2)  # 移动到文件末尾
        while True:
            line = file.readline().strip()
            if line:
                logger.info('MAA: ' + line)
                if '任务已全部完成' in line:
                    logger.info('All finished, go to sleep.')
                    return
            else:
                #logger.info('MAA: Wait MAA finish 1s.')
                time.sleep(1)

def go_to_exit(exit_code):
    # 进入睡眠前，收拾一下
    close_mumu_simulator()
    close_MAA()
    config = get_config()
    adb_path = config.get("adb", "path", fallback="adb")
    subprocess.run([adb_path, 'kill-server'], capture_output=True, text=True)

    # 报告执行状态
    logger = logging.getLogger()
    logger.info("Entering sleep mode after 10s...")
    time.sleep(10)  # 额外等待 10 秒
    report(exit_code)
    subprocess.run(['shutdown', '/h', '/f'], capture_output=True, text=True)


def wait_network_avaliable():
    logger = logging.getLogger()
    while not is_network_available():
        logger.info('Network is not avaliable, I\'ll check it after 1s')
        time.sleep(1)

def report(exit_code):
    config = get_config()

    mail_from = config.get("report", "from", fallback="xxx@163.com")
    mail_from_passwd = config.get("report", "password", fallback="passwd")
    print('mail_from: ' + mail_from)
    print('mail_from_passwd: ' + mail_from_passwd)
    mail_to = config.get("report", "to", fallback="xxx@163.com")
    log_path = config.get("arknights", "log_path", fallback=r"D:\script\AutoArknights.log")
    print('mail_to: ' + mail_to)
    print('log_path: ' + log_path)
    yag = yagmail.SMTP(mail_from, mail_from_passwd, host="smtp.163.com", port=587, smtp_ssl=True)
    # 发送邮件
    if (exit_code <= 0):
        yag.send(
            to=mail_to,                              # 收件人
            subject="Arknights周期日常运行报告",            # 邮件主题
            contents="运行成功"  # 邮件正文
        )
    else:
        with open(log_path, "rb") as f:
            encoded_file = base64.b64encode(f.read()).decode()
        html_content = f"""
            <p>运行失败，请检查log附件。</p>
            <p>你的邮件服务可能拦截了附件。</p>
            <p>请复制以下 Base64 代码，并在 <a href="https://www.base64decode.net/">Base64 解码网站</a> 粘贴解码。</p>
            <pre>{encoded_file}</pre>
            """
        yag.send(
            to=mail_to,
            subject="Arknights周期日常运行报告",
            contents=html_content
        )
    yag.close()


if __name__ == '__main__':
    setup_logging()
    logger = logging.getLogger()
    wait_network_avaliable()
    logger.info('Start to exec AutoArknights.')
    reboot_adb_server()
    reboot_mumu_simulator()
    startup_arknights()
    reboot_MAA()
    logger.info('Hello MAA and Byebye!')
    # 至此，把控制权交给MAA了，以后就MAA负责处理后续事宜
    wait_MAA_finished()
    go_to_exit(0)


