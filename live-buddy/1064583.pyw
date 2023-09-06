import sys
import time
import asyncio
import qasync
import os
import logging
import json
import traceback

import pyttsx4

from bilibili_api import live, Credential, Danmaku, sync
from concurrent.futures import ThreadPoolExecutor
from PyQt5.QtWidgets import QMainWindow, QTextEdit, QLineEdit, QVBoxLayout, QWidget
from PyQt5.QtCore import Qt

class Window(QMainWindow):
    def __init__(self, room:live.LiveRoom):
        super().__init__()
        # 朗读功能
        self.pool = ThreadPoolExecutor(max_workers=1)
        self.speak = lambda text: self.pool.submit(pyttsx4.speak, text)
        # 弹幕功能
        self.room = room
        # 初始化UI
        self.setupUI()

    def setupUI(self):
        self.setWindowTitle('tts')
        self.setStyleSheet("background-color: transparent;")  # 设置窗口背景为透明
        self.resize(300, 500)

        self.text_box = QTextEdit(self)
        self.text_box.setReadOnly(True)
        self.text_box.setStyleSheet('color:white')

        self.entry = QLineEdit(self)
        self.entry.returnPressed.connect(self.on_enter)
        self.entry.setStyleSheet('color:white')

        self.layout = QVBoxLayout()
        self.layout.addWidget(self.text_box)
        self.layout.addWidget(self.entry)

        self.widget = QWidget(self)
        self.widget.setLayout(self.layout)
        self.setCentralWidget(self.widget)

        self.entry.setFocus(Qt.FocusReason.OtherFocusReason)

    def log_and_speak(self, text:str, read:str=None):
        # 发送到聊天框
        self.text_box.append(f"[{time.strftime('%H:%M:%S')}] {text}")
        # 朗读信息
        self.speak(text if read is None else read)

    def on_enter(self):
        text = self.entry.text()
        if text == "":
            return
        if text.startswith('/send'):
            text = text.split(maxsplit=1)[1]
            asyncio.create_task(self.room.send_danmaku(Danmaku(text)))
        else:
            self.log_and_speak(text)
        self.entry.clear()


async def main():
    # basic
    sessdata = '46fe8764%2C1707486611%2C32edb%2A82ZKSbhBHsBBiIAW6TwrWJ0MswKy6veMbbxTqygcbastBxvsKYhatQYeGYrtXxTJWjIVZJHAAAXQA'
    bili_jct = 'bce9c9735b12961ab991574c69914de7'
    buvid3 = 'E7A26BBB-6D43-725B-1C02-455B1E38CFA293294infoc'
    dedeuserid = '22225793'
    ac_time_value = '1c4a56bd16fa41f30e295b6f9eee2a82'
    credential = Credential(sessdata=sessdata, bili_jct=bili_jct, buvid3=buvid3, dedeuserid=dedeuserid, ac_time_value=ac_time_value)

    room_display_id = os.path.basename(sys.argv[0]).split('.')[0]
    danmuku = live.LiveDanmaku(room_display_id=room_display_id, credential=credential)
    room = live.LiveRoom(room_display_id=room_display_id, credential=credential)

    window = Window(room=room)
    window.show()

    # log
    script_dir = os.path.dirname(sys.argv[0])
    file_handler = logging.FileHandler(filename=f'{script_dir}/{room_display_id}.log', mode='a', encoding='utf8')
    file_handler.setFormatter(logging.Formatter(fmt='%(asctime)s - %(levelname)s: %(message)s'))
    file_handler.setLevel(logging.DEBUG)
    logger = logging.getLogger(f'{room_display_id}')
    logger.setLevel(logging.DEBUG)
    logger.addHandler(file_handler)

    @danmuku.on('DANMU_MSG')
    async def on_DANMU_MSG(event):
        # 收到弹幕
        uname = event['data']['info'][2][1]
        text = event['data']['info'][1]
        window.log_and_speak(f'{uname}：{text}', f'{uname}说：{text}')

    @danmuku.on('SEND_GIFT')
    async def on_SEND_GIFT(event):
        # 收到礼物'
        uname = event['data']['data']['uname']
        action = event['data']['data']['action']
        giftName = event['data']['data']['giftName']
        num = event['data']['data']['num']
        window.log_and_speak(f'{uname} {action} {giftName}，数量 {num}')

    @danmuku.on('INTERACT_WORD')
    async def on_INTERACT_WORD(event):
        # 有人进入直播间
        uname = event['data']['data']['uname']
        window.log_and_speak(f'{uname} 进入直播间')

    @danmuku.on('LIVE')
    async def on_LIVE(event):
        # 直播开始
        if 'live_time' in event['data']:
            roomid = event['data']['roomid']
            window.log_and_speak(f'直播开始，房间号{roomid}')

    @danmuku.on('VERIFICATION_SUCCESSFUL')
    async def on_VERIFICATION_SUCCESSFUL(event):
        # 成功连接到直播间
        room_display_id = event['room_display_id']
        window.log_and_speak(f'已成功连接至房间：{room_display_id}')
    

    @danmuku.on('ALL')
    async def on_all(event):
        # 日志
        logger.info(json.dumps(event))

    loop = asyncio.get_running_loop()
    await danmuku.connect()

if __name__ == "__main__":
    try:
        qasync.run(main())
    except Exception as e:
        traceback.print_exc()
        print(sys.version)
        print(sys.argv)
        input()
        sys.exit(0)