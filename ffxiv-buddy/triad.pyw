import sys
import cv2
import numpy as np
import pyautogui
from system_hotkey import SystemHotkey
from PyQt5.QtCore import Qt, QTimer, QThread
from PyQt5.QtGui import QImage, QPixmap
from PyQt5.QtWidgets import QApplication, QLabel, QWidget, QVBoxLayout, QHBoxLayout, QCheckBox, QSizePolicy, QLineEdit, QComboBox


class ScreenCapture(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Screen Capture")
        self.setWindowFlags(Qt.WindowStaysOnTopHint)
        self.layout = QVBoxLayout()
        
        self.mouse_label = QLineEdit()
        self.mouse_label.setAlignment(Qt.AlignLeft)
        self.mouse_label2 = QLineEdit()
        self.mouse_label2.setAlignment(Qt.AlignLeft)
        self.layout.addWidget(self.mouse_label)
        self.layout.addWidget(self.mouse_label2)

        self.iamge = None
        self.pixmap = None
        self.image_label = QLabel()
        self.image_label.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Ignored)
        self.image_label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.image_label)
        self.setLayout(self.layout)

        self.timer = QTimer()
        self.timer.timeout.connect(self.update_info)

        self.if_capture = True
        self.checkbox_capture = QCheckBox("capture")
        self.checkbox_capture.stateChanged.connect(self.toggle_capture)
        self.checkbox_capture.setChecked(self.if_capture)

        self.if_battle = True
        self.checkbox_battle = QCheckBox("battle")
        self.checkbox_battle.stateChanged.connect(self.toggle_battle)
        self.checkbox_battle.setChecked(self.if_battle)

        self.if_repeat = False
        self.checkbox_repeat = QCheckBox("repeat")
        self.checkbox_repeat.stateChanged.connect(self.toggle_repeat)
        self.checkbox_repeat.setChecked(self.if_repeat)

        SystemHotkey().register(('control', 'alt', 'z'), callback=lambda _:(
            # self.checkbox_battle.toggle(),
            self.checkbox_repeat.toggle(),
        ))

        self.cb_deck = QComboBox(self)
        self.cb_deck.addItems(['1', '2', '3', '4', '5'])

        checkbox_layout = QHBoxLayout()
        checkbox_layout.addWidget(self.checkbox_capture)
        checkbox_layout.addWidget(self.checkbox_battle)
        checkbox_layout.addWidget(self.checkbox_repeat)
        checkbox_layout.addWidget(self.cb_deck)
        self.layout.addLayout(checkbox_layout)

    def toggle_capture(self, state):
        if state == Qt.Checked:
            self.timer.start(500)  # 2 FPS
        else:
            self.timer.stop()

    def toggle_battle(self, state):
        if state == Qt.Checked:
            self.if_battle = True
        else:
            self.if_battle = False

    def toggle_repeat(self, state):
        if state == Qt.Checked:
            self.if_repeat = True
        else:
            self.if_repeat = False

    def update_info(self):
        screenshot = pyautogui.screenshot()
        im = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2GRAY)

        y, x = pyautogui.position() # w * h
        turn_pos = (456, 503)
        target_l = [(716, 619), (922, 700), (716, 782), (922, 862), (716, 943)]
        target_r = [(504, 1156), (504, 1360), (504, 1564),
                    (701, 1156), (701, 1360), (701, 1564),
                    (950, 1156), (950, 1360), (950, 1564),]
        view_turn = im[turn_pos]
        view_l, view_r = [im[a] for a in target_l], [im[a] for a in target_r] 
        self.mouse_label.setText(f"pos: ({x}, {y}). val[{x}, {y}]: {im[(x,y)]}. turn: {view_turn}. deck: {self.cb_deck.currentText()}")
        self.mouse_label2.setText(f"l: {view_l}. r: {view_r}.")
        
        if self.if_battle:
            turn = im[turn_pos]
            val_l, val_r = [im[a] for a in target_l], [im[a] for a in target_r]
            pxl_l, pxl_r = [150, 202, 76], [102]
            if turn > 127 and sum(val in pxl_l for val in val_l) == 1 and sum(val in pxl_r for val in val_r) == 1:
                pos_l = filter(lambda pos: im[pos] in pxl_l, target_l).__next__()
                pos_r = filter(lambda pos: im[pos] in pxl_r, target_r).__next__()
                pyautogui.moveTo(pos_l[1] - 75, pos_l[0])
                pyautogui.click()
                pyautogui.moveTo(pos_r[1] - 75, pos_r[0])
                pyautogui.click()

            self.image = QImage(
                cv2.cvtColor(im, cv2.COLOR_GRAY2RGB),
                im.shape[1],
                im.shape[0],
                QImage.Format_RGB888
            )
            self.pixmap = QPixmap.fromImage(self.image).scaled(self.image_label.size(), aspectRatioMode=Qt.KeepAspectRatio)    
            self.image_label.setPixmap(self.pixmap)
        else:
            self.image_label.clear()

        if self.if_repeat:
            # 再战
            anchor_list = [((877, 1406), 48), ((793, 1282), 94), ((874, 1388), 78)]
            click_pos = (866, 1170)
            if all(im[pos] == anchor_color for pos, anchor_color in anchor_list):
                pyautogui.moveTo(click_pos[1], click_pos[0])
                pyautogui.click()
            # 挑战
            anchor_list = [((491, 1230), 133), ((740, 1040), 182), ((415, 1537), 159)]
            click_pos = (1001, 1175)
            if all(im[pos] == anchor_color for pos, anchor_color in anchor_list):
                pyautogui.moveTo(click_pos[1], click_pos[0])
                pyautogui.click()
            # 选卡组
            anchor_list = [((428, 1281), 195), ((553, 1304), 125), ((919, 1259), 208)]
            click_pos = (472 + (int(self.cb_deck.currentText()) - 1) * 123, 1283)
            if all(im[pos] == anchor_color for pos, anchor_color in anchor_list):
                pyautogui.moveTo(click_pos[1], click_pos[0])
                pyautogui.click()


    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ScreenCapture()
    window.resize(800, 300)  # 设置窗口固定大小
    window.show()
    sys.exit(app.exec_())