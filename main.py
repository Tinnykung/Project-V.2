import cv2
import threading
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk

import HandTrackingModule as htm
import pyautogui, autopy
import numpy as np
import math
import time
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

class HandControlApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Computer screen control system using AI hand gesture detection")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        self.root.configure(bg="#222831")

        # Live feed frame
        self.video_frame = tk.Label(self.root) 
        self.video_frame.pack(pady=10)

        # Start/Stop buttons
        self.btn_frame = tk.Frame(self.root, bg="#222831")
        self.btn_frame.pack()

        self.start_btn = tk.Button(self.btn_frame, text="▶️ Start", width=15, command=self.start)
        self.start_btn.grid(row=0, column=0, padx=10)

        self.stop_btn = tk.Button(self.btn_frame, text="⏹️ Stop", width=15, state='disabled', command=self.stop)
        self.stop_btn.grid(row=0, column=1, padx=10)

        self.running = False
        self.cap = None

        self.volPer = 0
        self.volBar = 400

        self.prev_x, self.prev_y = 0, 0
        self.smoothening = 4  # ค่ามาก = เคลื่อนช้า แต่เนียน

        self.slide_start_x = None
        self.slide_detected = False

        self.last_slide_time = 0  # เวลา slide ล่าสุด
        self.slide_delay = 1.0    # delay 1 วินาที

        self.history_x, self.history_y = [], []
        self.history_size = 3  # เก็บ 5 เฟรมล่าสุด
        
    def start(self):
        self.running = True
        self.start_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        self.thread = threading.Thread(target=self.run_hand_control)
        self.thread.start()

    def stop(self):
        self.running = False
        self.start_btn.config(state='normal')
        self.stop_btn.config(state='disabled')

    def update_image(self, imgtk):
        self.video_frame.imgtk = imgtk
        self.video_frame.configure(image=imgtk)

    def putText(self, img, mode, loc=None, color=(0, 255, 255)):
        if loc is None:
            text_size = cv2.getTextSize(str(mode), cv2.FONT_HERSHEY_COMPLEX_SMALL, 3, 3)[0]
            img_h, img_w = img.shape[:2]
            x = (img_w - text_size[0]) // 2
            y = img_h - 60  # 🔧 ปรับจากเดิม -30 เป็น -60 เพื่อเลี่ยงซ้อน
            loc = (x, y)
        cv2.putText(img, str(mode), loc, cv2.FONT_HERSHEY_COMPLEX_SMALL, 3, color, 3)
        return img

    def run_hand_control(self):
        wCam, hCam = 1280, 720
        self.cap = cv2.VideoCapture(0)
        self.cap.set(3, wCam)
        self.cap.set(4, hCam)

        detector = htm.handDetector(maxHands=1, detectionCon=0.85, trackCon=0.8)

        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        volRange = volume.GetVolumeRange()
        minVol = -63
        maxVol = volRange[1]

        hmin, hmax = 50, 200
        vol = 0
        volPer = 0
        self.volPer = volPer
        color = (0, 215, 255)
        tipIds = [4, 8, 12, 16, 20]

        mode = ''
        active = 0
        pTime = 0

        pyautogui.FAILSAFE = False

        while self.running:
            success, img = self.cap.read()
            img = detector.findHands(img)
            img = cv2.resize(img, (640, 480))
            img = cv2.flip(img, 1)
            lmList = detector.findPosition(img, draw=False)
            fingers = []

            if len(lmList) != 0:
                # Thumb
                fingers.append(1 if lmList[tipIds[0]][1] > lmList[tipIds[0] - 1][1] else 0)

                for id in range(1, 5):
                    fingers.append(1 if lmList[tipIds[id]][2] < lmList[tipIds[id] - 2][2] else 0)

                if (fingers == [0, 0, 0, 0, 0]) and (active == 0):
                    mode = 'N'
                elif (fingers == [0, 1, 0, 0, 0] or fingers == [0, 1, 1, 0, 0]) and (active == 0):
                    mode = 'Scroll'
                    active = 1
                elif fingers == [1, 1, 0, 0, 0] and active == 0:
                    mode = 'Volume'
                    active = 1
                elif fingers == [1, 1, 1, 1, 1] and active == 0:
                    mode = 'Cursor'
                    active = 1
                elif fingers == [1, 1, 1, 0, 0] and active == 0:
                    mode = 'Slide'
                    active = 1

            # Scroll Mode
            if mode == 'Scroll':
                active = 1
            #   print(mode)
                cv2.rectangle(img, (200, 410), (245, 460), (255, 255, 255), cv2.FILLED)
                if len(lmList) != 0:
                    if fingers == [0,1,0,0,0]:
                    #print('up')
                    #time.sleep(0.1)
                        self.putText(img, mode = 'U', loc=(200, 455), color = (0, 255, 0))
                        pyautogui.scroll(300)
                    if fingers == [0,1,1,0,0]:
                        #print('down')
                    #  time.sleep(0.1)
                        self.putText(img, mode = 'D', loc =  (200, 455), color = (0, 0, 255))
                        pyautogui.scroll(-300)
                    elif fingers == [0, 0, 0, 0, 0]:
                        active = 0
                        mode = 'N'

            # Volume Mode
            if mode == 'Volume':
                active = 1
            #print(mode)
                if len(lmList) != 0:
                    if fingers[-1] == 1:
                        active = 0
                        mode = 'N'
                        print(mode)
                    else:
                        #   print(lmList[4], lmList[8])
                            x1, y1 = lmList[4][1], lmList[4][2]
                            x2, y2 = lmList[8][1], lmList[8][2]
                            cx, cy = (x1 + x2) // 2, (y1 + y2) // 2
                            cv2.circle(img, (x1, y1), 10, color, cv2.FILLED)
                            cv2.circle(img, (x2, y2), 10, color, cv2.FILLED)
                            cv2.line(img, (x1, y1), (x2, y2), color, 3)
                            cv2.circle(img, (cx, cy), 8, color, cv2.FILLED)

                            length = math.hypot(x2 - x1, y2 - y1)
                            # print(length)

                            # Volume Range -65 - 0
                            vol = np.interp(length, [hmin, hmax], [minVol, maxVol])
                            volBar = np.interp(vol, [minVol, maxVol], [400, 150])
                            volPer = np.interp(vol, [minVol, maxVol], [0, 100])
                            print(vol)
                            volN = int(vol)
                            if volN % 4 != 0:
                                volN = volN - volN % 4
                                if volN >= 0:
                                    volN = 0
                                elif volN <= -64:
                                    volN = -64
                                elif vol >= -11:
                                    volN = vol

                            #print(int(length), volN)
                            volume.SetMasterVolumeLevel(vol, None)
                            if length < 50:
                                cv2.circle(img, (cx, cy), 11, (0, 0, 255), cv2.FILLED)

                            cv2.rectangle(img, (30, 150), (55, 400), (209, 206, 0), 3)
                            cv2.rectangle(img, (30, int(volBar)), (55, 400), (215, 255, 127), cv2.FILLED)
                            cv2.putText(img, f'{int(volPer)}%', (25, 430), cv2.FONT_HERSHEY_COMPLEX, 0.9, (209, 206, 0), 3)

            # Cursor Mode
            if mode == 'Cursor':
                cv2.rectangle(img, (110, 20), (620, 350), (255, 255, 255), 3)
                if fingers[1:] == [0, 0, 0, 0]:
                    active = 0
                    mode = 'N'
                elif len(lmList) != 0:
                    # คำนวณตำแหน่งกลางฝ่ามือ
                    keypoints_x = [lmList[0][1], lmList[5][1], lmList[9][1], lmList[13][1], lmList[17][1]]
                    keypoints_y = [lmList[0][2], lmList[5][2], lmList[9][2], lmList[13][2], lmList[17][2]]
                    center_x = int(sum(keypoints_x)/len(keypoints_x))
                    center_y = int(sum(keypoints_y)/len(keypoints_y))

                    # ขนาดหน้าจอ
                    w, h = autopy.screen.size()
                    X = int(np.interp(center_x, [110, 620], [w-1, 0])) 
                    Y = int(np.interp(center_y, [70, 350], [0, h-1]))

                    # === Moving Average Filter ===
                    if not hasattr(self, 'history_x'):
                        self.history_x, self.history_y = [], []
                        self.history_size = 3

                    self.history_x.append(X)
                    self.history_y.append(Y)

                    if len(self.history_x) > self.history_size:
                        self.history_x.pop(0)
                        self.history_y.pop(0)

                    avg_x = sum(self.history_x) / len(self.history_x)
                    avg_y = sum(self.history_y) / len(self.history_y)

                    # === Dynamic Smoothening ===
                    dx = abs(avg_x - self.prev_x)
                    dy = abs(avg_y - self.prev_y)
                    speed = math.hypot(dx, dy)
                    dynamic_smooth = max(1, 2 - min(speed / 100, 1))

                    smooth_x = self.prev_x + (avg_x - self.prev_x) / dynamic_smooth
                    smooth_y = self.prev_y + (avg_y - self.prev_y) / dynamic_smooth

                    autopy.mouse.move(smooth_x, smooth_y)
                    self.prev_x, self.prev_y = smooth_x, smooth_y

                    # ตรวจจับการคลิก
                    x1, y1 = lmList[4][1], lmList[4][2]   # นิ้วโป้ง
                    x2, y2 = lmList[8][1], lmList[8][2]   # นิ้วชี้
                    length = math.hypot(x2 - x1, y2 - y1)

                    click_threshold = 40
                    if fingers[1] == 0:  # คลิกซ้าย
                        cv2.circle(img, (lmList[8][1], lmList[8][2]), 10, (0, 0, 255), cv2.FILLED)
                        time.sleep(0.25)
                        pyautogui.click()

                    if fingers[2] == 0:  # คลิกขวา
                        cv2.circle(img, (lmList[12][1], lmList[12][2]), 10, (0, 255, 0), cv2.FILLED)
                        time.sleep(0.25)
                        pyautogui.rightClick()

                    # Mouse hold (drag & drop)
                    if length < click_threshold:
                        if not hasattr(self, 'holding'):
                            self.holding = False
                        if not self.holding:
                            self.holding = True
                            pyautogui.mouseDown()
                            print("🖱️ Mouse Hold Down")
                    else:
                        if hasattr(self, 'holding') and self.holding:
                            self.holding = False
                            pyautogui.mouseUp()
                            print("🖱️ Mouse Released")


            #Slide mode
            if mode == "Slide":
                 active = 1
                 if len(lmList) != 0:
                    x8 = lmList[8][1]  # นิ้วชี้
                    x12 = lmList[12][1]  # นิ้วกลาง
                    x_center = (x8 + x12) // 2
                    # วาดเส้นแบ่งกลางจอ เพื่อช่วยเล็งแนวปัดซ้าย–ขวา
                    cv2.line(img, (320, 0), (320, 480), (255, 255, 255), 2)  # เส้นแนวตั้งกลางจอ
                    cv2.putText(img, "Swipe Left <---", (20, 50), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 0, 0), 2)
                    cv2.putText(img, "---> Swipe Right", (340, 50), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2)

                    # เริ่มจับตำแหน่งเมื่อเข้าสู่โหมด
                    if self.slide_start_x is None:
                        self.slide_start_x = x_center
                        self.slide_detected = False
                    else:
                        diff_x = x_center - self.slide_start_x

                        current_time = time.time()
                        if abs(diff_x) > 60 and (current_time - self.last_slide_time) > self.slide_delay:
                            if diff_x > 0:
                                print("👉 Slide Right")
                                self.putText(img, mode='➡️', loc=(300, 400), color=(0, 255, 255))
                                # ใส่คำสั่ง slide ขวา เช่น pyautogui.press('right')
                                pyautogui.press('right')
                                self.last_slide_time = current_time

                            else:
                                print("👈 Slide Left")
                                self.putText(img, mode='⬅️', loc=(300, 400), color=(0, 255, 255))
                                # ใส่คำสั่ง slide ซ้าย เช่น pyautogui.press('left')
                                pyautogui.press('left')
                                self.last_slide_time = current_time

                            self.slide_detected = True
                    # รีเซ็ตเมื่อปล่อยนิ้วทั้งหมด
                    if fingers == [0, 0, 0, 0, 0]:
                        self.slide_start_x = None
                        self.slide_detected = False
                        mode = 'N'
                        active = 0

            display_mode = {
                'Scroll': 'Scroll Mode',
                'Volume': 'Volume Mode',
                'Cursor': 'Cursor Mode',
                'Slide': 'Slide Mode',
                'N': ''
            }.get(mode, '')
            img = self.putText(img, display_mode)

            # FPS
            cTime = time.time()
            fps = 1 / (cTime - pTime + 0.01)
            pTime = cTime
            cv2.putText(img, f'FPS: {int(fps)}', (480, 50), cv2.FONT_ITALIC, 1, (255, 0, 0), 2)

            # Convert for Tkinter
            img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(img)
            imgtk = ImageTk.PhotoImage(image=img)
            self.root.after(0, self.update_image, imgtk)

        self.cap.release()


    def on_closing(self):
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            self.stop()
            self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = HandControlApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()