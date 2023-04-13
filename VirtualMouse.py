import cv2
import numpy as np
import HandDetector as hd
import autopy
import time
import sys
import math
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
volume.GetMute()
volume.GetMasterVolumeLevel()
volRange = volume.GetVolumeRange()
minVol = volRange[0]
maxVol = volRange[1]


Wcam, Hcam = 1280, 720
pTime = 0
Hframe = 250    
Lframe = 444
smoothening = 10
pLocX, pLocY = 0, 0
cLocX, cLocY = 0, 0

detector = hd.HandDetector(maxHands=2, detectionCon=0.7)
handType = []


cap = cv2.VideoCapture(0)
cap.set(3, Wcam)
cap.set(4, Hcam)
wSc, hSc = autopy.screen.size()


while True:
    success, img = cap.read()
    img, handType = detector.findHands(img)
    lmlist, bbox = detector.findPosition(img) 
    if len(lmlist) != 0:
        x1, y1 = lmlist[8][1:]
        x2, y2 = lmlist[12][1:]
        x3, y3 = lmlist[4][1:]
        x4, y4 = lmlist[20][1:]
        cv2.rectangle(img, (Lframe, Hframe), (Wcam-Lframe, Hcam-Hframe), (255, 0, 255), 2)
        fingers = detector.fingersUp()  
        if handType == "Left":
            if fingers[1] == 1 and fingers[2] == 0:
                x5 = np.interp(x1, (Lframe, Wcam-Lframe), (0, wSc))
                y5 = np.interp(y1, (Hframe, Hcam-Hframe), (0, hSc))
                cLocX = pLocX + (x5-pLocX)/smoothening
                cLocY = pLocY + (y5-pLocY)/smoothening
                try:
                    autopy.mouse.move(wSc-cLocX, cLocY)
                    cv2.circle(img, (x1, y1), 15, (255, 0, 255), cv2.FILLED)
                    pLocX, pLocY = cLocX, cLocY
                except:
                    pass
                        
            if fingers[1]==1 and fingers[2]==1:
                length, img, lineInfo = detector.findDistance(8, 12, img)  
                length1, img1, lineInfo1 = detector.findDistance(8, 4, img)   
                if length < 40:
                    cv2.circle(img, (lineInfo[4], lineInfo[5]), 15, (0, 255, 0), cv2.FILLED)
                    autopy.mouse.toggle(autopy.mouse.Button.LEFT, False)
                    autopy.mouse.toggle(autopy.mouse.Button.MIDDLE, False)
                    # autopy.mouse.click(autopy.mouse.Button.LEFT, True)  
                    autopy.mouse.toggle(autopy.mouse.Button.LEFT, True) 
                    time.sleep(0.1)
                    autopy.mouse.toggle(autopy.mouse.Button.LEFT, False)
                if length1 < 40:
                    autopy.mouse.toggle(autopy.mouse.Button.LEFT, True) 
                    x5 = np.interp(x1, (Lframe, Wcam-Lframe), (0, wSc))
                    y5 = np.interp(y1, (Hframe, Hcam-Hframe), (0, hSc))
                    cLocX = pLocX + (x5-pLocX)/smoothening
                    cLocY = pLocY + (y5-pLocY)/smoothening
                    try:
                        autopy.mouse.move(wSc-cLocX, cLocY)
                        cv2.circle(img, (x1, y1), 15, (255, 0, 255), cv2.FILLED)
                        pLocX, pLocY = cLocX, cLocY
                    except:
                        pass
            if fingers[4]==1:
                length, img, lineInfo = detector.findDistance(4, 20, img)
                if length < 40:
                    cv2.circle(img, (lineInfo[4], lineInfo[5]), 15, (0, 255, 0), cv2.FILLED)
                    autopy.mouse.toggle(autopy.mouse.Button.LEFT, False)
                    autopy.mouse.toggle(autopy.mouse.Button.MIDDLE, True)
                    time.sleep(0.1)
                    autopy.mouse.toggle(autopy.mouse.Button.MIDDLE, False)
            if fingers[1] == 0 and fingers[2] == 0 and fingers[3] == 0 and fingers[4] == 0:
                cv2.circle(img, (lineInfo1[4], lineInfo1[5]), 15, (0, 255, 0), cv2.FILLED)
                autopy.mouse.toggle(autopy.mouse.Button.LEFT, False)
                # autopy.mouse.click(autopy.mouse.Button.MIDDLE, False)  
                autopy.mouse.toggle(autopy.mouse.Button.RIGHT, True) 
                time.sleep(0.1)
                autopy.mouse.toggle(autopy.mouse.Button.RIGHT, False)
        elif handType == "Right":
            cv2.circle(img, (x1, y1), 15, (255, 0, 255), cv2.FILLED)
            cv2.circle(img, (x3, y3), 15, (255, 0, 255), cv2.FILLED)
            cv2.line(img, (x1, y1), (x3, y3), (255, 0, 255), 3)
            
            length = int(math.hypot(x3-x1, y3-y1))
            vol = np.interp(length, [50, 250], [minVol, maxVol])
            volume.SetMasterVolumeLevel(vol, None)

    if cv2.waitKey(1) == 27:
        cap.release()
        cv2.destroyAllWindows()
        sys.exit()



    cTime = time.time()
    fps = 1/(cTime-pTime)
    pTime = cTime
    cv2.putText(img, str(int(fps)),(20, 50), cv2.FONT_HERSHEY_PLAIN, 3, (255, 0, 0), 3)


    cv2.imshow("Virtual Mouse", img)