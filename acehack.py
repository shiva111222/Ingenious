import cv2
import time
import numpy as np
import pyfirmata
import xlsxwriter
workbook=xlsxwriter.Workbook("percent1.xlsx")
worksheet=workbook.add_worksheet()
cap = cv2.VideoCapture(1)
board=pyfirmata.Arduino('COM4')
_, frame = cap.read()
h,w = frame.shape[:2]
nop=h*w
i=0
c=0
n=0
row=0
while(1):
 c=0
 _, frame = cap.read()
 for x in range(h):
   for y in range(w):
    (B,G,R)=frame[x,y]
   
    if ((B,G,R)>(200,200,200)):
     c=c+1
    
 percent=(c/nop)*100
 worksheet.write(row, 0,percent)
 row+=1
 n=n+1
 print('plant has blast disease % ',percent)
 cv2.imshow('frame',frame)
 board.digital[8].write(1)
 board.digital[9].write(0)
 board.digital[10].write(0)
 board.digital[11].write(1)
 time.sleep(2)
 board.digital[8].write(0)
 board.digital[9].write(0)
 board.digital[10].write(0)
 board.digital[11].write(0)
 if n==10:
   workbook.close()
 k = cv2.waitKey(5) & 0xFF
 if k == 27:
  break

cv2.destroyAllWindows()
cap.release()
