import cv2
import xlrd
import numpy as np
import matplotlib.pyplot as plt

loc = (r"C:\Users\Administrator\Desktop\list.xlsx")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

img = cv2.imread(r"C:\Users\Administrator\Desktop\certficate.jpg")
drawing = False
ix, iy = -1, -1

fix_img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
img_copy = fix_img

list = ["MUSKAN BHANDARI","BEST ACTOR" , "02-02-2020" , "Ambrish"]
l=[]
i=0

def certi(event,x,y,flags,params):
    
    global ix, iy, drawing, fix_img, img_copy,list,i
    
    
    if event == cv2.EVENT_LBUTTONDOWN:
        ix, iy = x, y
        drawing = True
        img_copy = fix_img.copy()
        
    elif event == cv2.EVENT_MOUSEMOVE:
        if drawing == True:
            cv2.rectangle(fix_img, (ix,iy), (x,y), (0,255,0), -1)
    elif event == cv2.EVENT_LBUTTONUP:
        drawing = False
        cv2.rectangle(fix_img, (ix, iy), (x,y), (0,255,0), -1)
        
        if cv2.waitKey(3000) & 0xFF == ord('c'):
            fix_img = img_copy
            size= (y - iy )/20
            font = cv2.FONT_HERSHEY_PLAIN
            cv2.putText(fix_img,text=list[i],org=(ix,y),fontFace=font,fontScale=size,color=(0,0,0),thickness=3,lineType=cv2.LINE_AA)
            
            l.append([(ix,y),size])
            
            i+=1
            
        else: fix_img = img_copy
  
cv2.namedWindow('draw', cv2.WINDOW_NORMAL)
cv2.setMouseCallback('draw',certi)

while True:
    cv2.imshow('draw', fix_img)
    
    if cv2.waitKey(1) & 0xFF == 27:
        break
        
cv2.destroyAllWindows()
fix_img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    
if i==4:
    for j in range(sheet.nrows):
        r = sheet.row_values(j)
        print(r)
        for  k in range (i):
             cv2.putText(fix_img,text=r[k],org=l[k][0],fontFace= cv2.FONT_HERSHEY_PLAIN,fontScale=l[k][1],color=(0,0,0),thickness=3,lineType=cv2.LINE_AA)
        cv2.imwrite('cert'+str(j)+'.png',fix_img)
        fix_img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            
        
