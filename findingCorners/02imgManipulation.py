import cv2
import numpy as np

img = cv2.imread('/Users/ruslanabdulin/Desktop/Python/01SA/img/multipleConnected.png')
gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

ret,thresh = cv2.threshold(gray,50,255,0)
contours,hierarchy = cv2.findContours(thresh, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
print(contours)

print("Number of contours detected:", len(contours))

for cnt in contours:
   x1,y1 = cnt[0][0]
   approx = cv2.approxPolyDP(cnt, 0.01*cv2.arcLength(cnt, True), True)
   if len(approx) == 4:
      x, y, w, h = cv2.boundingRect(cnt)
      print("New line: ", x, y, w, h, "\n")

cv2.imshow("Points", img)
cv2.waitKey(0)
cv2.destroyAllWindows()