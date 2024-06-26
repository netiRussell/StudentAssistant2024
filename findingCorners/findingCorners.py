import cv2 as cv
import numpy as np
import matplotlib.pyplot as plt

# Load the img and prepare it
img = cv.imread('/Users/ruslanabdulin/Desktop/Python/01SA/img/multipleDisconected.png')
gray = cv.cvtColor(img, cv.COLOR_BGR2GRAY)

# find Harris corners
gray = np.float32(gray)
dst = cv.cornerHarris(gray,2,3,0.04)
dst = cv.dilate(dst,None)
ret, dst = cv.threshold(dst,0.01*dst.max(),255,0)
dst = np.uint8(dst)

# find centroids
ret, labels, stats, centroids = cv.connectedComponentsWithStats(dst)

# define the criteria to stop and refine the corners
criteria = (cv.TERM_CRITERIA_EPS + cv.TERM_CRITERIA_MAX_ITER, 100, 1)
corners = cv.cornerSubPix(gray,np.float32(centroids),(5,5),(-1,-1),criteria)

# Now draw them
res = np.hstack((centroids,corners))
res = np.intp(res)
img[res[:,1],res[:,0]]=[0,0,255]
img[res[:,3],res[:,2]] = [0,255,0]


# Show side by side
fig, (ax1, ax2) = plt.subplots(1, 2)
fig.suptitle('Plots')
ax1.scatter(corners[:, 0], corners[:, 1])
ax2.imshow(img)
plt.show()

# cv.imshow('dst',img)
# if cv.waitKey(0) & 0xff == 27:
#  cv.destroyAllWindows()

# ! Make sure shapes with a small area are disregarded