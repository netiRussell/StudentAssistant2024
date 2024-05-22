import pandas
import numpy
from glob import glob
import cv2
import matplotlib.pylab as plt

# Reading images
pictures = glob("./img/sample.png")

img_mpltlb = plt.imread(pictures[0])
img_cv2 = cv2.imread(pictures[0])

# Display image
fig, ax = plt.subplots(1, 1, figsize=(10,10))
# ax[0].imshow(img_mpltlb)
ax.imshow(img_cv2)
# ax[0].set_title('Matplotlib')
# ax[1].set_title('OpenCV2')
plt.show()

# Image manipulation