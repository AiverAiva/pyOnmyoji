import numpy as np
import cv2 as cv
import time

from modules import ImageTransformer

transformer = ImageTransformer()
while "Screen Capturing":
    screenshot = np.array(transformer.getScreenshot('Calculator'))
    # target_area=screenshot
    cv.imshow("Target Area", screenshot)
    time.sleep(0.03)

    if cv.waitKey(25) & 0xFF == ord("q"):
        cv.destroyAllWindows()
        break