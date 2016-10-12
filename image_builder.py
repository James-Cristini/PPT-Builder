"""
from PIL import Image
import sys
from time import sleep

img_1 = Image.open("img_test/AUS_Adexli-02.jpg")
img_2 = Image.open("img_test/AUS_Adexli-03.jpg")
img_3 = Image.open("img_test/AUS_Adexli-04.jpg")
img_4 = Image.open("img_test/AUS_Adexli-05.jpg")
blank_image = Image.open("img_test/blank.jpg")
blank_image.show()
img_1.show(img_1)
img_2.show(img_2)
img_3.show(img_3)
img_4.show(img_4)

final_img = "img_test/full.jpg"

blank_image.paste(img_1, (0,0))
blank_image.paste(img_2, (400,0))
blank_image.paste(img_3, (0,300))
blank_image.paste(img_4, (400,300))
blank_image.save(final_img)

im = Image.open(final_img)
im.show()
"""
