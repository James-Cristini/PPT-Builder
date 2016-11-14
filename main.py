
"""
The ppt builder is intended for use by the Market Research department at Addison Whitney.

The program assists in putting together handwriting sample slides (for now) for brand name evaluation decks.
Takes in a folder of handwriting samples of potential brand names organized by country and builds a number of
slides (200-300+) in a matter of minutes (rather than the hours it was taking prior)

Please note that handwriting samples are not included for confidentiality purposes

This program is intended for use on Windows and will not build and execute a Power Point presentation on Mac
"""

__author__ = "James Cristini"
__credits__ = ["James Cristini", 'Adam Tilly (for the "Cristini Genie" name)']
__version__ = "1.3"
__maintainer__ = "James Cristini"
__email__ = "jacristi0428@gmail.com"

import os
import sys
import sip
from PyQt4 import QtGui, QtCore, uic
from PyQt4.QtGui import QApplication, QMainWindow, QMessageBox
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
#from logic import build_deck

class MainWindow(QMainWindow):

    INSTRUCTION_TEXT = """Copy the entire HW folder (and make sure the folder is called HW)
to the ppt_builder folder before pressing the "Build Slides" button!

Be sure your folders are structed as the example below:

ppt_builer (parent folder containing the data folder and ppt_builer.exe)
 - HW (folder containg country folders - remember this folder must be named HW)
   - AUS (country folder containing handwriting sample image files)
   - USA (be sure these folder are named for the 3-letter country code)
   - CAN
   - JPN
"""

    def __init__(self):
        super(MainWindow, self).__init__()
        #Loads the UI file; kept the file separate since it can be easily stored ain the data folder
        self.ui = uic.loadUi('data/UI.ui')
        self.ui.setWindowIcon(QtGui.QIcon("data/pencil.ico"))

        self.ui.progress_bar.setValue(0)

        self.ui.text_output.setPlainText(self.INSTRUCTION_TEXT)

        self.ui.start_btn.clicked.connect(self.build_deck)

        self.ui.show()


    # Parses out brand names from image files
    def get_list_of_names(self, start_dir, country_list):
        names_from_files = []

        for x in country_list:
            country_path = start_dir + x
            file_list = [x for x in os.listdir(country_path)]
            for i in file_list:
                line = i.split("-")
                if i[-3:] == "jpg":
                    line2 = line[0].split("_")
                    if line2[1] not in names_from_files:
                        names_from_files.append(line2[1])

        return names_from_files

    # Builds the actual PPT deck
    def build_deck(self) :
        start_dir = os.getcwd() + "\HW\\"
        country_list = [x for x in os.listdir(start_dir)]
        names_list = self.get_list_of_names(start_dir, country_list)
        HW = {}

        total = len(country_list) * len(names_list)
        self.ui.progress_bar.setMaximum(total)


        for c in country_list:
            HW[c] = {}
            for n in names_list:
                HW[c][n] = []

        # Creates a dictionary mapping out countries, names, and file names for each
        # Sets prority on certain countries in terms of placement in the PPT
        for c in HW:
            for n in HW[c]:
                count = 1
                c_range = 5
                if c == "USA" or c == "EU":
                    c_range = 10
                for x in range(0, c_range):
                    if count == 10:
                        file_name = c + "_" + n + "-10.jpg"
                    else:
                        file_name = c + "_" + n + "-0" + str(count) + ".jpg"
                    HW[c][n].append(file_name)
                    count += 1


        #Stars a new PPT deck
        prs_slides = {}

        prs = Presentation('ppts/aw_template.pptx') # Opens AW master template
        hw_slide_layout = prs.slide_layouts[5] # Slide used for handwriting samples

        # Sets an error image in case the image cannot be found
        ERROR_IMG = "data/error_img.png"

        # Progress bar starting value
        bar_val = 0

        #Text to output in the UI hile building deck
        out_text = ""

        # Build slides
        for name in names_list:
            for country in country_list:

                count = 0
                # Updated UI output text changes as slides are built, update progress bar accordingly
                out_text = "{} {} slide added\n".format(name, country)
                self.ui.text_output.setPlainText(out_text)
                self.ui.progress_bar.setValue(bar_val)

                # New slide is created and added to the pr_slides dictionary / PPT deck
                sld = "Slide %d" % (x)
                prs_slides[sld] = prs.slides.add_slide(hw_slide_layout)
                title = prs_slides[sld].placeholders[18] #Title text in Red
                subtitle = prs_slides[sld].placeholders[16] # Subtitle text in Grey

                title.text = "HANDWRITING SAMPLES"
                subtitle.text = name
                textbox = prs_slides[sld].shapes[2] #Gets and removes the uneeded text area
                sp = textbox.element
                sp.getparent().remove(sp)

                # Add grey box

                left = Inches(1.25)
                top = Inches(1.5)
                ht = Inches(4.5)
                wd = Inches(7.5)
                prs_slides[sld].shapes.add_picture("data/grey_box.png", left, top, wd, ht)

                # Add flag image

                flag_img = "data/flag_{}.png".format(country)
                left = Inches(8)
                top = Inches(0.7)
                ht = Inches(0.7)
                wd = Inches(0.7)
                try:
                    prs_slides[sld].shapes.add_picture(flag_img, left, top, wd, ht)
                except: # Add erroor image if the flag is not found
                    prs_slides[sld].shapes.add_picture(ERROR_IMG, left, top, wd, ht)

                # Add HW sample iamges to each
                for x in range(len(HW[country][name])):
                    img_file = "HW\\" + country + "\\" + HW[country][name][x]
                    try:
                        im = Image.open(img_file)
                    except:
                        im = Image.open(ERROR_IMG)
                    ht = Inches(1.25)
                    wd = Inches(2.0)

                    # Position based on the number found in the image file name
                    if "01" in img_file or "06" in img_file or "02" in img_file:
                        top = Inches(1.75)
                    if "07" in img_file or "03" in img_file or "08" in img_file:
                        top = Inches(3.125)
                    if "04" in img_file or "09" in img_file or "05" in img_file or "10" in img_file:
                        top = Inches(4.5)

                    if "01" in img_file or "07" in img_file or "04" in img_file:
                        left = Inches(1.5)
                    if "06" in img_file or "03" in img_file or "09" in img_file:
                        left = Inches(4.0)
                    if "02" in img_file or "08" in img_file or "05" in img_file or "10" in img_file:
                        left = Inches(6.5)
                    try:
                        prs_slides[sld].shapes.add_picture(img_file, left, top, wd, ht)
                    except: # Add erroor image if the sample is not found
                        prs_slides[sld].shapes.add_picture(ERROR_IMG, left, top, wd, ht)

                # Update progress bar, process events so progress bar updates are seen
                bar_val += 1
                QApplication.processEvents()

        try: # Save the file and open, if the old file is open, show a message box detailing the problem
            prs.save('hw_slides.pptx')
            self.open_deck('hw_slides.pptx')
        except IOError:
            print "File already open"
            QMessageBox.warning(self, "File already open", "The hw_slides file is currently open; close this file first and try again")
        print "DONE"

        # Final updates to progress bar and output text
        self.ui.progress_bar.setValue(total)
        self.ui.text_output.setPlainText("All slides have been built!")

    # Opens the deck after completion
    def open_deck(self, file_name):
        try:
            os.system("start " + str(file_name))
        except:
            print file_name, "not found"


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle("Plastique")
    app.setWindowIcon(QtGui.QIcon("book.ico"))
    window = MainWindow()

    sip.setdestroyonexit(False)
    sys.exit(app.exec_())
