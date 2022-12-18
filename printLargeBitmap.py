import win32print
import win32ui
from PIL import Image, ImageWin, ImageShow
import math
import argparse

argParser = argparse.ArgumentParser()
argParser.add_argument("-p", "--pixelSize", type=int, required=True, help="Size of a drawing feature in pixels")
argParser.add_argument("-s", "--sizeOnPaper", type=int, required=True, help="Size of a feature as drawn on paper in inches")
args = argParser.parse_args()
#
# Constants for GetDeviceCaps
# HORZRES / VERTRES = printable area
#
HORZRES = 8
VERTRES = 10
#
# LOGPIXELS = dots per inch
#
LOGPIXELSX = 88
LOGPIXELSY = 90
#
# PHYSICALWIDTH/HEIGHT = total area
#
PHYSICALWIDTH = 110
PHYSICALHEIGHT = 111
#
# PHYSICALOFFSETX/Y = left / top margin
#
PHYSICALOFFSETX = 112
PHYSICALOFFSETY = 113

printer_name = win32print.GetDefaultPrinter ()
file_name = "c:/Users/david/Pictures/IMG_20221215_0001_half_section.png"

#
# You can only write a Device-independent bitmap
#  directly to a Windows device context; therefore
#  we need (for ease) to use the Python Imaging
#  Library to manipulate the image.
#
# Create a device context from a named printer
#  and assess the printable size of the paper.
#
hDC = win32ui.CreateDC ()
hDC.CreatePrinterDC (printer_name)
printable_area = hDC.GetDeviceCaps (HORZRES), hDC.GetDeviceCaps (VERTRES)
printer_resolution = hDC.GetDeviceCaps(LOGPIXELSX), hDC.GetDeviceCaps(LOGPIXELSY)
printer_size = hDC.GetDeviceCaps (PHYSICALWIDTH), hDC.GetDeviceCaps (PHYSICALHEIGHT)
printer_margins = hDC.GetDeviceCaps (PHYSICALOFFSETX), hDC.GetDeviceCaps (PHYSICALOFFSETY)
print("print area x= ", printable_area[0], " y= ", printable_area[1], " size x= ", printer_size[0], " y= ", printer_size[1], " margins x= ", printer_margins[0], " y= ", printer_margins[1], " resolution x= ", printer_resolution[0], " y= ", printer_resolution[1])
scalingFactor = printer_resolution[0] / (args.pixelSize / args.sizeOnPaper)
print("Scaling factor = ", scalingFactor)

#
# Open the image, rotate it if it's wider than it is high
#
bmp = Image.open(file_name)
print("Image original size (before rotation) x=", bmp.size[0], " y= ", bmp.size[1])
if bmp.size[0] > bmp.size[1]:
    print("Rotatating...")
    bmp = bmp.rotate(90, expand = 1)

print("Image original size (after rotation) x=", bmp.size[0], " y= ", bmp.size[1])
bmp = bmp.resize((int(bmp.size[0] * scalingFactor), int(bmp.size[1] * scalingFactor)))
xNumberOfSheets = math.ceil(bmp.size[0] / printable_area[0])
yNumberOfSheets = math.ceil(bmp.size[1] / printable_area[1])
print("Number of sheets: x= ", xNumberOfSheets, " y= ", yNumberOfSheets)

for x in range(0, xNumberOfSheets):
    for y in range(0, yNumberOfSheets):
        subBmp = bmp.crop((x * printable_area[0], y * printable_area[1], (x * printable_area[0]) + printable_area[0] - 1, (y * printable_area[1]) + printable_area[1] - 1))
        #ImageShow.show(subBmp)
        
        #
        # Start the print job, and draw the bitmap to
        #  the printer device at the scaled size.
        #
        hDC.StartDoc ("temp" + str(x) + str(y))
        hDC.StartPage ()

        dib = ImageWin.Dib (subBmp)
        dib.draw (hDC.GetHandleOutput (), (printer_margins[0], printer_margins[1], printer_margins[0] + printable_area[0], printer_margins[1] + printable_area[1]))

        hDC.EndPage ()
        hDC.EndDoc ()
hDC.DeleteDC ()
