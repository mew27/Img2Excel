
from PIL import Image
import numpy
import xlsxwriter
from tqdm import tqdm
import sys

if len(sys.argv) < 3:
    print("Usage: python Img2Excel.py image outputfile [ratio]")
    exit()

image_name = sys.argv[1]
outputfile = sys.argv[2]

if len(sys.argv) >= 4:
    try:
        ratio = int(sys.argv[3])
        print(f"Converting pic with a ratio of 1:{ratio}")

        if ratio < 1:
            print("Ratio should be a number bigger than 1")
            exit()
    except:
        print("Ratio should be a number bigger than 1")
        exit()
else:
    ratio = 10

im = Image.open(image_name)
imArray = numpy.asarray(im)

workbook = xlsxwriter.Workbook(outputfile)
worksheet = workbook.add_worksheet()


for i in tqdm(range(0, imArray.shape[0], ratio)):
    worksheet.set_row_pixels(i, 10)
    for j in range(0, imArray.shape[1], ratio):
        r = imArray[i][j][0]
        g = imArray[i][j][1]
        b = imArray[i][j][2]

        rgb_string = f"#{r:02x}{g:02x}{b:02x}".upper()
        rgb_format = workbook.add_format()
        rgb_format.set_bg_color(rgb_string)

        #rgb_format = findRGBPalette(pltte, r, g, b)
        worksheet.write(i//ratio, j//ratio, '', rgb_format)

worksheet.set_column_pixels(0, imArray.shape[1] // ratio + 1, 20)

print("Writing to file...")
workbook.close()
print("Done!")