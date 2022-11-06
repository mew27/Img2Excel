
from PIL import Image
import numpy
import xlsxwriter
from tqdm import tqdm


im = Image.open("a1ff76b04fd4d2bfcca0285e4c081021.jpeg")
imArray = numpy.asarray(im)

workbook = xlsxwriter.Workbook('converted.xlsx')
worksheet = workbook.add_worksheet()


for i in tqdm(range(0, imArray.shape[0], 10)):
    worksheet.set_row_pixels(i, 10)
    for j in range(0, imArray.shape[1], 10):
        r = imArray[i][j][0]
        g = imArray[i][j][1]
        b = imArray[i][j][2]

        rgb_string = f"#{r:02x}{g:02x}{b:02x}".upper()
        rgb_format = workbook.add_format()
        rgb_format.set_bg_color(rgb_string)

        #rgb_format = findRGBPalette(pltte, r, g, b)
        worksheet.write(i//10, j//10, '', rgb_format)

worksheet.set_column_pixels(0, imArray.shape[1] // 10 + 1, 20)

workbook.close()