import openpyxl
from PIL import Image

def image_to_excel(file, output, percentage, zoom):

    # Open picture and create a workbook instance
    im = Image.open(file)
    wb = openpyxl.Workbook()
    sheet = wb.active

    # Resize image with spreadsheet's columns
    width, height = im.size
    cols = width * percentage/100
    print('Old image size: ' + str(im.size))
    imgScale = cols/width
    newSize = (int(width*imgScale), int(height*imgScale))
    im = im.resize(newSize)

    # Get new picture's dimensions
    cols, rows = im.size
    print('New image size: ' + str(im.size))

    # Formatting spreadsheet's cell size: height = 0.6 and width = 0.0625 (1 pixel)
    zoom = 100
    for i in range(1, rows):
        sheet.row_dimensions[i].height = zoom*0.6/100
    for j in range(1, cols):
        column_letter = openpyxl.utils.get_column_letter(j)
        sheet.column_dimensions[column_letter].width = zoom*0.0625/100

    # Convert image to RGB
    rgb_im = im.convert('RGB')

    # Formatting cell's color 
    for i in range(1, rows):
        for j in range(1, cols):
            c = rgb_im.getpixel((j, i))
            rgb2hex = lambda r,g,b: f"{r:02x}{g:02x}{b:02x}"
            c = rgb2hex(*c)
            sheet.cell(row = i, column = j).value = " "
            customFill = openpyxl.styles.PatternFill(start_color=c, end_color=c, fill_type='solid')
            sheet.cell(row = i, column = j).fill = customFill

    # Save workbook
    wb.save(output) 
    wb.close()
    

# Export (input file, output file, image scale (%), zoom(%))
image_to_excel('jangada.jpg', 'jangada.xlsx', 100, 100)