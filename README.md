# ExcelArt
Convert an image to excel colored cells

![imgonline-com-ua-twotoone-j1XV8OoC7hpBQ](https://user-images.githubusercontent.com/56649205/81574966-dd861080-937c-11ea-94d5-1158b6c81acb.jpg)

## Requirements
* [Pillow](https://pillow.readthedocs.io/en/stable/)
* [Openpyxl](https://openpyxl.readthedocs.io/en/stable/)


## How to use
Just run the function:
```Python
# Export (input file, output file, image scale (%), zoom(%))
image_to_excel('muitos.jpg', 'muitos.xlsx', 100, 100)
```
- Image scale (%): change the image size (100 is default)
- Zoom (%): excel's zoom (100 is default)

## Limitation
On large images, excel's */xl/styles.xml part (Styles)* can be corrupted.
This code is a small modification of [this one](https://github.com/joelibaceta/pix-to-xls). 
