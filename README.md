# ExcelArt
Convert an image to excel colored cells or create an animation with excel colored cells from a GIF file. 

![imgonline-com-ua-twotoone-j1XV8OoC7hpBQ](https://user-images.githubusercontent.com/56649205/81574966-dd861080-937c-11ea-94d5-1158b6c81acb.jpg)

## Requirements
* [Pillow](https://pillow.readthedocs.io/en/stable/)
* [Openpyxl](https://openpyxl.readthedocs.io/en/stable/)


## How to use
### Create colored cells
Just run the function:
```Python
# Export (input file, output file, image scale (%), zoom(%))
image_to_excel('muitos.jpg', 'muitos.xlsx', 100, 100)
```
- Image scale (%): change the image size (100 is default)
- Zoom (%): excel's zoom (100 is default)

### Animation with colored cells
1) Select your preference GIF file and output folder:
```Python
# gif_to_img(file.gif, output_folder_name, step_frames) 
gif_to_img('nazare.gif', 'nazare_photos', 2)
```

2) Change *nazare_photos* to your output folder to get all the frame images:
```Python
images = glob.glob("nazare_photos/*.jpg")
gif_to_excel(images, 'nazare.xlsx', 100, 100)
```

3) Open the excel file (like *nazare.xlsx* in this example), click `ALT+F11` to create an excel macro, then copy and paste the code:
```VBA
Sub LoopSheets() 'VBA macro to Loop through sheets and sorts Cols A:E in Ascending order.
Dim ws As Worksheet

Do While k < 100

For Each ws In Sheets 'Start of the VBA looping procedure.
ws.Activate
Application.Wait (Now + 0.00001)
Next ws
Sheets(1).Activate
k = k + 1
Loop

End Sub
```

## Limitation
On large images, excel's */xl/styles.xml part (Styles)* can be corrupted. I think this limitation is relative to the number of different colours on the image. For more info, look at [Excel specifications](https://support.office.com/en-ie/article/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3).

This code is a small modification of [this one](https://github.com/joelibaceta/pix-to-xls). 
