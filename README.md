# PicToExel
## Convert a table image to an Excel file.
The user selects an image of a table (.jpeg / png) and the program converts it to an Excel file (.xlsx).
Use of libraries: tkinter, pandas, opencv.
Tkinter: creating the GUI.
Pandas: handling data and storing it in an Excel file.
Opencv: image processing.\
Summary of the project:\
1. Importing all necessary libraries, including tkinter, pandas, OpenCV and other libraries.
2. Creating a screen - Tkinter's user interface and designing a screen so that the screen includes:\
• Button that opens a dialog box for selecting an image.\
• A GIF used as a loading animation while the image is being processed.\
• Button to save the prepared Excel cube.\
  3. Functionality to select a file, save it as a JPEG and display it on the screen.
4. The code reads the image and converts it to a binary image, and then it detects the horizontal and vertical lines in the image.
  5. The code identifies all the bounding boxes for each cell in the table and by its row and column.
  6. The OCR is performed on each of the cells using pytesseract, and the output is stored in a list.
  7. The data is then stored in a Pandas DataFrame and written to an Excel file with a save button to save the file, which opens a dialog to save the file to a selected directory.
The project is documented in code - for each function, and a code block.
You can also see the steps of the execution by pictures that were created, and saved in the pics folder, for example:

Choose the image:\
![image](https://github.com/TamarShayo/PicToExel/assets/120498576/91746a67-2dde-4672-ab24-5af700bfd224)

Converting the image to binary:\
![image](https://github.com/TamarShayo/PicToExel/assets/120498576/9e83154d-b344-4feb-86ff-c3f638ae6f8a)

vertical lines:\
![image](https://github.com/TamarShayo/PicToExel/assets/120498576/82d75c01-842f-43a2-a3eb-b60a5cf71b4c)

horizon lines:\
![image](https://github.com/TamarShayo/PicToExel/assets/120498576/1a36e6f1-8b96-4672-b2cd-b96559e89404)

Combining the lines:\
![image](https://github.com/TamarShayo/PicToExel/assets/120498576/567f87ab-6619-49e4-a802-73d162a62f41)

Created the output.xlsx file:\
![image](https://github.com/TamarShayo/PicToExel/assets/120498576/2ef44ecc-474a-429e-bb58-eeba24de1a28)
