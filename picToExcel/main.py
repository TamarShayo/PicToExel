# Import necessary libraries:
# This code block imports various libraries such as tkinter, opencv, pandas and PIL.
# These libraries are necessary for creating the GUI, image processing and converting the image to Excel.
from tkinter import *
from tkinter import ttk
from tkinter import filedialog,messagebox
import cv2
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import csv
import pytesseract
from PIL import Image, ImageTk, ImageSequence
import os
import shutil

# function to select an image:
# This function allows the user to select an image file from their computer.
# The selected image is saved as table.jpg in the pics folder.
# The function also displays the selected image in the GUI and calls the loud() function to display the GIF.
def select_file():
    path = filedialog.askopenfilename(title="Select an Image", filetype=(('image files','*.jpg'),('all files','*.*')))
    img = Image.open(path)
    file_name = f'{os.path.split(path)[-1]}'
    img.save(f'pics/table.jpg')
    img.thumbnail((200,200))
    img = ImageTk.PhotoImage(img)
    loud()
    label = Label(win, image=img)
    label.image = img
    label.pack()
    read_file()

# Function to refresh the image (GIF):
# This function displays a GIF in the GUI. It reads the GIF file and converts it into individual frames.
# It then displays the first frame of the GIF in the GUI and calls the animate_gif() function to display the remaining frames of the GIF.
# Finally, it calls the read_file() function to convert the selected image to Excel.
def loud():
   gif = Image.open("pics/Rounded_stripes.gif")
   frames = []
   for frame in ImageSequence.Iterator(gif):
      frames.append(ImageTk.PhotoImage(frame))
   label = Label(win, image=frames[0],bd=0)
   label.pack()

   def animate_gif(frame):
      label.config(image=frames[frame])
      win.after(50, animate_gif, (frame + 1) % len(frames))

   animate_gif(0)

# Reading the selected file and converting it to Excel:
# This code defines a function called read_file which reads an image file (table.jpg) and performs several image processing operations on it.
# First, the image is converted to a binary image using thresholding and then inverted.
# Next, vertical and horizontal kernels are defined to detect the vertical and horizontal lines in the image, respectively.
# These kernels are then used to detect the vertical and horizontal lines in the image
# and save them as separate images (vertical.jpg and horizontal.jpg).
# The vertical and horizontal lines are then combined in a new image (img_vh) using the cv2.addWeighted function.
# The resulting image is then eroded and thresholded to obtain a binary image (img_vh.jpg).
# Bitwise XOR and NOT operations are then performed on the original grayscale image and the img_vh image.
# Finally, contours are detected in the img_vh image using the cv2.findContours function.
def read_file():
    file = r'pics/table.jpg'
    img = cv2.imread(file, 0)
    # thresholding the image to a binary image
    thresh, img_bin = cv2.threshold(img, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    # inverting the image to binary image
    img_bin = 255 - img_bin
    cv2.imwrite('pics/cv_inverted.png', img_bin)
    #The definition of the horizontal vertical lines of the table:
    kernel_len = np.array(img).shape[1] // 100
    # Defining a vertical kernel to detect all vertical lines of image
    ver_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, kernel_len))
    # Defining a horizontal kernel to detect all horizontal lines of image
    hor_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (kernel_len, 1))
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))

    # Use vertical kernel to detect and save the vertical lines in a jpg
    image_1 = cv2.erode(img_bin, ver_kernel, iterations=3)
    vertical_lines = cv2.dilate(image_1, ver_kernel, iterations=3)
    cv2.imwrite("pics/vertical.jpg", vertical_lines)

    # Use horizontal kernel to detect and save the horizontal lines in a jpg
    image_2 = cv2.erode(img_bin, hor_kernel, iterations=3)
    horizontal_lines = cv2.dilate(image_2, hor_kernel, iterations=3)
    cv2.imwrite("pics/horizontal.jpg", horizontal_lines)

    # Combine horizontal and vertical lines in a new third image, with both having same weight.
    img_vh = cv2.addWeighted(vertical_lines, 0.5, horizontal_lines, 0.5, 0.0)
    # Eroding and thesholding the image
    img_vh = cv2.erode(~img_vh, kernel, iterations=2)
    thresh, img_vh = cv2.threshold(img_vh, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    cv2.imwrite("pics/img_vh.jpg", img_vh)
    bitxor = cv2.bitwise_xor(img, img_vh)
    bitnot = cv2.bitwise_not(bitxor)

    # Detect contours for following box detection
    contours, hierarchy = cv2.findContours(img_vh, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

    #This function receives the pointer to the list of lines that identify the objects in the image
    # and arranges them in a certain order according to the method parameter.
    # The function uses the method parameter to decide whether to sort the lines from left to right, right to left, top to bottom or bottom to top.
    # The implementation of the function uses operations like cv2.
    # boundingRect to find the bounding boxes of the objects and sort them according to the method parameter.
    # Finally, the function returns the ordered list of the lines and their bounding boxes.
    def sort_contours(cnts, method="left-to-right"):
        # initialize the reverse flag and sort index
        reverse = False
        i = 0
        # handle if we need to sort in reverse
        if method == "right-to-left" or method == "bottom-to-top":
            reverse = True
        # handle if we are sorting against the y-coordinate rather than
        # the x-coordinate of the bounding box
        if method == "top-to-bottom" or method == "bottom-to-top":
            i = 1
        # construct the list of bounding boxes and sort them from top to
        # bottom
        boundingBoxes = [cv2.boundingRect(c) for c in cnts]
        (cnts, boundingBoxes) = zip(*sorted(zip(cnts, boundingBoxes),
                                            key=lambda b: b[1][i], reverse=reverse))
        # return the list of sorted contours and bounding boxes
        return (cnts, boundingBoxes)

    # Sort all the contours by top to bottom.
    contours, boundingBoxes = sort_contours(contours, method="top-to-bottom")
    # Creating a list of heights for all detected boxes
    heights = [boundingBoxes[i][3] for i in range(len(boundingBoxes))]
    # Get mean of heights
    mean = np.mean(heights)
    # Create list box to store all boxes in
    box = []
    # Get position (x,y), width and height for every contour and show the contour on image
    for c in contours:
        x, y, w, h = cv2.boundingRect(c)
        if (w < 1000 and h < 500):
            image = cv2.rectangle(img, (x, y), (x + w, y + h), (0, 255, 0), 2)
            box.append([x, y, w, h])

    # Creating two lists to define row and column in which cell is located
    row = []
    column = []
    j = 0

    # Sorting the boxes to their respective row and column
    for i in range(len(box)):
        if (i == 0):
            column.append(box[i])
            previous = box[i]
        else:
            if (box[i][1] <= previous[1] + mean / 2):
                column.append(box[i])
                previous = box[i]
                if (i == len(box) - 1):
                    row.append(column)
            else:
                row.append(column)
                column = []
                previous = box[i]
                column.append(box[i])
    print(column)
    print(row)

    # calculating maximum number of cells
    countcol = 0
    for i in range(len(row)):
        countcol = len(row[i])
        if countcol > countcol:
            countcol = countcol

    # Retrieving the center of each column
    center = [int(row[i][j][0] + row[i][j][2] / 2) for j in range(len(row[i])) if row[0]]
    center = np.array(center)
    center.sort()
    print(center)
    # Regarding the distance to the columns center, the boxes are arranged in respective order
    # This code is used to extract the text from each of the table cells and create a new Excel file containing the data in the table.
    # The following actions are performed:
    # - Creates an empty list called finalboxes.
    # - Performs a double loop that goes through all the cells in the table and splits the cells by columns.
    # - uses the pytesseract library to read the text from each cell in the table.
    # - Creates a new list named outer that contains the text read from each cell in the table, where each cell is represented as a string.
    # - uses the pandas library to create a new dataframe that contains the data found in the table.
    # - Displays the new frame on the screen and saves it as a new Excel file called output.xlsx.
    finalboxes = []
    for i in range(len(row)):
        lis = []
        for k in range(countcol):
            lis.append([])
        for j in range(len(row[i])):
            diff = abs(center - (row[i][j][0] + row[i][j][2] / 4))
            minimum = min(diff)
            indexing = list(diff).index(minimum)
            lis[indexing].append(row[i][j])
        finalboxes.append(lis)

    # from every single image-based cell/box the strings are extracted via pytesseract and stored in a list
    outer = []
    for i in range(len(finalboxes)):
        for j in range(len(finalboxes[i])):
            inner = ''
            if (len(finalboxes[i][j]) == 0):
                outer.append(' ')
            else:
                for k in range(len(finalboxes[i][j])):
                    y, x, w, h = finalboxes[i][j][k][0], finalboxes[i][j][k][1], finalboxes[i][j][k][2], \
                                 finalboxes[i][j][k][3]
                    finalimg = bitnot[x:x + h, y:y + w]
                    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 1))
                    border = cv2.copyMakeBorder(finalimg, 2, 2, 2, 2, cv2.BORDER_CONSTANT, value=[255, 255])
                    resizing = cv2.resize(border, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
                    dilation = cv2.dilate(resizing, kernel, iterations=1)
                    erosion = cv2.erode(dilation, kernel, iterations=2)
                    out = pytesseract.image_to_string(erosion)
                    if (len(out) == 0):
                        out = pytesseract.image_to_string(erosion, config='--psm 3')
                    inner = inner + " " + out
                outer.append(inner)

    # Creating a dataframe of the generated OCR list
    arr = np.array(outer)
    dataframe = pd.DataFrame(arr.reshape(len(row), countcol))
    print(dataframe)
    data = dataframe.style.set_properties(align="left")
    # Converting it in a excel-file
    data.to_excel("pics/output.xlsx")
    save_button = ttk.Button(win, text="Save Excel file", command=save_file)
    save_button.pack()
def save_file():
    root = filedialog.askdirectory(initialdir="/", title="Select directory to save file")
    # if a directory is selected perform file saving operation here
    if root:
        # copy the Excel file to the selected directory with the name "Output.xlsx"
        shutil.copy("pics/output.xlsx", root + "/Output.xlsx")
        messagebox.showinfo("Success", "File saved successfully")

# Creating the GUI:
win = Tk()
win.geometry("1000x700")
win.title("picToExcel")
win.configure(bg="white")
# Creating a label and a button to open the dialog box:
# This code block creates a label and button in the GUI.
# When the button is clicked, the select_file() function is called.
Label(win, text=" convert image to excel file", font=('Caveat 20 bold')).pack(pady=20)
logoimg = ImageTk.PhotoImage(Image.open("pics/logo.png"))
label = Label(win, image=logoimg,bd=0)
label.pack()
style = ttk.Style()
style.configure("Custom.TButton", background="black", bordercolor=(0, 0, 0), padding=10)
button = ttk.Button(win, text="Select an image 👉", command=select_file, style="Custom.TButton")
button.pack(ipadx=5, pady=15)

#run the GUI
win.mainloop()
