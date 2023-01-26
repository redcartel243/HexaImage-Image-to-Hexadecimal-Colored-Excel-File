# HexaImage-Image-to-Hexadecimal-Colored-Excel-File
## **Takes images and converts them into hexadecimal colored excel files**
This script is a Python script that converts images to hexadecimal colored excel files. It uses the openpyxl library to create an excel file and the PIL library to open and manipulate the image.

The script starts by prompting the user to input the number of images they would like to process. It then prompts the user to input the path of each image and the names of the result files without color and with color.

The script then reads and resizes the image using PIL, converts the image into an array object, and converts the R,G,B values of the image into hexadecimal values. These hexadecimal values are then populated in a DataFrame and saved as an excel file.

After saving the excel file, the script opens the excel file and colors each cell with the corresponding hexadecimal value of the image. The final excel file with color is then saved.

The script can be useful for creating color palettes, performing image analysis, compressing images, training machine learning models for image recognition tasks, color matching, and data visualization.

To run the script, you will need to have Python, openpyxl, and PIL library installed. Please make sure to run the script in a virtual environment and run the command "pip install -r requirements.txt" to install the necessary dependencies.

Note that this script will only run on .jpg, .jpeg, .png and .bmp files.

## **Results**

**Image before conversion**
<img src='https://github.com/redcartel243/HexaImage-Image-to-Hexadecimal-Colored-Excel-File/blob/main/Arch2O-record-breaker-shanghai-tower-the-worlds-second-tallest-building-completed-in-shanghai.jpg'/>


**Hexadecimal State(See NoColor0.xlsx)**


**Colored State(see Color0.xlsx)**

