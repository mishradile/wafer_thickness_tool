# Wafer Thickness Plotter
For plotting wafer thickness given CSV files

## Installation 
- Download the entire `wafer_thickness_tool` folder onto Desktop. If you're downloading from Github, click the green `Code` button on top of the list of files, and click `Download Zip`. Move the downloaded folder onto your Desktop
- You need to have Python version>=3.9 to run the script. 
- Check if Python is installed in your computer:
    - Open command line. In Windows, press Windows logo + S to launch search window, and search `cmd`. 
    - In command line, type: `python --version`
    - If Python has been installed, the version that is in your computer will be returned. 
![Check python ver](https://user-images.githubusercontent.com/105037297/169487975-c7da6c6f-da46-44d2-bda3-5d8dd35987d7.PNG)
    - If an error message is returned, Python has not been installed. You can download it here: https://www.python.org/downloads/
- If Python has been installed in your computer, you need to proceed to install the required Python libraries that is needed to run the script. In the command line, type `cd %HOMEPATH%/Desktop/wafer_thickness_tool/` and press Enter. After that, type `pip install -r requirements.txt` and wait for installation to be completed.
- If no error messages appeared on the terminal, you have successfully installed the packages needed to run the script, please refer to Usage section below on how to run the script. 

## Troubleshooting
- If after entering `pip install -r requirements.txt` there's an error message `'pip' is not recognized as an internal or external command`, try the following:
  - Download Python again (or clicking the existing installation exe file which has been used to download Python and then clicking “Customize installation”), and checking that the “pip” checkbox is ticked to set the path for pip.
![pythoninstall](https://user-images.githubusercontent.com/105037297/175252187-2681279f-16b6-4e63-a583-1d06caa34270.PNG)
  - If the above doesn't work, you can set the PATH element from Window's cmd line: enter `setx PATH "%PATH%;C:\Python34\Scripts”`. After hitting enter, close the current terminal. Open a new terminal and try running `pip install -r requirements.txt` again. 
- If after running the script, terminal shows an error `ValueError: Points cannot contain NaN`, please check that all the data are flushed to the top left corner of each Excel sheets.

## Usage
Video demo of code usage: 
https://user-images.githubusercontent.com/105037297/177476376-20421ca0-2e59-4b12-9ef8-6c4915e5fc17.mp4
Note that if any data points have values/coordinates that are non-numeric or empty in the Excel sheet, it will be ignored by the program, and the wafer map will be generated using the remaining data. A message will be printed to the terminal after the process is completed, listing the worksheets which have empty cells/erroneous values. Please also ensure that data are flushed to the top left corner of each Excel worksheet to avoid generating the error message. 
![sd](https://user-images.githubusercontent.com/105037297/178438790-266d10e0-b4e9-4c01-be8c-288b52b63ef8.PNG)




