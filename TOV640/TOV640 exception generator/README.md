# TOV640 Exception Generator

The **TOV640_exception_generator** takes in a .datac file which stores the collected data for some section of the line EAL or TML, and outputs a .xlsx file which tells the categories (wire wear, low height, high height, left stagger, right stagger) and the location of the line (in meters) that fail to meet the standard.

## Before Using the Program
1. Make sure the naming of the metadata files are correct and they are stored in the ```metadata``` folder.
    - ```./metadata/EAL metadata.xlsx```
    - ```./metadata/TML metadata.xlsx```
2. Make sure the unit for location in the metadata is in meters
3. Install Python pandas
    - Enter ```pip install pandas``` in your terminal
4. Install openpyxl
    - Enter ```pip install openpyxl``` in your terminal

## How To Use the Program
### Method 1
1. Run the Python Script by clicking the **Run** button or enter ```python TOV640\ exception\ generator/TOV640_exception_generator.py``` in the terminal (assuming you are in the parent directory of ```TOV640 exception generator```)
2. Click **ENTER** four times to skip the instructions.
3. Enter either ```EAL``` or ```TML``` in capital letters.
    - If entering ```EAL```, enter ```y``` or ```n``` in lowercase letters to the questions asking if the data is in the **LMC**, **RAC**, **LOW S1**, or **SCL** section.
    - Otherwise, enter the section range of the data, e.g. ```UNI-TAP```.
4. Enter ```UP``` or ```DN``` in capital letters for up or down track.
5. Select the .datac file in your directory to be analyzed.
6. When asked, press ```ENTER``` to select the location to save the .xlsx file.

### Method 2
1. In auto_input.py, change the input values in the list ```input```.
2. Run auto_input.py. When prompted, choose the .datac file to be analyzed and the location to save the .xlsx file.