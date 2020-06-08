
# CAEN InventoryGUI
Semi-automation of the CAEN laptop inventory process

* Removes human error of misreading/mistyping Serial# and LAN MAC Address for laptops
* Semi-automatically and near-instantaneously inputs all required information for adding inventory

Runs on Python and Javascript, compatible with Windows.

# Setup (Windows)
1. Download the InventoryGUI_exe.zip file from the releases tab or within the repo
2. Extract the files to a folder and run inventoryGUI.exe
3. Done

# Setup (Python)
1. Install [Python](https://www.python.org/downloads/) and add to PATH
2. Download ```InventoryGUI.py```, ```fields.txt```, and ```dropdowns.txt``` to the same folder
3. Install [Microsoft Visual C++ Redistributable Package 2013 (x64)](https://www.microsoft.com/en-us/download/details.aspx?id=40784)
4. Using pip, install the modules in ```requirements.txt```
5. Run ```python inventoryGUI.py```
6. Done

# Usage (Windows)
1. Take one picture of each laptop box label, clearly displaying the Serial# and LAN MAC Address barcodes
2. Transfer pictures to computer in a single folder
3. Run ```inventoryGUI.exe```
4. When the GUI opens up, select the image folder from Step 2 and fill out/edit each field
	> Note: It is best to pull the exact spelling and formatting from the CAEN inventory website
5. Click the "Start" button. This will create:
	1. An Excel spreadsheet with the Serial#s, LAN MAC Addresses, and (optional) Names for each laptop
	2. A text file with the Javascript commands for each laptop
	3. An Excel spreadsheet with the Javascript commands for each laptop
	> Note: it doesn't matter which of ii and iii you use, just a matter of personal preference
6. Open the CAEN inventory website and "Add Inventory"
7. Right click, then click "Inspect Element" (or press Ctrl Shift I)
8. Click on "Console"
9. From 5.ii or 5.iii, copy the respective script, paste into the Console, and hit the "Enter" key. This should automatically fill in the required fields, save, and clone to a new inventory entry
10. Repeat Step 9 for all laptop inventory scripts generated

# Usage (Python)
1. Take one picture of each laptop box label, clearly displaying the Serial# and LAN MAC Address barcodes
2. Transfer pictures to computer in a single folder
3. In CMD, navigate to inventoryGUI's folder and run ```python inventoryGUI.py```
4. When the GUI opens up, select the image folder from Step 2 and fill out/edit each field
	> Note: the default values can be easily edited in inventoryGUI.py. It is best to pull the exact spelling and formatting from the CAEN inventory website
5. Click the "Start" button. This will create:
	1. An Excel spreadsheet with the Serial#s, LAN MAC Addresses, and (optional) Names for each laptop
	2. A text file with the Javascript commands for each laptop
	3. An Excel spreadsheet with the Javascript commands for each laptop
	> Note: it doesn't matter which of ii and iii you use, just a matter of personal preference
6. Open the CAEN inventory website and "Add Inventory"
7. Right click, then click "Inspect Element" (or press Ctrl Shift I)
8. Click on "Console"
9. From 5.ii or 5.iii, copy the respective script, paste into the Console, and hit the "Enter" key. This should automatically fill in the required fields, save, and clone to a new inventory entry
10. Repeat Step 9 for all laptop inventory scripts generated
