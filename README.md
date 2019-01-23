# Popup - An Excel Add-in

A beautiful and simple calendar widget that does all that you want it to, easily input dates into your spreadsheet.  

## How it Works
### Open and Close
Open the form by clicking the button from the ribbon on the Home tab or the button in the context menu.   
Keyboard shortcut: **`Ctrl + Shift + D`** to open straight from the spreadsheet.  
You can then quickly use the Escape key **`{Esc}`** to close the form.  

### Function
A single click on any date will populate all of the selected cells with that value.

A *single click* on the **`Today`** label at the bottom of the form will take you back to today's month and select today's date.  
*Double clicking* on the **`Today`** label will input today's date to the selected cells.

If the active cell is blank when the form is opened, Popup's default date will be today's month, highlighted light blue.  
If the cell already contains a date then Popup will show that date's month and year, highlighting the date light yellow.

### Navigation
You can use the **`Left`** and **`Right`** arrows to slowly go to the previous or next month, to go a little bit quicker, use the scroll wheel on your mouse.   
Clicking on the **`Month`** or **`Year`** label at the top of the form will display all the Months of the year and a whole heap of Years to choose form.  

----

## Browse for the Addin Folder
To manually locate the default Excel AddIns folder, follow the steps below.

- Click the Developer tab on the Excel Ribbon. If it isn't visible, follow the steps here.
- Click the Add-Ins command
- In the Add-Ins window, select any add-in in the list, and click the Browse button. That will open the Browse window, at the AddIns folder.
- Right-click on the path at the top of the Browse window, and click "Copy Address as Text"
- Click Cancel, to close the Browse window
- Click Cancel, to close the Add-Ins window.
-Open Windows Explorer, and paste the copied address into the address bar, then press Enter

## Installation
- Unzip the zip file to extract the contents
- Move the add-in to your local Microsoft AddIns directory.
- Open Excel
- Click the File tab
- Click Options
- Click the Add-Ins category on the left bar of the Excel Options Dialog Box
- In the Manage box near the bottom, click Excel Add-ins, and then click Go. The Add-Ins dialog box appears.
- Click Browse
- Navigate to the xla or xlam add-in and select it.
- Make sure the checkbox next to the add-in is checked, then click OK.

If the add-in keeps disappearing when you restart Excel, follow the directions below.  
***Important Note:*** There are a few additional steps to this installation due to an Office Security Update released in July 2016.


#### Trust the Folder Location
The folder that the add-in file is saved in needs to be added as a Trusted Location in Excel. The instructions on how to trust the folder location are below.  
***Open the Excel Options menu:***
- File > Options
- Open the Trust Center menu and Add a new location
- Trust Center > Trust Center Settings… > Trusted Locations > Add new location
- Browse for the folder that you saved the add-in file in.
- Press OK, the folder should now appear in the Trusted Locations list.
- Press OK on the Trust Center and Excel Options menu to close the menus.  
 
The add-in is now in a trusted location and the add-in will appear every time you open Excel.  


#### Unblock the File
Some users still have issues with the add-in’s ribbon disappearing after trusting the folder location. If this happens to you, you will need to Unblock the file by changing a file property.  

- Locate the Add-in file (.xla, .xlam) in Windows Explorer.
- Right-click the file and select Properties.
- At the bottom of the General tab you should see a Security section. Check the box that says Unblock.
- Press the OK button.

Close Excel completely and re-open it. The add-in should now load and any custom ribbons will appear.  

You will only need to do this unblock one time. However, if you download an updated version of the file then you will have to repeat the steps above to unblock it.  
 
Hopefully these additional security steps will be fixed in a future update to Office. If you completed all the steps above then you should see the add-ins ribbon tab load every time you open Excel.

----  

## Source Material
Popup is an adaption of the date picker created by Trevor Eyre.  
This add-in has a lot of code added and and some modified code to suit the formatting that I prefer.  
Here is a link to Trevor's Date Picker which was v1.5.2 at time of writing:  
https://trevoreyre.com/portfolio/excel-datepicker/  

If this is deprecated you can see the source file and overview here at my Popup repository:  
https://github.com/ExcelTitan/Popup/tree/master/source

The form's functionality is completed at this stage. Some work still remains to remove unnecessary code 
