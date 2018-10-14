# Popup - An Excel Add-in
A beautiful calendar widget for Excel

[Download Popup.xlam](https://github.com/ExcelTitan/Popup/raw/master/Popup.xlam){: .btn .btn--info}

## What is it?
A beautiful and simple calendar widget that does all that you want it to, easily input dates into your spreadsheet.  

## How it Works
Open the form by clicking the button from the ribbon on the Home tab or the button in the context menu.

To Navigate through **Months** and **Years** you can click the left and right arrows, choose the month or year by clicking on the label at the top of the form. You can also scroll the mouse wheel to scroll through the months really fast.  
Clicking on 'Today' or the date at the bottom of the form will take you back to today's month and select today's date.  
You can click on the month or year then choose from the drop downs provided.

A single click on date will input to the current selection, it's for this reason that the form is modal (you can't click on another cell once it's opened).
You can also double click on the Today label at the bottom to put today's date into the currently selected cells.

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


## Trust the Folder Location
The folder that the add-in file is saved in needs to be added as a Trusted Location in Excel. The instructions on how to trust the folder location are below.
-  Open the Excel Options menu
   - File > Options
   - Open the Trust Center menu and Add a new location
   - Trust Center > Trust Center Settings… > Trusted Locations > Add new location
   - Browse for the folder that you saved the add-in file in.
   - Press OK, the folder should now appear in the Trusted Locations list.
   - Press OK on the Trust Center and Excel Options menu to close the menus.  
 
The add-in is now in a trusted location and the add-in will appear every time you open Excel.


## UNBLOCK THE FILE
Some users still have issues with the add-in’s ribbon disappearing after trusting the folder location. If this happens to you, you will need to Unblock the file by changing a file property.  

- Locate the Add-in file (.xla, .xlam) in Windows Explorer.
- Right-click the file and select Properties.
- At the bottom of the General tab you should see a Security section. Check the box that says Unblock.
- Press the OK button.

Close Excel completely and re-open it. The add-in should now load and any custom ribbons will appear.  

You will only need to do this unblock one time. However, if you download an updated version of the file then you will have to repeat the steps above to unblock it.  
 
Hopefully these additional security steps will be fixed in a future update to Office. If you completed all the steps above then you should see the add-ins ribbon tab load every time you open Excel.

## Source Material
Popup is an adaption of the date picker created by Trevor Eyre.  
This add-in has a lot of code added and and some modified code to suit the formatting that I prefer.  
Here is a link to Trevor's Date Picker which was v1.5.2 at time of writing:  
https://trevoreyre.com/portfolio/excel-datepicker/  

If this is deprecated you can see the source file and overview here at my Popup repository:  
https://github.com/ExcelTitan/Popup/tree/master/source

The form's functionality is completed at this stage. Some work still remains to remove unnecessary code 
