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

## Source Material
Popup is an adaption of the date picker created by Trevor Eyre.  
This add-in has a lot of code added and and some modified code to suit the formatting that I prefer.  
Here is a link to Trevor's Date Picker which was v1.5.2 at time of writing:  
https://trevoreyre.com/portfolio/excel-datepicker/  

If this is deprecated you can see the source file and overview here at my Popup repository:  
https://github.com/ExcelTitan/Popup/tree/master/source

The form's functionality is completed at this stage. Some work still remains to remove unnecessary code 
