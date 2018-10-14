# Excel VBA Date Picker created by Trevor Eyre

[Download: CalendarForm+v1.5.2.zip](https://github.com/ExcelTitan/Popup/raw/master/source/CalendarForm%2Bv1.5.2.zip)  

## Overview
The goal in creating this form was first and foremost to overcome the monstrosity that is the Microsoft MonthView control. If you're reading this, you probably already know what I'm talking about. Many others have been in my place and have come up with their own date pickers to solve this problem. So why yet another custom date picker?  
 
I was most interested in the following features:  

- Ease of use. I wanted a completely self-contained form that could be imported into any VBA project and used without any additional coding.
- Simple, attractive design. While a lot of custom date pickers on the internet look good and work well, none of them quite nailed it for me in terms of style and UI design.
- Fully customizable functionality and look. I tried to include as many of the options from the MonthView control as I could, without getting too messy.  
 
Since none of the date pickers I have been able to find in all my searching have quite completed my checklist, here we are! Now my hope is that some other tired soul may also benefit from my labors.  

If you encounter any bugs, or have any great ideas or feature requests that could improve this bad boy, please send me a message.  

### Importing the Date Picker
To use the Excel VBA date picker, you must first import the userform into your project. Start by clicking the link above to download **`CalendarForm v1.5.2.zip`**. Extract the files in the zip archive, and save the **`CalendarForm.frm`** and **`CalendarForm.frx`** files on your computer.  
 
Open a new Excel workbook, and press **`alt-F11`** to open the VBA project window. Right-click on the left hand side of the project window, and select **`Import File`**... from the menu. Find the **`CalendarForm.frm`** file you saved on your computer, and click **`Open`**. This will import the date picker into your workbook, so it is ready to use.  
 
### Using the Date Picker
Once the **`CalendarForm`** date picker has been imported into your workbook, it can be used by calling the **`CalendarForm.GetDate`** function. The **`GetDate`** function is the single point of entry into the date picker. It controls everything. Every argument in the function is optional, meaning your function call an be as simple as **`dateVariable = CalendarForm.GetDate`**. That's all there is to it. The date picker initializes and pops up, the user selects a date, and the selection is returned to your variable.  

```vb
Dim dateVariable as Date
dateVariable = CalendarForm.GetDate
```
From there, you can use as many or as few arguments as you want in order to get the desired calendar that suits your needs. Below are some examples of different calendars you can get from the various arguments in the **`GetDate`** function.  
  
All default values are also set in the GetDate function, so if you want to change default colors or behavior without having to explicitly do so in every function call, you can set those values in the argument list for the function. By setting the various arguments, you can obtain vastly different calendars, all with the same function.
