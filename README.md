# Check-BOM
Check through a Solid Works bill of materials and see if any of the part numbers are miss matched with the file names

**Installation:** 

Open any drawing.
Right click on the toolbar, then select “Customize…”

Click on the “Commands” tab, select Macro, and then Drag and Drop the last icon (“New Macro Button”) onto the empty space on the toolbar.

Click on the 3 dots, navigate to “E:\WORK\Outdwg\SolidworksBomOutputs” then select the file called “ExportBOMasExel.swp”. Click “Open”.

Give it a name (or leave it as is). Click “OK” on the “Customize Macro Button” page. And click “OK” again on the “Customize” page.

Your macro is now available on your toolbar.



**How to use:**

When on an assembly drawing, click the macro and then click “OK”.

Wait for a few seconds until the Excel Sheet has popped up.

Click on the Excel application, and press “Ctrl + r ”.

Wait another couple seconds until the macro has finished running.

The Excel sheet should look like this. The first section is the original BOM. The second section shows all of the discrepancies between the “Title” and “File Name”. The third section shows only the discrepancies whose Title is a number and does not contain any letters.

You will find a copy of the third section on “Sheet 2”. Note: the second sheet only shows the top 45 discrepancies to allow for a one-sheet print. If you want more, you can copy, then “special paste” the rest of the values onto the second page.



When done, close the Excel file first, and then select whether you want or do not want to delete the Excel file.
 
**Note**: do not click “Yes” before closing the Excel file. This will cause the system to give you an error due to a lack of permission.
