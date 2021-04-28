# Creating a PowerPoint slides from Excel data

The purpose of this code is to automate the creation of PowerPoint using the large amount of data available in Excel.

You can create more columns in Excel, but you will need to add more shapes in PowerPoint and change the code part, as shown below:

```bash
AddS.Shapes(<number of shape that u wanna put the title>).TextFrame.TextRange = DataRow.Cells(2, 1)
        For i = 1 To <total of columns/shapes> Step 1 'Cells "For"
```
**Excel Workbook and PPT template download**
https://drive.google.com/drive/folders/1PkaVgj7f8mtjGhuCEzuATUJ8ADF6M8Sl?usp=sharing
