Attribute VB_Name = "M�dulo1"
Sub LoopRowsSelected()
       
    Dim DataRange As Range
    Dim DataRow As Range
    Dim PPT
    Dim Pres
    Dim AddS
    Dim lDados(30) As String
    Dim cellFor As Integer
    Dim pptPath As String
    Dim MyPPT As Object
    
    Set MyPPT = CreateObject("Powerpoint.application")
    
'------- Open PowerPoint that is specific in the "Menu" sheets -------
    
    MyPPT.Visible = True
    pptPath = ThisWorkbook.Path & ActiveWorkbook.Sheets("Menu").Range("E7")
    
    'path to powerpoint
    MyPPT.presentations.Open pptPath
   
'------- Select data be used to create PPT -------
    
    ActiveWorkbook.Sheets("Data").Activate
    ActiveWorkbook.Sheets("Data").Range("B2").Select
    ActiveWorkbook.Sheets("Data").Range(Selection, Selection.End(xlUp)).Select
    ActiveWorkbook.Sheets("Data").Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Sheets("Data").Range(Selection, Selection.End(xlToRight)).Select
    
'------- Assign the selected data above to variable "DataRange" -------
   
    Set DataRange = Selection

    ActiveWorkbook.Sheets("Menu").Activate

'------- WHERE THE MAGIC HAPPENS: There are 2 "For" one that creates a slide
'------- for each line and the other that adds the data for each cell in the
'------- line to a shape in powerpoint

    Set Pres = MyPPT.ActivePresentation
        
    For Each DataRow In DataRange.Rows 'Line "For"
        
        Set AddS = Pres.Slides.AddSlide(Pres.Slides.Count + 1, Pres.SlideMaster.CustomLayouts(1))
        AddS.Shapes(17).TextFrame.TextRange = DataRow.Cells(2, 1)
        For i = 1 To 16 Step 1 'Cells "For"
           cellFor = i + 1
           lDados(i) = DataRow.Cells(2, cellFor)
           AddS.Shapes(i).TextFrame.TextRange = lDados(i)
        Next i
        
    Next DataRow
    
End Sub
