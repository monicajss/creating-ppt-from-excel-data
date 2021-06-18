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
    Dim pptName As String
    Dim newPPT As Object
    Dim fileName As String
    Dim pptFullName As String
    
    Set MyPPT = CreateObject("Powerpoint.application")
    
'------- Open PowerPoint that is specific in the "Menu" sheets -------
    
    MyPPT.Visible = True
    'path to powerpoint
    pptFullName = ThisWorkbook.path & ActiveWorkbook.Sheets("Menu").Range("E7")
    
    pptPath = pptFullName & ".pptx"
              
    MyPPT.presentations.Open pptPath

'------- Define a name for the new ppy
    fileName = pptFullName & Format(Now(), "yyyymmdd")
   
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
        
        If DataRow.Cells(2, 1) <> "" Then
                     
               Set AddS = Pres.Slides.AddSlide(Pres.Slides.Count + 1, Pres.SlideMaster.CustomLayouts(1))
                            
               AddS.Shapes(17).TextFrame.TextRange = DataRow.Cells(2, 1)
               
               For i = 1 To 16 Step 1 'Cells "For"
                  cellFor = i + 1
                  lDados(i) = DataRow.Cells(2, cellFor)
                  AddS.Shapes(i).TextFrame.TextRange = lDados(i)
               Next i

    Next DataRow
       
    '------- Save the new ppt and quit all of them

    Set newPPT = MyPPT.ActivePresentation
    
    newPPT.SaveCopyAs fileName
    MyPPT.Quit
    MsgBox ("Created powerpoint!")
    
End Sub
