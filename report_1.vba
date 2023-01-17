Sub auto_open()

    Application.ScreenUpdating = False
    Application.AlertBeforeOverwriting = False
    Application.DisplayAlerts = False
    
    Range("A4:W4", Range("B4:W4").End(xlDown).End(xlToRight)).Clear

    Workbooks.Open "C:\xxxxxxx\master_source.csv"
    Windows("master_source.csv").Activate
    Range("AH2:BD2", Range("AI2:BD2").End(xlDown).End(xlToRight)).Select
    Range("AH2").Activate
    Selection.Copy
    Windows("table_3.xlsm").Activate
    Sheet1.Select
    Range("A4").Select
    Sheet1.Paste
    Windows("master_source.csv").Application.CutCopyMode = False
    Windows("master_source.csv").Close
    
    Range("A2:W2", Range("B2:W2").End(xlDown).End(xlToRight)).Borders.LineStyle = XlLineStyle.xlContinuous
    Range("A2:W2", Range("B2:W2").End(xlDown).End(xlToRight)).HorizontalAlignment = xlCenter
    Range("A2:W2", Range("B2:W2").End(xlDown).End(xlToRight)).VerticalAlignment = xlCenter
    
    Dim MyRange As Range
    Dim MyRange2 As Range
    Dim MyRange3 As Range
    
    Set MyRange = Range("G4:G200")
    Set MyRange2 = Range("H4:H200")
    Set MyRange3 = Range("E4:E200")
    
    Dim strFormula As String
    Dim strFormula2 As String
    Dim strFormula3 As String
    
    strFormula = "=OR(AND(A4=""ROM"", G4>0.042361111), AND(A4=""CZE"", G4>0.022916667), AND(A4=""SVK"", G4>0.022916667), AND(A4=""POL"", G4>0.022916667), AND(A4=""BGR"", G4>0.042361111), AND(A4=""IND"", G4>0.031944444))"
    strFormula2 = "=OR(AND(A4=""ROM"", H4>0.021527778), AND(A4=""CZE"", H4>0.022916667), AND(A4=""SVK"", H4>0.022916667), AND(A4=""POL"", H4>0.0125), AND(A4=""BGR"", H4>0.021527778), AND(A4=""IND"", H4>0.021527778))"
    
    
    MyRange.FormatConditions.Delete
    MyRange2.FormatConditions.Delete
    MyRange3.FormatConditions.Delete
    
    MyRange.FormatConditions.Add Type:=xlExpression, Operator:=xlGreater, _
    Formula1:=strFormula
    
    MyRange2.FormatConditions.Add Type:=xlExpression, Operator:=xlGreater, _
    Formula1:=strFormula2
    
  
    Dim iRange As Range
    Dim condition1 As FormatCondition
    Dim condition2 As FormatCondition
    Set iRange = Range("E4:E30")
    iRange.FormatConditions.Delete
    
    Set condition1 = iRange.FormatConditions.Add(xlExpression, xlGreater, "=D4=0")
    Set condition2 = iRange.FormatConditions.Add(xlCellValue, xlGreater, "=D4")
    
    With condition1
    .Interior.Color = vbWhite
    .Font.Color = vbBlack
    .StopIfTrue = True
    End With
    
      
    With condition2
     .Interior.Color = vbRed
     .Font.Color = vbBlack
    End With
    
    MyRange.FormatConditions(1).Interior.Color = RGB(255, 51, 51)
    MyRange2.FormatConditions(1).Interior.Color = RGB(255, 51, 51)
   

    
    ActiveWorkbook.Save
    
   
    Application.Quit
  
End Sub
