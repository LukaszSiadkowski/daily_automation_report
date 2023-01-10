Sub auto_open()

    Application.ScreenUpdating = False
    Application.AlertBeforeOverwriting = False
    Application.DisplayAlerts = False
    
    
    Range("A4:R200").Clear
    
    
    Workbooks.Open "C:\xxxxxxxxx\master_source.csv"
    Windows("master_source.csv").Activate
    Range("A2:R2", Range("B2:R2").End(xlDown).End(xlToRight)).Select
    Range("A2").Activate
    Selection.Copy
    Windows("table_1.xlsm").Activate
    Sheet1.Select
    Range("A4").Select
    Sheet1.Paste
    Windows("master_source.csv").Application.CutCopyMode = False
    Windows("master_source.csv").Close
    
    Range("A2:R2", Range("B2:R2").End(xlDown).End(xlToRight)).Borders.LineStyle = XlLineStyle.xlContinuous
    Range("A2:R2", Range("B2:R2").End(xlDown).End(xlToRight)).HorizontalAlignment = xlCenter
    Range("A2:R2", Range("B2:R2").End(xlDown).End(xlToRight)).VerticalAlignment = xlCenter

    

    
    ActiveWorkbook.Save
    
    

    
    Application.Quit
  
End Sub
