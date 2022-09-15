Sub auto_open()

    Application.ScreenUpdating = False
    Application.AlertBeforeOverwriting = False
    Application.DisplayAlerts = False

    Workbooks.Open "xxxxxxxxxx\Daily\main_team.csv"
    Windows("main_market.csv").Activate
    Range("A2:U300").Select
    Range("A2").Activate
    Selection.Copy
    Windows("market_1.xlsm").Activate
    Sheet1.Select
    Range("A5").Select
    Sheet1.Paste
    Windows("main_team.csv").Application.CutCopyMode = False
    Windows("main_team.csv").Close
    
    
    Workbooks.Open "xxxxxxxxx\Daily\external_team.csv"
    Windows("external_team.csv").Activate
    Range("A2:U300").Select
    Range("A2").Activate
    Selection.Copy
    Windows("market_1.xlsm").Activate
    Sheet1.Select
    Range("Y5").Select
    Sheet1.Paste
    Windows("external_team.csv").Application.CutCopyMode = False
    Windows("external_team.csv").Close
    
    Range("A1:AU300").Borders.LineStyle = XlLineStyle.xlContinuous
    Range("A1:AU300").HorizontalAlignment = xlCenter
    Range("A1:AU300").VerticalAlignment = xlCenter


   
    Dim strbody1, strbody2, strbody3, strbody4, strbody5, strbody6, strbody7, hour, day As String
    Dim ahora As Long
    Dim Body, Gname, Gpath As Variant
    Dim chartG As Chart
    Dim grf As ChartObject
    Dim rng1 As Range
    Dim rng2 As Range
    Dim rng3 As Range
    Dim rng4 As Range
    Dim OutApp, OutMail As Object
    
    Set rng1 = Nothing
    On Error Resume Next
    
    Set grf = Nothing
    On Error Resume Next
    
    day = Application.WorksheetFunction.Floor_Math(now() - 1)
    day = Format(day, "dd/mm")
    
    hour = Application.WorksheetFunction.Floor_Math(now() - 1 / 24, 0.5 / 24)
    hour = Format(hour, "hh:mm")
    
    
    'ranges:
    
    
    Set rng1 = Range("A2:P2", Range("B2:P2").End(xlDown).End(xlToRight))
    Set rng2 = Range("Q2:W2", Range("Q2:W2").End(xlDown).End(xlToRight))
    Set rng3 = Range("Y2:AN2", Range("Z2:AN2").End(xlDown).End(xlToRight))
    Set rng4 = Range("AO2:AU2", Range("AO2:AU2").End(xlDown).End(xlToRight))
    
     
    'email body pieces:
    
    strbody1 = "Hello Team,<br><br>" & _
                "Hope this email finds you well.<br><br>" & _
            "Please find below the update of xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx for previous day for xxxxxxxxxxxxxxxxxxxxxx </b><br>" & _
            "Disclaimer in the bottom" & _
            "<h3>xxxxxxxxxxxxxx</h3>" & _
            "<h4>xxxxxxxxxxxxxxx</h4><br>"
            
    strbody2 = "<h4>xxxxxxxxxxxx metrics:</h4><br>"
    
    strbody3 = "<h3>xxxxxxxxxxx</h3>" & _
                "<h4>xxxxxxxxxxxxx</h4><br>"
    
    strbody4 = "<h4>xxxxxxxxxxxxxxxx</h4><br>"
                           
                
    strbody5 = "<br><U>If you will have any questions please don't hesitate to contact me.</U><br>" & _
               "<br><br>Kind regards,<br>Lukasz Siadkowski"
    
    strbody6 = "Hello Team,<br><br>" & _
                "Hope this email finds you well.<br><br>" & _
            "Please find below the update of aaaaaaaaaaaaaaaaaaaaaaaaaa for previous day for aaaaaaaaaaaaaaaaaaaaaaaaaaa </b><br>" & _
            "Disclaimer in the bottom<br><br>" & _
            "<h3>aaaaaaaaaaaaaaaaa</h3>" & _
            "<h4>aaaaaaaaaaaaaaaaaaaa</h4>"
            
    strbody7 = "<br><br><b>Disclaimer:</b><br><br>" & _
                xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx <br><br>" & _
            "aaaaaaaaaaaaaaaaaaaaaa</b><br>" & _
            "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaa</b><br>" & _
            "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa</b><br>" & _
            "aaaaaaaaaaaaaaaaaaaaaaaa</b><br>" & _
            "aaaaaaaaaaaaaaaaaaaaaaa</b><br>"

            
      
    
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    On Error Resume Next
            
    W = Application.WorksheetFunction.WeekNum(now())
    Y = Year(now())
        
    'email body:
    Body = strbody1 & RangetoHTML(rng1) & "<br>" & strbody2 & RangetoHTML(rng2) & "<br>" & strbody3 & RangetoHTML(rng3) & "<br>" & strbody4 & RangetoHTML(rng4) & "<br>" & strbody5 & strbody7
    
    
    With OutMail
        .To = "xxxxxxxxxxxxxxx"
        .CC = "aaaaaaaaaaaaaaa"
        .BCC = ""
        .Subject = "xxxxxxxxxxxxxxxxx " & day
        .HTMLBody = Body
        '.SentOnBehalfOfName = "aaaaaaaaaaaaaaaaaaa"
        .SentOnBehalfOfName = "xxxxxxxxxxxxxxxxxx"
        '.Display
        .Send
    End With
    
    On Error GoTo 0
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
    
    
    
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    On Error Resume Next
            
    W = Application.WorksheetFunction.WeekNum(now())
    Y = Year(now())
    
    Body_1 = strbody6 & RangetoHTML(rng3) & "<br>" & strbody4 & RangetoHTML(rng4) & "<br>" & strbody5 & strbody7
    
    With OutMail
        .To = "xxxxxxxxxxxxxxx"
        .CC = "aaaaaaaaaaaaaaa"
        .BCC = ""
        .Subject = "xxxxxxxxxxxxxxxxx " & day
        .HTMLBody = Body
        '.SentOnBehalfOfName = "aaaaaaaaaaaaaaaaaaa"
        .SentOnBehalfOfName = "xxxxxxxxxxxxxxxxxx"
        '.Display
        .Send
    End With
    
    On Error GoTo 0
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
    
    ActiveWorkbook.SaveAs "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx " & Format(now() - 1, "DD-MMM-YYYY"), 51
    
    
    Application.Quit
    


End Sub


Function RangetoHTML(rng As Range)

    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function

