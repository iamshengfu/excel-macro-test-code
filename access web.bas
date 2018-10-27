Sub ºê1()

'
    Application.CutCopyMode = False
    With ActiveWorkbook.XmlMaps("html_Ó³Éä")
        .ShowImportExportValidationErrors = False
        .AdjustColumnWidth = True
        .PreserveColumnFilter = True
        .PreserveNumberFormatting = True
        .AppendOnImport = False
    End With
    Application.CutCopyMode = False
    ActiveWorkbook.XmlImport Url:= _
        "https://mail.126.com/js6/main.jsp?sid=mASckbVBonQZNvNfzDBBSzWdNYzdzPte&df=mail126_letter%23module=welcome.WelcomeModule%7C%7B%7D" _
        , ImportMap:=Nothing, Overwrite:=True, Destination:=Range("$A$1")
End Sub
Sub ºê2()
'
' ºê2 ºê
'

'
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;https://mail.google.com/mail/u/0/#inbox", Destination:=Range("$A$21"))
        .CommandType = 0
        .Name = "#inbox"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = True
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = False
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
End Sub
