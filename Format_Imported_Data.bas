Attribute VB_Name = "Module2"
Private Sub FormatData()
'
' FormatData Macro
'

'
    ' Clear any existing sort fields
    With ActiveWorkbook.Worksheets("Sheet 1").Sort
        .SortFields.Clear
        ' Add sorting by column F in descending order
        .SortFields.Add2 Key:=Range("F14:F183"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        ' Add sorting by column E in ascending order
        .SortFields.Add2 Key:=Range("E14:E183"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ' Apply sorting to the range A13:M183
        .SetRange Range("A13:M183")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' Format the range A14:M183
    With Range("A14:M183")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Font.Bold = False
    End With

    ' Auto-fit specific columns
    Worksheets("Sheet 1").Range("I:I, K:K, M:M").EntireColumn.AutoFit
    
    ' Delete any empty rows below the last non-empty row in column A
    Range("A14").End(xlDown).Offset(1, 0).Select
    Do Until Selection.Value = ""
        Selection.EntireRow.Delete
    Loop
End Sub

