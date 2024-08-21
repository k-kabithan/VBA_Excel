VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

    Dim folderExt As String ' Declare a variable to hold the folder path
    Dim folderName As String ' Declare a variable to hold the folder name
    Dim hyphenPos As Integer ' Declare a variable to hold the position of the hyphen
    Dim prefix As String ' Declare a variable to hold the prefix before the hyphen
    
    folderExt = ActiveWorkbook.Path ' Get the path of the active workbook
    
    If Right(folderExt, 1) = "\" Then
        folderExt = Left(folderExt, Len(folderExt) - 1) ' Remove the trailing backslash if it exists
    End If
    
    folderName = Split(folderExt, "\")(UBound(Split(folderExt, "\")) - 1) ' Get the folder name (second to last part of the path)
    
    hyphenPos = InStr(folderName, "-") ' Find the position of the hyphen in the folder name
    
    If hyphenPos > 0 Then
        prefix = Trim(Left(folderName, hyphenPos - 1)) ' Get the prefix before the hyphen
        
        If IsNumeric(prefix) Then
            folderName = Trim(Mid(folderName, hyphenPos + 2)) ' Remove the numeric prefix and hyphen
        End If
    End If
    
    Worksheets("Data Sheet 3 Legs").Range("C10").Value = folderName ' Set the value of cell C10 on the "Data Sheet 3 Legs" sheet to the modified folder name
    
End Sub

