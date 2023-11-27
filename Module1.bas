Attribute VB_Name = "Module1"
Sub Monday_CopyDataBasedOnDate()
    ' Define the source and destination workbooks
    Dim sourceWorkbook As Workbook
    Dim destinationWorkbook As Workbook
    
    ' Change the file paths as needed
    Dim sourcePath As String
    sourcePath = Workbooks("Automation_main").Sheets("Sheet1").Range("B1").Value & "\" & Workbooks("Automation_main").Sheets("Sheet1").Range("B2").Value & "." & Workbooks("Automation_main").Sheets("Sheet1").Range("B3").Value
    
    Dim destinationPath As String
    destinationPath = Workbooks("Automation_main").Sheets("Sheet1").Range("B29").Value & "\" & Workbooks("Automation_main").Sheets("Sheet1").Range("B30").Value & "." & Workbooks("Automation_main").Sheets("Sheet1").Range("B31").Value
    
    ' Open the source workbook in the background without updating links
    Set sourceWorkbook = Workbooks.Open(sourcePath, UpdateLinks:=False)
    
    ' Open the destination workbook in the background without updating links
    Set destinationWorkbook = Workbooks.Open(destinationPath, UpdateLinks:=False)
    
    ' Get the sheet name from B4 of Workbooks("Automation_main")
    Dim sourceSheetName As String
    sourceSheetName = Workbooks("Automation_main").Sheets("Sheet1").Range("B4").Value
    
    ' Set the source sheet using the retrieved sheet name
    Dim sourceSheet As Worksheet
    Set sourceSheet = sourceWorkbook.Sheets(sourceSheetName)
    
    ' If the source sheet is not found, handle the error
    If sourceSheet Is Nothing Then
        MsgBox "Source sheet not found!"
        Exit Sub
    End If
    
    ' Get the sheet name from B32 of Workbooks("Automation_main")
    Dim destinationSheetName As String
    destinationSheetName = Workbooks("Automation_main").Sheets("Sheet1").Range("B32").Value
    
    ' Set the destination sheet using the retrieved sheet name
    Dim destinationSheet As Worksheet
    Set destinationSheet = destinationWorkbook.Sheets(destinationSheetName)
    
    ' If the destination sheet is not found, handle the error
    If destinationSheet Is Nothing Then
        MsgBox "Destination sheet not found!"
        Exit Sub
    End If
    
    ' Define the date column in source and destination sheets
    Dim dateColumnSource As Range
    Set dateColumnSource = sourceSheet.Columns(Workbooks("Automation_main").Sheets("Sheet1").Range("B5").Value)
    Dim singledateColumnSource As Variant
    singledateColumnSource = Workbooks("Automation_main").Sheets("Sheet1").Range("B5").Value
    
    Dim dateColumnDestination As Range
    Set dateColumnDestination = destinationSheet.Columns(Workbooks("Automation_main").Sheets("Sheet1").Range("B33").Value)
    
    ' Find the last row with data in the date column of the source sheet
    Dim lastRowSource As Long
    lastRowSource = sourceSheet.Cells(sourceSheet.Rows.Count, dateColumnSource.Column).End(xlUp).Row

    ' Loop through each row in the source sheet
    On Error Resume Next
    For i = 1 To lastRowSource
    
    ' Check if the date in the source sheet matches any date in the destination sheet
    Dim matchIndex As Variant
    matchIndex = Application.Match(CLng(sourceSheet.Cells(i, singledateColumnSource).Value), dateColumnDestination, 0)
        
        ' Copy data only for matching dates
        Dim sourceProfitColumn As String
        sourceProfitColumn = Workbooks("Automation_main").Sheets("Sheet1").Range("B6").Value
        Dim destinationProfitColumn As String
        destinationProfitColumn = Workbooks("Automation_main").Sheets("Sheet1").Range("B34").Value

        ' If the destination row is found and it's not an error, copy the profit value
        If Not IsError(matchIndex) Then
            destinationSheet.Cells(matchIndex, destinationProfitColumn).Value = sourceSheet.Cells(i, sourceProfitColumn).Value

        End If
    
Next i
On Error GoTo 0

End Sub
