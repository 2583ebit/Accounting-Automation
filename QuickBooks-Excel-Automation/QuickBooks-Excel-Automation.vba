Option Explicit

Sub QBO_Import()

Dim FileToOpen As Variant
Dim FileCnt As Byte
Dim SelectedBook As Workbook
Dim ws As Worksheet
Dim pt As PivotTable

Application.ScreenUpdating = False

With Sheet3
    Dim lastRow As Long, lastCol As Long
    lastRow = .Cells(.Rows.Count, "B").End(xlUp).row
    lastCol = .Cells(3, .Columns.Count).End(xlToLeft).Column

    If lastRow >= 3 And lastCol >= 2 Then
        .Range("B3", .Cells(lastRow, lastCol)).ClearContents
    End If
End With

On Error GoTo Handle

'pick files to import using multiselect'
FileToOpen = Application.GetOpenFilename(filefilter:="Excel Files (*.xlsx),*.xlsx", Title:="Select Workbook to Import", MultiSelect:=True)

If IsArray(FileToOpen) Then
    For FileCnt = 1 To UBound(FileToOpen)
        Set SelectedBook = Workbooks.Open(Filename:=FileToOpen(FileCnt))
        
        With SelectedBook.Worksheets("Sheet1")
            Dim startCell As Range
            
            
            Set startCell = .Range("B3")
            lastRow = .Cells(.Rows.Count, startCell.Column).End(xlUp).row
            lastCol = .Cells(startCell.row, .Columns.Count).End(xlToLeft).Column
            
            .Range(startCell, .Cells(lastRow, lastCol)).Copy
        End With
        
        Sheet3.Range("B3").PasteSpecial xlPasteValues
        Application.CutCopyMode = False
        
        SelectedBook.Close SaveChanges:=False
    Next FileCnt
    
    ' Set the worksheet that contains the pivot table
    Set ws = ThisWorkbook.Sheets("Summary")  
    
    ' Set the pivot table object by name
    Set pt = ws.PivotTables("PivotTable1")
    
    ' Refresh the pivot table
    pt.RefreshTable
    
    Call GetClosestRate
    
    MsgBox "Data Imported Successfully", vbInformation, "!!......  :)  ......!!"
    
End If 'isarray

Sheet1.Select

Exit Sub

Handle:
If Err.Number = 1004 Then
    MsgBox "Workbook does not contain this security"
    Else
    MsgBox "An error has occured"
    End If

Application.ScreenUpdating = True

End Sub

Sub GetClosestRate()

    Dim http As Object
    Dim html As Object
    Dim table As Object, row As Object
    Dim dateVal As Date, rateVal As Double
    Dim targetDate As Date
    Dim closestDiff As Long: closestDiff = 100000
    Dim closestDate As Date
    Dim closestRate As Double
    Dim maxAllowedDiff As Long: maxAllowedDiff = 10 ' Max number of days difference allowed
    
    Application.ScreenUpdating = False
    
    ' Get target date from Summary sheet cell C6
    If IsDate(Sheets("Summary").Range("C6").Value) Then
        targetDate = Sheets("Summary").Range("C6").Value
    Else
        MsgBox "Invalid date in Summary!C6"
        Exit Sub
    End If
    
    ' Create HTTP request to fetch the page
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "https://www.federalreserve.gov/releases/h10/hist/dat00_hk.htm", False
    http.send

    ' Load the HTML document
    Set html = CreateObject("htmlfile")
    html.body.innerHTML = http.responseText
    
    ' Find the first table on the page
    Set table = html.getElementsByTagName("table")(0)
    
    ' Loop through each row starting after the header
    For Each row In table.getElementsByTagName("tr")
        If row.Children.Length = 2 Then
            On Error Resume Next
            dateVal = CDate(row.Children(0).innerText)
            rateVal = CDbl(row.Children(1).innerText)
            On Error GoTo 0
            
            If IsDate(dateVal) And IsNumeric(rateVal) Then
                Dim daysDiff As Long
                daysDiff = Abs(DateDiff("d", dateVal, targetDate))
                
                If daysDiff <= maxAllowedDiff Then
                    If daysDiff < closestDiff Then
                        closestDiff = daysDiff
                        closestDate = dateVal
                        closestRate = rateVal
                    End If
                End If
            End If
        End If
    Next row
    
    ' Output closest result to M9 and M10
    If closestDiff <= maxAllowedDiff Then
        With Sheets("Summary")
            '.Range("M9").Value = "Closest Date"
            .Range("l10").Value = closestDate
            '.Range("N9").Value = "Rate"
            .Range("m10").Value = closestRate
        End With
    Else
        MsgBox "No matching date found within " & maxAllowedDiff & " days.", vbExclamation
    End If

Application.ScreenUpdating = True

End Sub

Sub ExportSummaryToPDF()
    Dim ws As Worksheet
    Dim exportRange As Range
    Dim filePath As String
    Dim exportDate As String
    Dim timeStamp As String
    Dim desktopPath As String

    ' Set worksheet and range
    Set ws = ThisWorkbook.Sheets("Summary")
    Set exportRange = ws.Range("B4:D15")
    
    ' Get date from C6 and format
    If Not IsDate(ws.Range("C6").Value) Then
        MsgBox "Invalid date in cell C6", vbExclamation
        Exit Sub
    End If
    exportDate = Format(ws.Range("C6").Value, "yyyy-mm-dd")
    
    ' Get timestamp
    timeStamp = Format(Now, "hhmmss")
    
    ' Get desktop path
    desktopPath = Environ("USERPROFILE") & "\Desktop\"
    
    ' Set file path
    filePath = desktopPath & "Foreign Sub P&L - " & exportDate & " - " & timeStamp & ".pdf"

    ' Set page setup to center horizontally
    With ws.PageSetup
        .CenterHorizontally = True
        .CenterVertically = False
    End With
    
    ' Export to PDF
    On Error GoTo ErrHandler
    exportRange.ExportAsFixedFormat Type:=xlTypePDF, Filename:=filePath, Quality:=xlQualityStandard
    
    ' Open the PDF after export
    Shell "explorer.exe """ & filePath & """", vbNormalFocus
    
    MsgBox "Exported to: " & filePath, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error exporting PDF: " & Err.Description, vbCritical
End Sub
