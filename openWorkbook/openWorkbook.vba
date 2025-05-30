Option Explicit

Sub Import_CapTablesCommon()

Dim FileToOpen As Variant
Dim FileCnt As Byte
Dim SelectedBook As Workbook

On Error GoTo Handle  'on error will display error message

FileToOpen = Application.GetOpenFilename(filefilter:="Excel Files (*.xlsx),*.xlsx", Title:="Select Workbook to Import", MultiSelect:=True)  'opens the selected excel workbook from files.  Can make multiple selections.

If IsArray(FileToOpen) Then  'checks if mutliple files are being opened
    For FileCnt = 1 To UBound(FileToOpen)  'if yes, loops through each file
    Set SelectedBook = Workbooks.Open(Filename:=FileToOpen(FileCnt))  'opens the workbook(s)
    Next FileCnt  'moves to next file
    
    MsgBox "File Opened Successfully", vbInformation, "!!......  :)  ......!!"  'message that file opened successfully
    
End If  

Exit Sub

Handle:  'error handler
If Err.Number = 1004 Then  'checks if error 1004 triggered
    MsgBox "Workbook does not contain this security"  '1004 error message
    Else  'if another error
    MsgBox "An error has occured"  'general error message
    End If

End Sub
