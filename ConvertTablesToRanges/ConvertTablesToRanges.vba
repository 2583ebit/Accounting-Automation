Option Explicit

Sub ConvertTableToRangeCommon()

Dim xSheet As Worksheet
Dim xList As ListObject

Application.ScreenUpdating = False
    
    Set xSheet = Worksheets("Worksheet Name") 'selects the named worksheet

      For Each xList In xSheet.ListObjects  'for each table in the worksheet
          xList.Unlist  'convert to range
      Next

Application.ScreenUpdating = True
    
End Sub
