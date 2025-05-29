Sub Convert_Common_Table()

Application.ScreenUpdating = False

Sheets("Sheet Name").Select  'selects the sheet you want to convert

  ActiveSheet.Range("A2").Select  'selects the beginning cell of the range you want to convert
  SheetName.ListObjects.Add(xlSrcRange, Selection.CurrentRegion, , xlYes).Name = "Sheet Name"  'converts the range to a table

Application.ScreenUpdating = True

End Sub
