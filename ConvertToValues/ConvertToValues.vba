Option Explicit  
 
Sub PasteValues()  
   
Dim ws As Worksheet  
   
  For Each ws In Worksheets  
    ws.Cells.Copy  'copies all cells in worksheets
    ws.Cells.PasteSpecial xlValues  'pastes all cells as values
  Next ws  
     
End Sub  
