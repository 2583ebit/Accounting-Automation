Option Explicit  

Sub Unprotect()  

Dim ws As Worksheet  
   
  For Each ws In Worksheets 'looks through all workbooks 
    ws.Unprotect "Password" 'uses your password to unprotect sheets 
  Next ws  
    
End Sub  
