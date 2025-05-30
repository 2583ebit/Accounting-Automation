Option Explicit
Function Progress(Progress_Percentage As Single, Optional Unload_After_Completion As Boolean = True)  'updates progress bar on user form and unloads when completed

Dim Total_width, Current_W As Single
Dim FileCnt As Byte

Me.Show False  'allows user to continue interacting with Excel while code runs

VBA.DoEvents  'temporarily pauses Excel so other events can process

Total_width = 200  'width of progress bar
Current_W = (Total_width / 100) * Progress_Percentage  'caclulates current width based on % complete

Me.lbl_progress.Left = Current_W  'move progress bar horizontally
Me.lbl_value.Caption = Format(Progress_Percentage, "0") & "%"  'show % complete

If Unload_After_Completion = True And Progress_Percentage = 100 Then Unload Me  'close progress bar once finished

End Function


Private Sub lbl_progress_Click()

End Sub
