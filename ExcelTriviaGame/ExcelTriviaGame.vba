Option Explicit

Private currentRow As Integer
Private totalScore As Integer

Private Sub UserForm_Initialize()

    currentRow = 2 ' Starting row
    totalScore = 0 ' Reset score
    lblScore.Caption = "Score: " & totalScore  'shows total score at this label
    LoadQuestion 'loads first question

End Sub

Private Sub LoadQuestion()

    With Sheets("Questions")
        If .Cells(currentRow, 1).Value = "" Then 'when you hit an empty row in column A
            MsgBox "Game over! Final score: " & totalScore, vbInformation 'game ends and you get your final score
            Unload Me
            Exit Sub
                        
        End If

        lblQuestion.Caption = .Cells(currentRow, 1).Value 'this label loads the questions from the worksheet into the form
        opt1.Caption = .Cells(currentRow, 2).Value 
        opt2.Caption = .Cells(currentRow, 3).Value
        opt3.Caption = .Cells(currentRow, 4).Value
        
        'resets radio buttons
        opt1.Value = False
        opt2.Value = False
        opt3.Value = False
        
    End With
    
End Sub

Private Sub btnNext_Click()

    Dim selectedAnswer As String

    'stores user's answer to selected radio button
    If opt1.Value Then selectedAnswer = opt1.Caption
    If opt2.Value Then selectedAnswer = opt2.Caption
    If opt3.Value Then selectedAnswer = opt3.Caption

    'message if no answer selected
    If selectedAnswer = "" Then
        MsgBox "Please select an answer."
        
        Exit Sub
        
    End If

    If selectedAnswer = Sheets("Questions").Cells(currentRow, 5).Value Then 'checks if users' answer matches correct answer in worksheet
        totalScore = totalScore + 1 'if correct it adds to your score
        MsgBox "Correct!", vbInformation
        
    Else
        MsgBox "Wrong! The correct answer was: " & Sheets("Questions").Cells(currentRow, 5).Value, vbExclamation 'displays correct answer if you are wrong
        
    End If

    currentRow = currentRow + 1 'moves to next row
    lblScore.Caption = "Score: " & totalScore 'updates score
    LoadQuestion
    
End Sub

