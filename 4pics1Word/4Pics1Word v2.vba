' Version 2.1

' Game Description:
' User is given an image, a keyboard containing a limited amount of letters, a label that shows their input, and a button that clears their 
' input. The user is given an unknown word that can only be described with the four photos they're given. If the user inputs the incorrect 
' answer say "WRONG!", else if they got the answer correctly say "CORRECT!". Upon getting the correct answer the user is given a NEXT button 
' that moves them to the next level/slide. Everything in the previous slide is resetted.

' Changelog:
'	- Added new btn_enabled(ByVal bool As Boolean) function
'	- Added new CHECK button
'		- Adds new check_Click() Sub
'		- Checks if user's answer is correct instead of checking for input length
'		- Replaces CLEAR button
'	- Added new reset() function
'		- Works similarly to Reset_Click()
        - Adds new ANSWER.Visible = True for changing answers
    - Removed Unused Function comments
'	- Removed CLEAR button
'	- Removed CLEAR_Click() Sub


Private Function userInput(ByVal btn As CommandButton) As Integer
        
    If Label.Caption = "CORRECT!" Or Label.Caption = "WRONG!" Then
        reset
    End If

    Button = btn.Caption ' Recieves String of Button Caption
    btn.Enabled = False ' Disables Button
    
    If Label.Caption = "Enter answer" Then
        Label.Caption = Button ' Checks if LABEL is empty (i.e. Enter Answer) and clears LABEL
    Else
        ' Simply adds letters to LABEL instead of clearing
        Label.Caption = Label.Caption & Button
    End If
    
End Function

Private Function reset()
    Label.Caption = "Enter answer"
    Label.BackColor = RGB(255, 255, 255)
    btnNext.Visible = False
    ANSWER.Visible = False
    ' ANSWER.Visible = True ' Uncomment this to change the answer. Comment to hide the answer.
    
    btn_enable
End Function

Private Function btn_enable()
    ' Resets buttons
    btn1.Enabled = True
    btn2.Enabled = True
    btn3.Enabled = True
    btn4.Enabled = True
    btn5.Enabled = True
    btn6.Enabled = True
    btn7.Enabled = True
    btn8.Enabled = True
    btn9.Enabled = True
    btn10.Enabled = True 
    
    ' Add Enabled Buttons here.
End Function

Sub OnSlideShowPageChange()
    Dim i As Integer
    i = ActivePresentation.SlideShowWindow.View.CurrentShowPosition
    ' If i <> 1 Then Exit Sub
    reset
End Sub

Private Sub btnNext_Click()
    ' Upon getting the correct answer the user is given a NEXT button that moves them to the next
    ' level/slide. Everything in the previous slide is resetted.

    ActivePresentation.SlideShowWindow.View.GotoSlide (ActivePresentation.SlideShowWindow.View.Slide.SlideIndex + 1)

End Sub

Private Sub check_Click()
    ' Matches user's input to answer
    If Label.Caption = ANSWER.Caption Then
        Label.Caption = "CORRECT!" ' Change LABEL text to CORRECT!
        Label.BackColor = RGB(0, 255, 0) ' Change LABEL background color to Green
        btnNext.Visible = True ' Make btnNext appear
    Else
        Label.Caption = "WRONG!" ' Change LABEL text to WRONG!
        Label.BackColor = RGB(255, 0, 0) ' Change Label background color to Red
        
        btn_enable
    End If
End Sub

Private Sub btn1_Click()
    userInput btn1
End Sub

Private Sub btn10_Click()
    userInput btn10
End Sub

Private Sub btn2_Click()
    userInput btn2
End Sub

Private Sub btn3_Click()
    userInput btn3
End Sub

Private Sub btn4_Click()
    userInput btn4
End Sub

Private Sub btn5_Click()
    userInput btn5
End Sub

Private Sub btn6_Click()
    userInput btn6
End Sub

Private Sub btn7_Click()
    userInput btn7
End Sub

Private Sub btn8_Click()
    userInput btn8
End Sub

Private Sub btn9_Click()
    userInput btn9
End Sub
