' Version 1.1

' Game Description:
' User is given an image, a keyboard containing a limited amount of letters, a label that shows their input, and a button that clears their 
' input. The user is given an unknown word that can only be described with the four photos they're given. If the user inputs the incorrect 
' answer say "WRONG!", else if they got the answer correctly say "CORRECT!". Upon getting the correct answer the user is given a NEXT button 
' that moves them to the next level/slide. Everything in the previous slide is resetted.

' Changelog for 1.1:
' Updated for Version 2.0 to be interchangeable between the two versions
'	- Updated Reset_Click to check_Click

Private Function userInput(ByVal btn As CommandButton) As Integer
    Button = btn.Caption ' Recieves String of Button Caption
    btn.Enabled = False ' Disables Button
    
    If Label.Caption = "Enter answer" Then
        Label.Caption = Button ' Checks if LABEL is empty (i.e. Enter Answer) and clears LABEL
    Else
        ' Simply adds letters to LABEL instead of clearing
        temp = Label.Caption
        Label.Caption = temp & Button
    End If
    

    If Len(Label.Caption) >= Len(ANSWER.Caption) Then ' Checks word length of user's inputs

        ' Matches user's input to answer
        If Label.Caption = ANSWER.Caption Then
            Label.Caption = "CORRECT!" ' Change LABEL text to CORRECT!
            Label.BackColor = RGB(0, 255, 0) ' Change LABEL background color to Green
            btnNext.Visible = True ' Make btnNext appear
        Else
            Label.Caption = "WRONG!" ' Change LABEL text to WRONG!
            Label.BackColor = RGB(255, 0, 0) ' Change Label background color to Red
        End If
    End If
    
End Function

Sub OnSlideShowPageChange()
    Dim i As Integer
    i = ActivePresentation.SlideShowWindow.View.CurrentShowPosition
    ' If i <> 1 Then Exit Sub
    check_Click
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

Private Sub btnNext_Click()
    
    ActivePresentation.SlideShowWindow.View.GotoSlide (ActivePresentation.SlideShowWindow.View.Slide.SlideIndex + 1)

End Sub

Private Sub check_Click()
    Label.Caption = "Enter answer"
    Label.BackColor = RGB(255, 255, 255)
    btnNext.Visible = False
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
End Sub