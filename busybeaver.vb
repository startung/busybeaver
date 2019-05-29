
Private Sub Reset_Button_Click()
    ResetDeck
End Sub

Private Sub start_button_Click()
    RunBusyBeaver
End Sub

Sub RunBusyBeaver()
    Dim CurrCard As Long
    Dim CurrCell As Long
    Dim Shifts As Long
    Dim Score As Long
    Dim Cards() As Byte
    Dim i As Long
    Dim Delay As Long
    Dim Card1 As Long
    Dim Card2 As Long
    Dim Card3 As Long
    
    CurrCard = 1
    CurrCell = 1
    TotalCell = 1
    Shifts = 0
    Score = 0
    Delay = 1
    
    ' Setup cards
    Cards = StrConv(Left(PowerPoint.ActivePresentation.Name, Len(PowerPoint.ActivePresentation.Name) - 5), vbFromUnicode)
    For i = 0 To UBound(Cards)
        Cards(i) = Cards(i) - 48
    Next
    
    ' Reset deck
    ResetDeck
    
    ' Go To Start
    ActivePresentation.SlideShowWindow.View.Next
    
        Do While CurrCard <> 0
            Shifts = Shifts + 1
            CurrVal = ActivePresentation.Slides(CurrCell + 1).Shapes.Item(1).TextFrame.TextRange.Text
            Card1 = Cards(((CurrCard - 1) * 6) + (3 * CurrVal) + 1)
            Card2 = Cards(((CurrCard - 1) * 6) + (3 * CurrVal) + 2)
            Card3 = Cards(((CurrCard - 1) * 6) + (3 * CurrVal) + 3)
                   
            If Cards(0) = 20 Then ' First character in filename D then turn on "Debugging"
                MsgBox "Card:" & CurrCard & "-" & Card1 & Card2 & Card3
            ElseIf Cards(0) = 33 Then ' First character in filename Q then turn on "Quick"
                Delay = 0
            End If
            
            DelayMe (Delay) ' Add pause
            
            'MsgBox "Cur:" & CurrVal & " New:" & Card1
            
            If Card1 = 0 Then
                If CurrVal = "1" Then
                    Score = Score - 1
                    ActivePresentation.Slides(CurrCell + 1).Shapes.Item(1).TextFrame.TextRange.Text = "0"
                End If
            ElseIf Card1 = 1 Then
                If CurrVal = "0" Then
                    Score = Score + 1
                    ActivePresentation.Slides(CurrCell + 1).Shapes.Item(1).TextFrame.TextRange.Text = "1"
                End If
            Else
                MsgBox "ERROR: filename use only binary values (except for the first character)"
                Exit Sub
            End If
            
            DelayMe (Delay) ' Add pause
            
            If Card2 = 0 Then
                If CurrCell = 1 Then
                    ActivePresentation.Slides(TotalCell + 2).Copy
                    ActivePresentation.Slides.Paste (2)
                    TotalCell = TotalCell + 1
                    CurrCell = CurrCell + 1
                End If
                
                DelayMe (Delay) ' Add pause
                
                CurrCell = CurrCell - 1
                ActivePresentation.SlideShowWindow.View.Previous
                
            ElseIf Card2 = 1 Then
                If CurrCell = TotalCell Then
                    ActivePresentation.Slides(TotalCell + 2).Copy
                    ActivePresentation.Slides.Paste (TotalCell + 2)
                    TotalCell = TotalCell + 1
                End If
                
                DelayMe (Delay) ' Add pause
                
                CurrCell = CurrCell + 1
                ActivePresentation.SlideShowWindow.View.Next
            Else
                MsgBox "ERROR: filename use only binary values (except for the first character)"
                Exit Sub
            End If

            CurrCard = Card3
            DelayMe (Delay) ' Add pause
        Loop
    
    Result Shifts, Score
End Sub

Sub ResetDeck()
    Dim i As Long
    
    Do While ActivePresentation.Slides.Count > 3
        If ActivePresentation.Slides(4).Shapes.Count = 2 Then ' Start slide
            ActivePresentation.Slides(4).MoveTo 2
        Else
            ActivePresentation.Slides(4).Delete
        End If
    Loop
    If ActivePresentation.Slides(3).Shapes.Count = 2 Then ' Start slide
        ActivePresentation.Slides(3).MoveTo 2
    End If
    ActivePresentation.Slides(2).Shapes.Item(1).TextFrame.TextRange.Text = "0"
    ActivePresentation.Slides(3).Shapes.Item(1).TextFrame.TextRange.Text = "0"
   
    SlideShowWindows(1).View.GotoSlide 1
End Sub

Sub Result(Shifts As Long, Score As Long)
    Dim MyMsg As String

    MyMsg = "You have reached card 0. The number of shifts was " & Shifts & ", and you scored " & Score & "."

    MsgBox (MyMsg)
End Sub

Sub DelayMe(Delay As Long)
    Start = Timer
    While Timer < Start + Delay
        DoEvents
    Wend
End Sub