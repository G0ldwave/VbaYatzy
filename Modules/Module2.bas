Attribute VB_Name = "Module2"
Dim rng As String
Dim one As Integer
Dim two As Integer
Dim three As Integer
Dim four As Integer
Dim five As Integer
Dim six As Integer
Sub toggle1(b)
    If b Then
        Ark1.ToggleButton1.BackColor = RGB(0, 255, 0)
    Else
        Ark1.ToggleButton1.BackColor = RGB(255, 255, 255)
    End If
    
End Sub
Sub toggle2(b)
    If b Then
        Ark1.ToggleButton2.BackColor = RGB(0, 255, 0)
    Else
        Ark1.ToggleButton2.BackColor = RGB(255, 255, 255)
    End If
End Sub
Sub toggle3(b)
    If b Then
        Ark1.ToggleButton3.BackColor = RGB(0, 255, 0)
    Else
        Ark1.ToggleButton3.BackColor = RGB(255, 255, 255)
    End If
End Sub
Sub toggle4(b)
    If b Then
        Ark1.ToggleButton4.BackColor = RGB(0, 255, 0)
    Else
        Ark1.ToggleButton4.BackColor = RGB(255, 255, 255)
    End If
End Sub
Sub toggle5(b)
    If b Then
        Ark1.ToggleButton5.BackColor = RGB(0, 255, 0)
    Else
        Ark1.ToggleButton5.BackColor = RGB(255, 255, 255)
    End If
End Sub

Function Random()
    Random = Int((6 - 1 + 1) * Rnd + 1)
End Function
Sub finalsum()
    Ark1.Range("C27") = Application.sum(Range("C10:C26"))
End Sub
Sub finalsum2()
    Ark1.Range("D27") = Application.sum(Range("D10:D26"))
End Sub
Sub Fill(round)
    If round < 7 Then
        Call face(round, "C")
        Call sum
        Call bonus
    ElseIf round = 7 Then
        Call par("C18")
    ElseIf round = 8 Then
        Call topar("C19")
    ElseIf round = 9 Then
        Call trelike("C20")
    ElseIf round = 10 Then
        Call firelike("C21")
    ElseIf round = 11 Then
        Call litenstraight("C22")
    ElseIf round = 12 Then
        Call storstraight("C23")
    ElseIf round = 13 Then
        Call hus("C24")
    ElseIf round = 14 Then
        Call sjanse("C25")
    ElseIf round = 15 Then
        Call yatzy("C26")
    End If
    Call finalsum
    Call resettoggle
End Sub

Sub Fill2(round)
    If round < 7 Then
        Call face(round, "D")
        Call sum2
        Call bonus2
    ElseIf round = 7 Then
        Call par("D18")
    ElseIf round = 8 Then
        Call topar("D19")
    ElseIf round = 9 Then
        Call trelike("D20")
    ElseIf round = 10 Then
        Call firelike("D21")
    ElseIf round = 11 Then
        Call litenstraight("D22")
    ElseIf round = 12 Then
        Call storstraight("D23")
    ElseIf round = 13 Then
        Call hus("D24")
    ElseIf round = 14 Then
        Call sjanse("D25")
    ElseIf round = 15 Then
        Call yatzy("D26")
    End If
    Call finalsum2
    Call resettoggle
End Sub
Sub sjanse(rng)
    Ark1.Range(rng) = Application.sum(Range("C2:C6"))
End Sub
Sub hus(player)
    Call updatevals
    If six > 4 Then
        Ark1.Range(player) = 30
    ElseIf six > 2 And five > 1 Then
        Ark1.Range(player) = 28
    ElseIf five > 2 And six > 1 Then
        Ark1.Range(player) = 27
    ElseIf six > 2 And four > 1 Then
        Ark1.Range(player) = 26
    ElseIf five > 2 And five > 1 Then
        Ark1.Range(player) = 25
    ElseIf four > 2 And six > 1 Or six > 2 And three > 1 Then
        Ark1.Range(player) = 24
    ElseIf five > 2 And four > 1 Then
        Ark1.Range(player) = 23
    ElseIf six > 2 And two > 1 Or four > 2 And five > 1 Then
        Ark1.Range(player) = 22
    ElseIf five > 2 And three > 1 Or three > 2 And six > 1 Then
        Ark1.Range(player) = 21
    ElseIf four > 2 And four > 1 Or six > 2 And one > 1 Then
        Ark1.Range(player) = 20
    ElseIf three > 2 And five > 1 Or five > 2 And two > 1 Then
        Ark1.Range(player) = 19
    ElseIf four > 2 And three > 1 Or two > 2 And six > 1 Then
        Ark1.Range(player) = 18
    ElseIf five > 2 And one > 1 Or three > 2 And four > 1 Then
        Ark1.Range(player) = 17
    ElseIf two > 2 And five > 1 Or four > 2 And two > 1 Then
        Ark1.Range(player) = 16
    ElseIf three > 2 And three > 1 Or one > 2 And six > 1 Then
        Ark1.Range(player) = 15
    ElseIf four > 2 And one > 1 Or two > 2 And four > 1 Then
        Ark1.Range(player) = 14
    ElseIf one > 2 And five > 1 Or three > 2 And two > 1 Then
        Ark1.Range(player) = 13
    ElseIf two > 2 And three > 1 Then
        Ark1.Range(player) = 12
    ElseIf three > 2 And one > 1 Or one > 2 And four > 1 Then
        Ark1.Range(player) = 11
    ElseIf two > 2 And two > 1 Then
        Ark1.Range(player) = 10
    ElseIf one > 2 And three > 1 Then
        Ark1.Range(player) = 9
    ElseIf two > 2 And one > 1 Then
        Ark1.Range(player) = 8
    ElseIf one > 2 And two > 1 Then
        Ark1.Range(player) = 7
    ElseIf one > 2 And one > 1 Then
        Ark1.Range(player) = 5
    End If
End Sub
Sub storstraight(player)
    Call updatevals
    If six > 0 Then
        If five > 0 Then
            If four > 0 Then
                If three > 0 Then
                    If two > 0 Then
                        Ark1.Range(player) = 20
                    End If
                End If
            End If
        End If
    End If
End Sub
Sub litenstraight(player)
    Call updatevals
    If five > 0 Then
        If four > 0 Then
            If three > 0 Then
                If two > 0 Then
                    If one > 0 Then
                        Ark1.Range(player) = 15
                    End If
                End If
            End If
        End If
    End If
End Sub
Sub firelike(player)
    Call updatevals
    If six > 3 Then
        Ark1.Range(player) = 6 * 4
    ElseIf five > 3 Then
        Ark1.Range(player) = 5 * 4
    ElseIf four > 3 Then
        Ark1.Range(player) = 4 * 4
    ElseIf three > 3 Then
        Ark1.Range(player) = 3 * 4
    ElseIf two > 3 Then
        Ark1.Range(player) = 2 * 4
    ElseIf one > 3 Then
        Ark1.Range(player) = 1 * 4
    End If
End Sub
Sub trelike(player)
    Call updatevals
    If six > 2 Then
        Ark1.Range(player) = 6 * 3
    ElseIf five > 2 Then
        Ark1.Range(player) = 5 * 3
    ElseIf four > 2 Then
        Ark1.Range(player) = 4 * 3
    ElseIf three > 2 Then
        Ark1.Range(player) = 3 * 3
    ElseIf two > 2 Then
        Ark1.Range(player) = 2 * 3
    ElseIf one > 2 Then
        Ark1.Range(player) = 1 * 3
    End If
End Sub
Sub bonus()
    If Application.sum(Range("C10:C15")) > 62 Then
        Ark1.Range("C17") = 50
    End If
End Sub
Sub bonus2()
    If Application.sum(Range("D10:D15")) > 62 Then
        Ark1.Range("D17") = 50
    End If
End Sub
Sub sum()
    Ark1.Range("C16") = Application.sum(Range("C10:C15"))
End Sub
Sub sum2()
    Ark1.Range("D16") = Application.sum(Range("D10:D15"))
End Sub
Sub topar(player)
    Call updatevals
    
    If six > 3 Then
        Ark1.Range(player) = 24
    ElseIf six > 1 And five > 1 Then
        Ark1.Range(player) = 22
    ElseIf six > 1 And four > 1 Or five > 3 Then
        Ark1.Range(player) = 20
    ElseIf six > 1 And three > 1 Or four > 1 And five > 1 Then
        Ark1.Range(player) = 18
    ElseIf six > 1 And two > 1 Or five > 1 And three > 1 Or four > 3 Then
        Ark1.Range(player) = 16
    ElseIf six > 1 And one > 1 Or four > 1 And three > 1 Or five > 1 And two > 1 Then
        Ark1.Range(player) = 14
    ElseIf four > 1 And two > 1 Or five > 1 And one > 1 Or three > 3 Then
        Ark1.Range(player) = 12
    ElseIf three > 1 And two > 1 Or one > 1 And four > 1 Then
        Ark1.Range(player) = 10
    ElseIf two > 3 Or one > 1 And three > 1 Then
        Ark1.Range(player) = 8
    ElseIf one > 1 And two > 1 Then
        Ark1.Range(player) = 6
    ElseIf one > 3 Then
        Ark1.Range(player) = 6
    End If
End Sub
Sub par(player)
    Call updatevals
    
    If six > 1 Then
        Ark1.Range(player) = 6 * 2
    ElseIf five > 1 Then
        Ark1.Range(player) = 5 * 2
    ElseIf four > 1 Then
        Ark1.Range(player) = 4 * 2
    ElseIf three > 1 Then
        Ark1.Range(player) = 3 * 2
    ElseIf two > 1 Then
        Ark1.Range(player) = 2 * 2
    ElseIf one > 1 Then
        Ark1.Range(player) = 1 * 2
    End If
End Sub
Sub yatzy(player)
    Call updatevals
    
    If six > 5 Then
        Ark1.Range(player) = 50
    ElseIf five > 5 Then
        Ark1.Range(player) = 50
    ElseIf four > 5 Then
        Ark1.Range(player) = 50
    ElseIf three > 5 Then
        Ark1.Range(player) = 50
    ElseIf two > 5 Then
        Ark1.Range(player) = 50
    ElseIf one > 5 Then
        Ark1.Range(player) = 50
    End If
End Sub
Sub face(round, player)
    Call updatevals
    
    If round = 1 Then
        Ark1.Range(player & "10") = one * 1
    ElseIf round = 2 Then
        Ark1.Range(player & "11") = two * 2
    ElseIf round = 3 Then
        Ark1.Range(player & "12") = three * 3
    ElseIf round = 4 Then
        Ark1.Range(player & "13") = four * 4
    ElseIf round = 5 Then
        Ark1.Range(player & "14") = five * 5
    ElseIf round = 6 Then
        Ark1.Range(player & "15") = six * 6
    End If
End Sub

Sub updatevals()
    one = Application.WorksheetFunction.CountIf(Ark1.Range("C2:C6"), 1)
    two = Application.WorksheetFunction.CountIf(Ark1.Range("C2:C6"), 2)
    three = Application.WorksheetFunction.CountIf(Ark1.Range("C2:C6"), 3)
    four = Application.WorksheetFunction.CountIf(Ark1.Range("C2:C6"), 4)
    five = Application.WorksheetFunction.CountIf(Ark1.Range("C2:C6"), 5)
    six = Application.WorksheetFunction.CountIf(Ark1.Range("C2:C6"), 6)
End Sub

Sub resettoggle()
    Ark1.ToggleButton1.Value = False
    Ark1.ToggleButton2.Value = False
    Ark1.ToggleButton3.Value = False
    Ark1.ToggleButton4.Value = False
    Ark1.ToggleButton5.Value = False
    Call toggle1(Ark1.ToggleButton1.Value)
    Call toggle1(Ark1.ToggleButton1.Value)
    Call toggle1(Ark1.ToggleButton1.Value)
    Call toggle1(Ark1.ToggleButton1.Value)
    Call toggle1(Ark1.ToggleButton1.Value)
    
    
End Sub
