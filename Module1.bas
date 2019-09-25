Attribute VB_Name = "Module1"
Dim rollCount As Integer
Public player As Boolean
Dim lvl1 As Integer
Dim lvl2 As Integer

Sub kast()
    Call dbug

    rollCount = rollCount + 1
    Call rollDice
    Call switch
    
End Sub

Sub start()
    Call resetdbug
    
    player = True
    Call switch
    lvl1 = 0
    lvl2 = 0
    rollCount = 0
    Ark1.kast.Enabled = True
    Ark1.Range("C10:H27").Value = 0
End Sub

Sub rollDice()
    If Ark1.ToggleButton1.Value = False Then
        Ark1.Range("C2") = Module2.Random
    End If
    If Ark1.ToggleButton2.Value = False Then
        Ark1.Range("C3") = Module2.Random
    End If
    If Ark1.ToggleButton3.Value = False Then
        Ark1.Range("C4") = Module2.Random
    End If
    If Ark1.ToggleButton4.Value = False Then
        Ark1.Range("C5") = Module2.Random
    End If
    If Ark1.ToggleButton5.Value = False Then
        Ark1.Range("C6") = Module2.Random
    End If
End Sub
Sub switch()
     
    If rollCount = 3 Then
    
        If player Then
            lvl1 = lvl1 + 1
            Call Fill(lvl1)
        Else
            lvl2 = lvl2 + 1
            Call Fill2(lvl2)
        End If
        
        
        player = Not player
        rollCount = 0
    End If
    
    
    If player Then
        Ark1.Range("K11") = "Player1"
    Else
        Ark1.Range("K11") = "Player2"
    End If
End Sub
Sub dbug()
    Ark1.Range("O1").Value = rollCount
    Ark1.Range("O2").Value = lvl1
    Ark1.Range("O3").Value = lvl2
End Sub

Sub resetdbug()
    Ark1.Range("O1").Value = ""
    Ark1.Range("O2").Value = ""
    Ark1.Range("O3").Value = ""
End Sub



