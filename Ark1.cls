VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ark1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub kast_Click()
    Call Module1.kast
End Sub

Private Sub start_Click()
    Call Module1.start
End Sub

Private Sub ToggleButton1_Click()
    Call Module2.toggle1(ToggleButton1.Value)
End Sub

Private Sub ToggleButton2_Click()
    Call Module2.toggle2(ToggleButton2.Value)
End Sub

Private Sub ToggleButton3_Click()
    Call Module2.toggle3(ToggleButton3.Value)
End Sub

Private Sub ToggleButton4_Click()
    Call Module2.toggle4(ToggleButton4.Value)
End Sub

Private Sub ToggleButton5_Click()
    Call Module2.toggle5(ToggleButton5.Value)
End Sub
