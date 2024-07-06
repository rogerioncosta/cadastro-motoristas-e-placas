VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plan1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Nome_Change()

    Selection.AutoFilter Field:=1, Criteria1:=CStr("*" + Nome.Text) + "*"

End Sub


Private Sub CPF_Change()

    Selection.AutoFilter Field:=2, Criteria1:=CStr("*" + CPF.Text) + "*"

End Sub

Private Sub Placa_Change()

    Selection.AutoFilter Field:=3, Criteria1:=CStr("*" + Placa.Text) + "*"

End Sub
