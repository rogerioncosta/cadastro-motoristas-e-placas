VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPasta_de_trabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)
'senha plan bloqueadas "142"
Application.DisplayAlerts = False
Application.Quit
End Sub


Sub Limpar_textbox()
Dim tbx As OLEObject
For Each tbx In ActiveSheet.OLEObjects
If TypeName(tbx.Object) = "TextBox" Then tbx.Object.Text = ""

Next
If ActiveSheet.Name = "pesquisamotorista" Then
        Range("C6:E6").AutoFilter
        Range("C6:E6").AutoFilter
    ElseIf ActiveSheet.Name = "pesquisafrete" Then
        Range("C7:E7").AutoFilter
        Range("C7:E7").AutoFilter
    ElseIf ActiveSheet.Name = "duvidas" Then
        Range("C4").AutoFilter
        Range("C4").AutoFilter
    ElseIf ActiveSheet.Name = "vlplt" Then
        Range("B9:C9").AutoFilter
        Range("B9:C9").AutoFilter
   Else
End If

End Sub
Sub Inicio_botao()
Sheets("inicio").Select
End Sub

Sub Pesquisa_mot_botao()
Sheets("pesquisamotorista").Select
End Sub

Sub Placas_contatos_botao()
Sheets("plcctt").Select
End Sub

Sub Pesquisa_frete_botao()
Sheets("pesquisafrete").Select
End Sub

Sub Duvidas_botao()
Sheets("duvidas").Select
End Sub

Sub Valepallet_botao()
Sheets("vlplt").Select
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

Dim Senha As String
Senha = "******"

If FormCadastro.Visible = True Or FormEdicao.Visible = True Then
Exit Sub

Else

    If InputBox("Digite a senha para Salvar, ou apenas feche.", "Proteção") = Senha Then
        Exit Sub
    Else
        If SaveAsUI = True Then
            MsgBox "Não é permitido Salvar Como"
            Cancel = True
            Exit Sub
        End If
        
        If SaveAsUI = False Then
            MsgBox "Não é permitido 'Salvar'"
            Cancel = True
            Exit Sub
        End If
    End If
End If

End Sub
