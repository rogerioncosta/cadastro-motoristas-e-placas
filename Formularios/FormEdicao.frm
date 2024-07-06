VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormEdicao 
   Caption         =   "Edição"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4830
   OleObjectBlob   =   "FormEdicao.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormEdicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Me.TextName.Enabled = False
    Me.TextName.BackColor = &H80000004
    Me.Label1.ForeColor = &H8000000A
End Sub

Private Sub TextCpf_Change() 'AfterUpdate

    Sheets("pesquisamotorista").Activate
    
    'Rows(1).Find(What:="Cargas", LookAt:=xlWhole).Select
    'Columns("D").Find(What:=Me.TextCpf.Value).Select
        
    If Not IsNumeric(Me.TextCpf.Value) Then
        MsgBox "São permitidos somente números neste campo!*", vbInformation
        'remove ponto, vírgulas, traços e barras do cpf ou renavam do form
        Me.TextCpf.Value = Replace(Me.TextCpf.Value, ".", "")
        Me.TextCpf.Value = Replace(Me.TextCpf.Value, ",", "")
        Me.TextCpf.Value = Replace(Me.TextCpf.Value, "-", "")
        Me.TextCpf.Value = Replace(Me.TextCpf.Value, "/", "")
        Exit Sub
    End If

    'remove espaços em branco a direita e esquerda do form
    Me.TextCpf.Value = WorksheetFunction.Trim(Me.TextCpf)
    
    Range("C6:E6").AutoFilter
    Range("C6:E6").AutoFilter
    
    'Limpa qualquer texto nas caixas de pesquisa da planilha nome e cpf
    Dim tbx As OLEObject
    For Each tbx In ActiveSheet.OLEObjects
        If TypeName(tbx.Object) = "TextBox" Then tbx.Object.Text = ""
    Next
    
    Dim rng As Range
    Dim valorProcurado As String
    
    valorProcurado = Me.TextCpf.Value
    
    If Len(valorProcurado) > 11 Then
        MsgBox "Valor não encontrado"
        Me.TextName.Value = ""
        UserForm_Initialize
        Range("A1").Select
        Exit Sub
    End If
    
'    If Len(valorProcurado) > 9 And ActiveCell.Offset(0, 1).Value = "PLACA" Then
'        MsgBox "Valor não encontrado"
'        Me.TextName.Value = ""
'        UserForm_Initialize
'        Range("A1").Select
'        Exit Sub
'    End If
    
    If Len(valorProcurado) = 9 Or Len(valorProcurado) = 11 Then ' Verifica se o número está completamente digitado
                
        Set rng = Columns("D").Find(What:=valorProcurado)
                
        If Not rng Is Nothing Then
        
            rng.Select
            Me.TextName.Enabled = True
            Me.TextName.BackColor = &H80000005
            Me.Label1.ForeColor = &H80000008
            Me.TextName = ActiveCell.Offset(0, -1).Value
            
        Else
            MsgBox "Valor não encontrado.", vbInformation
            'Me.TextName.Value = ""
            'UserForm_Initialize
            'Range("A1").Select
    
        End If
        
    End If

    
End Sub

Private Sub cancelar_Click()
    Range("A1").Select
    Unload Me
End Sub

Private Sub Editar_Click()

    'Verificar se tem campos em branco
'    If Me.TextName = "" Or Me.TextCpf = "" Then
'        MsgBox "Não são permitidos campos em branco!", vbInformation
'        Exit Sub
'    End If
    
    'Verificar se tem campos em branco
    If Me.TextName.Enabled = True And Me.TextCpf.Enabled = False Then
        If Me.TextName = "" Or Me.TextCpf = "" Then
            MsgBox "Não são permitidos campos em branco!", vbInformation
            Exit Sub
        End If
        
        'remove espaços em branco a direita e esquerda do form
        Me.TextName.Value = WorksheetFunction.Trim(Me.TextName)
        
        ActiveCell.Offset(0, -1).Select
        
        '==Realizar a edição dos dados
        ActiveCell.Value = Me.TextName.Text
        'transforma campo em maiusculo
        maiusc = UCase(ActiveCell.Value)
        ActiveCell.Value = maiusc
        'remove espaços em branco a direita e esquerda
        ActiveCell.Value = WorksheetFunction.Trim(ActiveCell)
        
    GoTo continue
    End If
    
    If Me.TextCpf.Enabled = True Then
    
        'Forçar somente números
        If Not IsNumeric(Me.TextCpf.Value) Then
            MsgBox "São permitidos somente números neste campo!*", vbInformation
            Exit Sub
        End If
        
        'Verificar se tem campos em branco
        If Me.TextName = "" Or Me.TextCpf = "" Then
            MsgBox "Não são permitidos campos em branco!", vbInformation
            Exit Sub
        End If
        
        If Len(Me.TextCpf.Value) < 11 And ActiveCell.Offset(0, 1).Value = "MOTORISTA" Then
            MsgBox "Digite a quantidade correta de números!", vbInformation
            Exit Sub
        End If
        
        If Len(Me.TextCpf.Value) > 9 And ActiveCell.Offset(0, 1).Value = "PLACA" Then
            MsgBox "Digite a quantidade correta de números!", vbInformation
            Exit Sub
        End If
        
        If Len(Me.TextCpf.Value) < 9 And ActiveCell.Offset(0, 1).Value = "PLACA" Then
            MsgBox "Digite a quantidade correta de números!", vbInformation
            Exit Sub
        End If
        
'        If Len(Me.TextCpf.Value) < 9 Or Len(Me.TextCpf.Value) < 11 Then
'            MsgBox "Digite a quantidade correta de números!"
'            Exit Sub
'        End If
    
        'remove espaços em branco a direita e esquerda do form
        'Me.TextName.Value = WorksheetFunction.Trim(Me.TextName)
        Me.TextCpf.Value = WorksheetFunction.Trim(Me.TextCpf)
        
        
        '==Realizar a edição dos dados
        ActiveCell.Value = Me.TextCpf.Text
        'converte para texto
        ActiveCell.NumberFormat = "@"
        'remove espaços em branco a direita e esquerda
        ActiveCell.Value = WorksheetFunction.Trim(ActiveCell)
    GoTo continue
    End If
       
continue:

    Range("A1").Select

    '*******
    ThisWorkbook.Save
    
    MsgBox ("Registro editado com sucesso!"), vbInformation
    
    Unload Me
    
    'Me.TextName = ""
    'Me.TextCpf = ""
    'Me.TextName.SetFocus

    
End Sub

Private Sub TextName_Change()
    If Me.TextName = "" Then
        MsgBox "Não são permitidos campos em branco!", vbInformation
        Exit Sub
    End If
End Sub

Private Sub TextName_Enter()
    Me.TextCpf.Enabled = False
    Me.TextCpf.BackColor = &H80000004
    Me.Label2.ForeColor = &HC0C0C0
End Sub



Private Sub TextCpf_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Cr = KeyAscii
    Módulo2.bloquerCaractere
    On Error Resume Next
    KeyAscii = Valor

End Sub
