Attribute VB_Name = "M�dulo2"
Public Cr As Double
Public Valor As String

Sub bloquerCaractere()

Valor = ""

If (Cr < 45 Or Cr > 57) Or Cr = 45 Or Cr = 46 Or Cr = 47 Then

    If Cr <> 8 Then
        If Cr <> 13 Then
        
        Valor = 0
        
        MsgBox "S�o permitidos somente n�meros neste campo!", vbInformation, "ERRO"
        
        End If
    End If
End If


End Sub
