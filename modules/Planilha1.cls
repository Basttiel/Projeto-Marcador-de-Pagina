VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
    
    Range("A1").Select
    
    If formCadastro.Visible = False And formLogin.Visible = False And formUser.Visible = False Then
        
        On Error Resume Next
        formLogin.Show
        On Error GoTo 0
    End If
    
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    
End Sub
