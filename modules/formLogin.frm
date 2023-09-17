VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formLogin 
   Caption         =   "Login"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5640
   OleObjectBlob   =   "formLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PlanUsers, PlanInicio As Worksheet
Dim temp1, temp2, conf

Private Sub btnCad_Log_Click()
    
    Unload formLogin
    formCadastro.Show
    Unload formCadastro
    formLogin.Show
    
End Sub

Private Sub btnEntrar_Log_Click()
    Application.ScreenUpdating = False
    
    Set PlanUsers = Sheets("Usu�rios Cadastrados")
    Set PlanInicio = Sheets("Inicial")
    temp1 = UCase(Me.txtUser_Log.Value)
    conf = 0
    
    PlanUsers.Activate
    Range("A1").Select
    
    If Me.txtUser_Log = "" Or Me.txtSenha_Log = "" Then
    
        MsgBox "Digite seu usu�rio e sua senha!", vbOKOnly + vbExclamation, "Aviso"
        PlanInicio.Activate
        Exit Sub
    End If
    
    Do While conf = 0
        
        ActiveCell.Offset(1, 0).Select
        temp2 = UCase(ActiveCell.Value)
       
        If temp2 = temp1 Then
        
            conf = 1
        ElseIf temp2 = "" Then
    
            conf = 2
        End If
    Loop
    
    If conf = 1 Then
        
        If ActiveCell.Offset(0, 1).Value = Me.txtSenha_Log Then
        
            MsgBox "Login realizado com sucesso!", vbOKOnly + vbInformation, "Aviso"
            'abrir futuro form
            conf = 3
        Else
        
            MsgBox "Usu�rio e/ou senha incorretos!", vbOKOnly + vbExclamation, "Aviso"
            Me.txtUser_Log.Value = ""
            Me.txtSenha_Log.Value = ""
        End If
        
    Else
        
        MsgBox "Usu�rio e/ou senha incorretos!", vbOKOnly + vbExclamation, "Aviso"
        Me.txtUser_Log.Value = ""
        Me.txtSenha_Log.Value = ""
    End If
    
    PlanInicio.Activate
    Me.txtUser_Log.SetFocus
    
    If conf = 3 Then
    
        Unload Me
    End If
    
    Application.ScreenUpdating = True
End Sub

Private Sub btnMostrarSenha_Log_Click()
    
    If Me.btnMostrarSenha_Log.Value Then
    
        Me.txtSenha_Log.PasswordChar = ""
    Else
    
        Me.txtSenha_Log.PasswordChar = "*"
    End If
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then
    
        'Cancel = True
    End If

End Sub
