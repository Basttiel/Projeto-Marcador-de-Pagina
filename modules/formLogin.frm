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
Dim temp1, temp2, conf, user, inpbox, resp
Dim senha As String

Private Sub btnCad_Log_Click()
    
    Unload formLogin
    formCadastro.Show
    Unload formCadastro
    formLogin.Show
    
End Sub

Private Sub btnEntrar_Log_Click()
    Application.ScreenUpdating = False
    
    Set PlanUsers = Sheets("Usuários Cadastrados")
    Set PlanInicio = Sheets("Inicial")
    temp1 = UCase(Me.txtUser_Log.Value)
    conf = 0
    
    PlanUsers.Activate
    Range("A1").Select
    
    If Me.txtUser_Log = "" Or Me.txtSenha_Log = "" Then
    
        MsgBox "Digite seu usuário e sua senha!", vbOKOnly + vbExclamation, "Aviso"
        PlanInicio.Activate
        Exit Sub
    End If
    
    Do While conf = 0
        
        ActiveCell.Offset(1, 0).Select
        temp2 = UCase(ActiveCell.Value)
       
        If temp2 = temp1 Then
        
            conf = 1
            user = ActiveCell.Value
        ElseIf temp2 = "" Then
    
            conf = 2
        End If
    Loop
    
    If conf = 1 Then
        senha = ActiveCell.Offset(0, 1).Value
        
        If senha = Me.txtSenha_Log Then
        
            'MsgBox "Login realizado com sucesso!", vbOKOnly + vbInformation, "Aviso"
            PlanInicio.Activate
            PlanInicio.Range("a1").Value = user
            ActiveWorkbook.RefreshAll
            Unload formLogin
            formUser.Show
            Unload formUser
            formLogin.Show
            Exit Sub
            conf = 3
        Else
        
            MsgBox "Usuário e/ou senha incorretos!", vbOKOnly + vbExclamation, "Aviso"
            Me.txtUser_Log.Value = ""
            Me.txtSenha_Log.Value = ""
        End If
        
    Else
        
        MsgBox "Usuário e/ou senha incorretos!", vbOKOnly + vbExclamation, "Aviso"
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

Private Sub btnEsqSenha_Log_Click()

    Set PlanUsers = Sheets("Usuários Cadastrados")
    Set PlanInicio = Sheets("Inicial")
    conf = 0
      
    inpbox = UCase(InputBox("Digite o seu nome de usuário!", "Esqueci minha senha"))
    
    If inpbox <> "" Then
        PlanUsers.Activate
        Range("A1").Select
        
        Do While conf = 0
            
            ActiveCell.Offset(1, 0).Select
            temp2 = UCase(ActiveCell.Value)
           
            If temp2 = inpbox Then
            
                conf = 1
                user = ActiveCell.Value
            ElseIf temp2 = "" Then
        
                conf = 2
            End If
        Loop
    End If
    
    If conf = 1 Then
        resp = InputBox(ActiveCell.Offset(0, 2).Value, "Pergunta de Segurança")
    
        If UCase(resp) = UCase(ActiveCell.Offset(0, 3).Value) Then
            MsgBox "Sua senha é: " & ActiveCell.Offset(0, 1).Value, vbOKOnly, "Senha"
            
        Else
            MsgBox "Resposta Incorreta!"
            
        End If
    
    Else
        MsgBox "Usuário não encontrado!", vbExclamation
        
    End If
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
    
        Cancel = True
    End If

End Sub
