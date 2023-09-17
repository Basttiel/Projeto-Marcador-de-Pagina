VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formCadastro 
   Caption         =   "Cadastro"
   ClientHeight    =   9000.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5640
   OleObjectBlob   =   "formCadastro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Dim PlanCad, PlanUsers, PlanInicio, NovaPlanilha As Worksheet
    Dim fim As Long

Private Sub btnCadastrar_Cad_Click()
    Set PlanCad = Sheets("Usuários Cadastrados")
    Set PlanInicio = Sheets("Inicial")
    
    Application.Run "Módulo1.ValidarUser"
    
    If PlanInicio.Range("B1").Value = 1 Then
    
        Exit Sub
    End If
    
    If Me.txtSenha_Cad <> Me.txtConfSenha_Cad Then
    
        MsgBox "As senhas não coincidem!", vbOKOnly + vbExclamation, "Aviso"
        Me.txtSenha_Cad.SetFocus
    
            ElseIf Me.txtUser_Cad = "" Or Me.txtSenha_Cad = "" Or Me.txtConfSenha_Cad = "" Or Me.txtPerg_Cad = "" Or Me.txtResp_Cad = "" Then
    
            MsgBox "Todos os campos não foram preenchidos!", vbOKOnly + vbExclamation, "Aviso"
            Me.txtUser_Cad.SetFocus
            
                Else
                    Application.ScreenUpdating = False
                    PlanCad.Activate
                    PlanCad.Range("A2").Select
                    fim = PlanCad.UsedRange.Rows.Count + 1
                        
                    PlanCad.Range("A" & fim).Value = Me.txtUser_Cad.Value
                    PlanCad.Range("B" & fim).Value = Me.txtSenha_Cad.Value
                    PlanCad.Range("C" & fim).Value = Me.txtPerg_Cad.Value
                    PlanCad.Range("D" & fim).Value = Me.txtResp_Cad.Value
                    
                    'Set NovaPlanilha = ThisWorkbook.Sheets.Add
                    'NovaPlanilha.Name = Me.txtUser_Cad.Value
                    PlanInicio.Activate
                    Application.ScreenUpdating = True
                    ActiveWorkbook.RefreshAll
                    Unload Me
                    MsgBox "Usuário cadastrado com sucesso!", vbOKOnly + vbInformation, "Aviso"
                    'formLogin.Show
    End If

End Sub

Private Sub btnMostrarSenha1_Cad_Click()

    If Me.btnMostrarSenha1_Cad.Value Then
    
        Me.txtSenha_Cad.PasswordChar = ""
    Else
    
        Me.txtSenha_Cad.PasswordChar = "*"
    End If

End Sub

Private Sub btnMostrarSenha2_Cad_Click()

    If Me.btnMostrarSenha2_Cad.Value Then
    
        Me.txtConfSenha_Cad.PasswordChar = ""
    Else
    
        Me.txtConfSenha_Cad.PasswordChar = "*"
    End If

End Sub

Private Sub btnVeriUser_Cad_Click()
    
    Set PlanInicio = Sheets("Inicial")
    
    Application.Run "Módulo1.ValidarUser"
    
    If PlanInicio.Range("B1").Value <> 1 Then
        MsgBox "Usuário disponível!", vbOKOnly + vbInformation, "Aviso"
        Me.txtSenha_Cad.SetFocus
    End If
    
End Sub
