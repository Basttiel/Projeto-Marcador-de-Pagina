Attribute VB_Name = "Módulo1"
Option Explicit

Dim PlanUsers, PlanInicio As Worksheet
Dim temp1, temp2 As String
Dim frm

Public Sub ValidarUser()
    Application.ScreenUpdating = False
    
    Set PlanUsers = Sheets("Usuários Cadastrados")
    Set PlanInicio = Sheets("Inicial")
    
    PlanUsers.Activate
    Range("A1").Select

    Do While ActiveCell.Value <> ""
        
        temp1 = UCase(ActiveCell.Value)
        temp2 = UCase(formCadastro.txtUser_Cad.Value)
        
        If temp1 = temp2 Then
        
            MsgBox "Usuário já cadastrado!", vbOKOnly + vbExclamation, "Aviso"
            PlanInicio.Range("B1").Value = 1
            formCadastro.txtUser_Cad.SetFocus
            
            Exit Sub
        
        End If
        
        ActiveCell.Offset(1, 0).Select
    
    Loop
    
    
    PlanInicio.Activate
    PlanInicio.Range("B1").Value = ""
    
    Application.ScreenUpdating = True
End Sub
