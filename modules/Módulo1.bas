Attribute VB_Name = "M�dulo1"
Option Explicit

Dim PlanUsers, PlanInicio, PlanCad As Worksheet
Dim temp1, temp2 As String
Dim frm, lin, i, temp

Public Sub ValidarUser()
    Application.ScreenUpdating = False
    
    Set PlanUsers = Sheets("Usu�rios Cadastrados")
    Set PlanInicio = Sheets("Inicial")
    
    PlanUsers.Activate
    Range("A1").Select

    Do While ActiveCell.Value <> ""
        
        temp1 = UCase(ActiveCell.Value)
        temp2 = UCase(formCadastro.txtUser_Cad.Value)
        
        If temp1 = temp2 Then
        
            MsgBox "Usu�rio j� cadastrado!", vbOKOnly + vbExclamation, "Aviso"
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

Public Sub AttLista()

    Set PlanCad = Sheets("Quadrinhos Cadastrados")
    lin = 2
    
    'Do Until PlanCad.Cells(lin, 1) = ""

       ' With formUser.listQuad
            
           ' .AddItem
            '.List(lin - 2, 0) = PlanCad.Cells(lin, 2)
            '.List(lin - 2, 1) = PlanCad.Cells(lin, 3)
            '.List(lin - 2, 2) = PlanCad.Cells(lin, 4)
            '.List(lin - 2, 3) = PlanCad.Cells(lin, 5)
            '.List(lin - 2, 4) = PlanCad.Cells(lin, 6)
            '.List(lin - 2, 5) = PlanCad.Cells(lin, 7)
            
        'End With
        
       'lin = lin + 1
    'Loop
    
    
    PlanCad.Activate
    Range("H2").Select
    temp = 0
    
    Do While ActiveCell.Value <> ""
    
    If ActiveCell.Value = formUser.lblUser.Caption Then
    
        With formUser.listQuad
            
            .AddItem
            .List(temp, 0) = PlanCad.Cells(lin, 2)
            .List(temp, 1) = PlanCad.Cells(lin, 3)
            .List(temp, 2) = PlanCad.Cells(lin, 4)
            .List(temp, 3) = PlanCad.Cells(lin, 5)
            .List(temp, 4) = PlanCad.Cells(lin, 6)
            .List(temp, 5) = PlanCad.Cells(lin, 7)
            temp = temp + 1
            
        End With
    
    End If
    
    lin = lin + 1
    ActiveCell.Offset(1, 0).Select
    Loop
End Sub
