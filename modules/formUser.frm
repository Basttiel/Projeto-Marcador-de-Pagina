VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formUser 
   ClientHeight    =   7080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16695
   OleObjectBlob   =   "formUser.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim PlanCad, PlanInicio As Worksheet
    Dim fim As Long
    Dim id As Integer
    Dim tabela As ListObject
    Dim user, lin, resp, temp

Private Sub btnAtt_Click()
    
    Set PlanCad = Sheets("Quadrinhos Cadastrados")

    If Me.listQuad.ListIndex < 0 Then
        MsgBox "Selecione na lista o que deseja editar!", vbInformation, "Aviso"
    
    Else
        
        PlanCad.Activate
        Range("A1").Select
        
        Do While ActiveCell.Value <> CInt(Me.listQuad.List(Me.listQuad.ListIndex, 0)) Or ActiveCell.Value = Null
        ActiveCell.Offset(1, 0).Select
            
        Loop
                
        If ActiveCell.Value = "" Then
                MsgBox "Não encontrado!", vbCritical
            
        Else
            FormAtt.lblUser.Caption = Me.lblUser.Caption
            FormAtt.lblID.Caption = CInt(Me.listQuad.List(Me.listQuad.ListIndex, 0))
            FormAtt.txtNome.Value = ActiveCell.Offset(0, 1).Value
            FormAtt.txtMarc.Value = ActiveCell.Offset(0, 2).Value
            FormAtt.txtFonte.Value = ActiveCell.Offset(0, 3).Value
            If ActiveCell.Offset(0, 4).Value = "Lendo" Then
                FormAtt.btnopLendo.Value = True
            
            ElseIf ActiveCell.Offset(0, 4).Value = "Completo" Then
                FormAtt.btnopComp.Value = True
            
            Else
                FormAtt.btnopPlan.Value = True
            
            End If
            FormAtt.txtNota = ActiveCell.Offset(0, 5).Value
            FormAtt.txtComen = ActiveCell.Offset(0, 6).Value
            FormAtt.Show
                    
        End If
               
    End If
    
End Sub

Private Sub btnCad_Click()
    Set PlanCad = Sheets("Quadrinhos Cadastrados")
    Set PlanInicio = Sheets("Inicial")
    id = PlanCad.Range("id").Value

    Application.ScreenUpdating = False
    PlanCad.Activate
    PlanCad.Range("A2").Select
    fim = PlanCad.UsedRange.Rows.Count + 1
    PlanCad.Range("id") = id + 1
                        
    PlanCad.Range("A" & fim).Value = PlanCad.Range("k1")
    PlanCad.Range("B" & fim).Value = Me.txtNome.Value
    PlanCad.Range("C" & fim).Value = Me.txtMarc.Value
    PlanCad.Range("D" & fim).Value = Me.txtFonte.Value
    
    If Me.btnopLendo = True Then
        PlanCad.Range("E" & fim).Value = Me.btnopLendo.Caption
    
    ElseIf Me.btnopComp = True Then
        PlanCad.Range("E" & fim).Value = Me.btnopComp.Caption
    
    ElseIf Me.btnopPlan = True Then
        PlanCad.Range("E" & fim).Value = Me.btnopPlan.Caption
    
    End If
    
    PlanCad.Range("F" & fim).Value = Me.txtNota.Value
    PlanCad.Range("G" & fim).Value = Me.txtComen.Value
    PlanCad.Range("H" & fim).Value = Me.lblUser.Caption
    
    Application.Run "Módulo1.AttLista"
    PlanInicio.Activate
    Application.ScreenUpdating = True
    ActiveWorkbook.RefreshAll
    
    Me.txtNome = ""
    Me.txtMarc = ""
    Me.txtFonte = ""
    Me.txtNota = ""
    Me.txtComen = ""
    Me.btnopComp = 0
    Me.btnopLendo = 0
    Me.btnopPlan = 0
    
    MsgBox "Cadastrado com sucesso!", vbOKOnly + vbInformation, "Aviso"

End Sub

Private Sub btnComen_Click()

    Set PlanCad = Sheets("Quadrinhos Cadastrados")

    If Me.listQuad.ListIndex < 0 Then
        MsgBox "Selecione a opção que deseja ver o comentário!", vbInformation, "Aviso"
        
    Else
        PlanCad.Activate
        Range("A1").Select
        
        Do While ActiveCell.Value <> CInt(Me.listQuad.List(Me.listQuad.ListIndex, 0)) Or ActiveCell.Value = Null
            ActiveCell.Offset(1, 0).Select
            
        Loop
        
        If ActiveCell.Value = "" Then
            MsgBox "Não encontrado!", vbCritical
            
        Else
            MsgBox ActiveCell.Offset(0, 6).Value, , "Comentário"
        
        End If
        
    End If

End Sub

Private Sub btnDel_Click()

    Set PlanCad = Sheets("Quadrinhos Cadastrados")

    If Me.listQuad.ListIndex < 0 Then
        MsgBox "Selecione na lista o que deseja excluir!", vbInformation, "Aviso"
    
    Else
        resp = MsgBox("Tem certeza que deseja excluir este dado?", vbYesNo + vbExclamation, "ALERTA")
            
        If resp = vbYes Then
            PlanCad.Activate
            Range("A1").Select
        
            Do While ActiveCell.Value <> CInt(Me.listQuad.List(Me.listQuad.ListIndex, 0)) Or ActiveCell.Value = Null
            ActiveCell.Offset(1, 0).Select
            
            Loop
                
            If ActiveCell.Value = "" Then
                MsgBox "Não encontrado!", vbCritical
            
            Else
                Rows(ActiveCell.Row).Delete
                Application.Run "Módulo1.AttLista"
                MsgBox "Deletado com sucesso!", vbInformation, "Aviso"
                    
            End If
        End If
    End If

End Sub

Private Sub btnPesq_Click()

    Application.Run "Módulo1.AttLista"

End Sub

Private Sub UserForm_Activate()
    
    Set PlanInicio = Sheets("Inicial")
    user = PlanInicio.Cells(1, 1).Value
    
    Me.lblInicio.Caption = "Bem Vindo " & user & "!"
    Me.lblUser.Caption = user
    
    Application.Run "Módulo1.AttLista"

End Sub

