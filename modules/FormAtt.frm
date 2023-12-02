VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormAtt 
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15000
   OleObjectBlob   =   "FormAtt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormAtt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PlanCad, PlanInicio As Worksheet
Dim fim As Long

Private Sub btnAtt_Click()

    Set PlanCad = Sheets("Quadrinhos Cadastrados")
    Set PlanInicio = Sheets("Inicial")

    Application.ScreenUpdating = False
    Rows(ActiveCell.Row).Delete
    PlanCad.Activate
    PlanCad.Range("A2").Select
    fim = PlanCad.UsedRange.Rows.Count + 1
                        
    PlanCad.Range("A" & fim).Value = Me.lblID.Caption
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
    Unload FormAtt
    MsgBox "Alterado com sucesso!", vbInformation, "Aviso"

End Sub
