Attribute VB_Name = "Módulo1"
Option Explicit

Dim PlanUsers, PlanInicio, PlanCad As Worksheet
Dim temp1, temp2 As String
Dim frm, lin, i, temp, tbl, rng

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
    Application.Run "Módulo1.Ordenar"
    PlanCad.Activate
    Range("H2").Select
    temp = 0
    For i = formUser.listQuad.ListCount - 1 To 0 Step -1
        formUser.listQuad.RemoveItem i
    Next i
    'formUser.listQuad.List = Array()
    
    Do While ActiveCell.Value <> ""
    
    If ActiveCell.Value = formUser.lblUser.Caption Then
    
        If formUser.txtPesq = "" Or InStr(1, ActiveCell.Offset(0, -6).Value, formUser.txtPesq.Value, vbTextCompare) > 0 Then
        
            With formUser.listQuad
                
                .AddItem
                .List(temp, 0) = PlanCad.Cells(lin, 1)
                .List(temp, 1) = PlanCad.Cells(lin, 2)
                .List(temp, 2) = PlanCad.Cells(lin, 3)
                .List(temp, 3) = PlanCad.Cells(lin, 4)
                .List(temp, 4) = PlanCad.Cells(lin, 5)
                .List(temp, 5) = PlanCad.Cells(lin, 6)
                .List(temp, 6) = PlanCad.Cells(lin, 7)
                temp = temp + 1
                
            End With
        
        End If
    
    End If
    
    lin = lin + 1
    ActiveCell.Offset(1, 0).Select
    Loop
    
End Sub

'Public Sub MostrarCom()

    'Set PlanCad = Sheets("Quadrinhos Cadastrados")

'End Sub

Sub Ordenar()

    ActiveWorkbook.Worksheets("Quadrinhos Cadastrados").ListObjects("tabQuad").Sort _
        .SortFields.Clear
    ActiveWorkbook.Worksheets("Quadrinhos Cadastrados").ListObjects("tabQuad").Sort _
        .SortFields.Add2 Key:=Range("tabQuad[[#All],[nome]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Quadrinhos Cadastrados").ListObjects("tabQuad") _
        .Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
