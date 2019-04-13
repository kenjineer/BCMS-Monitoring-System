Attribute VB_Name = "Module1"
'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-'
'    Title: BCMS Project Monitoring (All Sites)     '
'    Author: Engr. Kenneth Caro Karamihan           '
'    Company: PLDT Inc.                             '
'    Division: BC Governance and Reporting          '
'    Date: May 24, 2018                             '
'    Code version: 2.0                              '
'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-'

Sub Status()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Dim rng As Integer
    Dim stat_null, stat_com, stat_bs, stat_as As Boolean
    Dim result As String
    Dim cell As Range
    
    If ActiveCell.Row > 2 Then
        rng = ActiveCell.Row - 1
    Else
        rng = ActiveCell.Row
    End If
    
    Set cell = Cells(rng, 14)
    
    stat_null = Evaluate(cell.FormatConditions.Item(1).Formula1)
    stat_com = Evaluate(cell.FormatConditions.Item(2).Formula1)
    stat_bs = Evaluate(cell.FormatConditions.Item(3).Formula1)
    stat_as = Evaluate(cell.FormatConditions.Item(4).Formula1)
    
    
    If stat_null = True Then
        result = ""
    ElseIf stat_com = False And stat_as = False Then
        result = "BEHIND SCHEDULE"
        cell.Font.Size = 11
    ElseIf stat_bs = False And stat_as = False Then
        result = "ON TIME"
        cell.Font.Size = 12
    ElseIf stat_as = True Then
        result = "AHEAD SCHEDULE"
        cell.Font.Size = 11
    Else
    End If
    cell.Value = result
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


