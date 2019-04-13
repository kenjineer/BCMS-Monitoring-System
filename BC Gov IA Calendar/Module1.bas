Attribute VB_Name = "Module1"
'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+'
'    Title: BC Gov IA Calendar Automatic Email and Deadline Notifier     '
'    Author: Engr. Kenneth Caro Karamihan                                '
'    Company: PLDT Inc.                                                  '
'    Division: BC Governance and Reporting                               '
'    Date: May 7, 2018                                                   '
'    Code version: 1.0                                                   '
'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+'

Sub Notif()
    Dim i As Integer
    With Status
        .lst_L.Clear
        .lst_B.Clear
        .lst_DC.Clear
        .lst_DL.Clear
        .lst_G.Clear
        .lst_S.Clear
        .lst_GH.Clear
        .lst_C.Clear
        .lst_L.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_B.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_DC.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_DL.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_G.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_S.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_GH.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_C.AddItem "DUE DATES NEAR EXPIRATION:"
        .chk_APFMDL.Enabled = False
        .chk_APFMG.Enabled = False
        .chk_APFMS.Enabled = False
        .chk_APFMGH.Enabled = False
        .chk_APFMC.Enabled = False
    End With
    For i = 3 To 21
        If Sheet1.Cells(i, 2).Value = "-" Then
        ElseIf Date = Sheet1.Cells(i, 2).Value - 1 Or Date = Sheet1.Cells(i, 2).Value - 3 Or Date = Sheet1.Cells(i, 2).Value - 7 Then
            Status.lst_L.AddItem Sheet1.Cells(i, 2).Value & "      " & Sheet1.Cells(i, 1).Value
        End If
    Next i
    For i = 3 To 21
        If Sheet1.Cells(i, 3).Value = "-" Then
        ElseIf Date = Sheet1.Cells(i, 3).Value - 1 Or Date = Sheet1.Cells(i, 3).Value - 3 Or Date = Sheet1.Cells(i, 3).Value - 7 Then
            Status.lst_B.AddItem Sheet1.Cells(i, 3).Value & "      " & Sheet1.Cells(i, 1).Value
        End If
    Next i
    For i = 3 To 21
        If Sheet1.Cells(i, 4).Value = "-" Then
        ElseIf Date = Sheet1.Cells(i, 4).Value - 1 Or Date = Sheet1.Cells(i, 4).Value - 3 Or Date = Sheet1.Cells(i, 4).Value - 7 Then
            Status.lst_DC.AddItem Sheet1.Cells(i, 4).Value & "      " & Sheet1.Cells(i, 1).Value
        End If
    Next i
    For i = 3 To 21
        If Sheet1.Cells(i, 5).Value = "-" Then
        ElseIf Date = Sheet1.Cells(i, 5).Value - 1 Or Date = Sheet1.Cells(i, 5).Value - 3 Or Date = Sheet1.Cells(i, 5).Value - 7 Then
            Status.lst_DL.AddItem Sheet1.Cells(i, 5).Value & "      " & Sheet1.Cells(i, 1).Value
        End If
    Next i
    For i = 3 To 21
        If Sheet1.Cells(i, 6).Value = "-" Then
        ElseIf Date = Sheet1.Cells(i, 6).Value - 1 Or Date = Sheet1.Cells(i, 6).Value - 3 Or Date = Sheet1.Cells(i, 6).Value - 7 Then
            Status.lst_G.AddItem Sheet1.Cells(i, 6).Value & "      " & Sheet1.Cells(i, 1).Value
        End If
    Next i
    For i = 3 To 21
        If Sheet1.Cells(i, 7).Value = "-" Then
        ElseIf Date = Sheet1.Cells(i, 7).Value - 1 Or Date = Sheet1.Cells(i, 7).Value - 3 Or Date = Sheet1.Cells(i, 7).Value - 7 Then
            Status.lst_S.AddItem Sheet1.Cells(i, 7).Value & "      " & Sheet1.Cells(i, 1).Value
        End If
    Next i
    For i = 3 To 21
        If Sheet1.Cells(i, 8).Value = "-" Then
        ElseIf Date = Sheet1.Cells(i, 8).Value - 1 Or Date = Sheet1.Cells(i, 8).Value - 3 Or Date = Sheet1.Cells(i, 8).Value - 7 Then
            Status.lst_GH.AddItem Sheet1.Cells(i, 8).Value & "      " & Sheet1.Cells(i, 1).Value
        End If
    Next i
    For i = 3 To 21
        If Sheet1.Cells(i, 9).Value = "-" Then
        ElseIf Date = Sheet1.Cells(i, 9).Value - 1 Or Date = Sheet1.Cells(i, 9).Value - 3 Or Date = Sheet1.Cells(i, 9).Value - 7 Then
            Status.lst_C.AddItem Sheet1.Cells(i, 9).Value & "      " & Sheet1.Cells(i, 1).Value
        End If
    Next i
    
    If Status.lst_L.ListCount >= 2 Then
        Status.MultiPage1.Pages(0).Caption = "*" + Status.MultiPage1.Pages(0).Caption
        Status.Image1.Visible = True
    Else
        Status.Image1.Visible = False
    End If
    If Status.lst_B.ListCount >= 2 Then
        Status.MultiPage1.Pages(1).Caption = "*" + Status.MultiPage1.Pages(1).Caption
        Status.Image2.Visible = True
    Else
        Status.Image2.Visible = False
    End If
    If Status.lst_DC.ListCount >= 2 Then
        Status.MultiPage1.Pages(2).Caption = "*" + Status.MultiPage1.Pages(2).Caption
        Status.Image3.Visible = True
    Else
        Status.Image3.Visible = False
    End If
    If Status.lst_DL.ListCount >= 2 Then
        Status.MultiPage1.Pages(3).Caption = "*" + Status.MultiPage1.Pages(3).Caption
        Status.Image4.Visible = True
        Status.chk_APFMDL.Enabled = True
    Else
        Status.Image4.Visible = False
        Status.chk_APFMDL.Enabled = False
    End If
    If Status.lst_G.ListCount >= 2 Then
        Status.MultiPage1.Pages(4).Caption = "*" + Status.MultiPage1.Pages(4).Caption
        Status.Image5.Visible = True
        Status.chk_APFMG.Enabled = True
    Else
        Status.Image5.Visible = False
        Status.chk_APFMG.Enabled = False
    End If
    If Status.lst_S.ListCount >= 2 Then
        Status.MultiPage1.Pages(5).Caption = "*" + Status.MultiPage1.Pages(5).Caption
        Status.Image6.Visible = True
        Status.chk_APFMS.Enabled = True
    Else
        Status.Image6.Visible = False
        Status.chk_APFMS.Enabled = False
    End If
    If Status.lst_GH.ListCount >= 2 Then
        Status.MultiPage1.Pages(6).Caption = "*" + Status.MultiPage1.Pages(6).Caption
        Status.Image7.Visible = True
        Status.chk_APFMGH.Enabled = True
    Else
        Status.Image7.Visible = False
        Status.chk_APFMGH.Enabled = False
    End If
    If Status.lst_C.ListCount >= 2 Then
        Status.MultiPage1.Pages(7).Caption = "*" + Status.MultiPage1.Pages(7).Caption
        Status.Image8.Visible = True
        Status.chk_APFMC.Enabled = True
    Else
        Status.Image8.Visible = False
        Status.chk_APFMC.Enabled = False
    End If
    
    Status.Show
End Sub
