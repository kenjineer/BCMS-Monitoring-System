Attribute VB_Name = "Module2"
'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+'
'    Title: BCMS Docs Monitoring Automatic Email and Deadline Notifier     '
'    Author: Engr. Kenneth Caro Karamihan                                  '
'    Company: PLDT Inc.                                                    '
'    Division: BC Governance and Reporting                                 '
'    Date: May 22, 2018                                                    '
'    Code version: 1.0                                                     '
'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+'
    
Sub Init()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Dim c, rng As Range
    Dim stat_black, stat_red, stat_orange As Boolean
    With Notif
        .lst_GOV.Clear
        .lst_STRAT.Clear
        .lst_ENG.Clear
        .lst_LUCLS.Clear
        .lst_BCLS.Clear
        .lst_DCLS.Clear
        .lst_NGMMFxATOp.Clear
        .lst_FCNOD.Clear
        .lst_IGOD.Clear
        .lst_WGMMFxATOp.Clear
        .lst_FCNOS.Clear
        .lst_IGOS.Clear
        .lst_EGMMGHL.Clear
        .lst_FCNOGH.Clear
        .lst_EGMMGNT.Clear
        .lst_FCNOGNT.Clear
        .lst_VisFxATOp.Clear
        .lst_FCNOJNE.Clear
        .lst_STRATC.Clear
        .lst_ENGC.Clear
        .lst_GOV.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_STRAT.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_ENG.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_LUCLS.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_BCLS.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_DCLS.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_NGMMFxATOp.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_FCNOD.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_IGOD.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_WGMMFxATOp.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_FCNOS.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_IGOS.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_EGMMGHL.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_FCNOGH.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_EGMMGNT.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_FCNOGNT.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_VisFxATOp.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_FCNOJNE.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_STRATC.AddItem "DUE DATES NEAR EXPIRATION:"
        .lst_ENGC.AddItem "DUE DATES NEAR EXPIRATION:"
    End With
    
    'GOV
    Sheet2.Activate
    Set rng = Sheet2.Range("G4:G11,G35:G47,G49")
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_GOV.AddItem "3 Months Before Deadline:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_GOV.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_orange = False Then
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_GOV.AddItem "For Urgent Update:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_GOV.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_red = False Then
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_GOV.AddItem "Outdated:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_GOV.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_black = False Then
        End If
    Next
    
    'STRAT
    Set rng = Sheet2.Range("G3,G12:G22,G29:G30")
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_STRAT.AddItem "3 Months Before Deadline:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_STRAT.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_orange = False Then
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_STRAT.AddItem "For Urgent Update:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_STRAT.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_red = False Then
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_STRAT.AddItem "Outdated:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_STRAT.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_black = False Then
        End If
    Next
    
    'ENG
    Set rng = Sheet2.Range("G23:G28,G31:G34,G48")
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_ENG.AddItem "3 Months Before Deadline:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_ENG.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_orange = False Then
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_ENG.AddItem "For Urgent Update:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_ENG.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_red = False Then
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_ENG.AddItem "Outdated:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_ENG.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_black = False Then
        End If
    Next
    
    'LUCLS
    Sheet3.Activate
    Set rng = Sheet3.Range("G3:G8,G10:G23,G25:G48")
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_LUCLS.AddItem "3 Months Before Deadline:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If c.Row >= 10 And c.Row <= 21 And stat_orange = True Then
            Notif.lst_LUCLS.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(9, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 25 And c.Row <= 47 And stat_orange = True Then
            Notif.lst_LUCLS.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(24, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf stat_orange = True Then
            Notif.lst_LUCLS.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_orange = False Then
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_LUCLS.AddItem "For Urgent Update:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If c.Row >= 10 And c.Row <= 21 And stat_red = True Then
            Notif.lst_LUCLS.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(9, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 25 And c.Row <= 47 And stat_red = True Then
            Notif.lst_LUCLS.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(24, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf stat_red = True Then
            Notif.lst_LUCLS.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_red = False Then
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_LUCLS.AddItem "Outdated:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If c.Row >= 10 And c.Row <= 21 And stat_black = True Then
            Notif.lst_LUCLS.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(9, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 25 And c.Row <= 47 And stat_black = True Then
            Notif.lst_LUCLS.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(24, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf stat_black = True Then
            Notif.lst_LUCLS.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_black = False Then
        End If
    Next
    
    'BCLS
    Set rng = Sheet3.Range("Q3:Q8,Q10:Q28,Q30:Q53")
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_BCLS.AddItem "3 Months Before Deadline:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If c.Row >= 10 And c.Row <= 26 And stat_orange = True Then
            Notif.lst_BCLS.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(9, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf c.Row >= 30 And c.Row <= 52 And stat_orange = True Then
            Notif.lst_BCLS.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(29, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf stat_orange = True Then
            Notif.lst_BCLS.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(c.Row, "M").Text
        ElseIf stat_orange = False Then
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_BCLS.AddItem "For Urgent Update:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If c.Row >= 10 And c.Row <= 26 And stat_red = True Then
            Notif.lst_BCLS.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(9, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf c.Row >= 30 And c.Row <= 52 And stat_red = True Then
            Notif.lst_BCLS.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(29, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf stat_red = True Then
            Notif.lst_BCLS.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(c.Row, "M").Text
        ElseIf stat_red = False Then
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_BCLS.AddItem "Outdated:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If c.Row >= 10 And c.Row <= 26 And stat_black = True Then
            Notif.lst_BCLS.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(9, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf c.Row >= 30 And c.Row <= 52 And stat_black = True Then
            Notif.lst_BCLS.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(29, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf stat_black = True Then
            Notif.lst_BCLS.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(c.Row, "M").Text
        ElseIf stat_black = False Then
        End If
    Next
    
    'DCLS
    Set rng = Sheet3.Range("AA3:AA8,AA10:AA22,AA24:AA47")
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_DCLS.AddItem "3 Months Before Deadline:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If c.Row >= 10 And c.Row <= 20 And stat_orange = True Then
            Notif.lst_DCLS.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(9, "W").Text & ": " & Cells(c.Row, "W").Text
        ElseIf c.Row >= 24 And c.Row <= 46 And stat_orange = True Then
            Notif.lst_DCLS.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(23, "W").Text & ": " & Cells(c.Row, "W").Text
        ElseIf stat_orange = True Then
            Notif.lst_DCLS.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(c.Row, "W").Text
        ElseIf stat_orange = False Then
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_DCLS.AddItem "For Urgent Update:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If c.Row >= 10 And c.Row <= 20 And stat_red = True Then
            Notif.lst_DCLS.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(9, "W").Text & ": " & Cells(c.Row, "W").Text
        ElseIf c.Row >= 24 And c.Row <= 46 And stat_red = True Then
            Notif.lst_DCLS.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(23, "W").Text & ": " & Cells(c.Row, "W").Text
        ElseIf stat_red = True Then
            Notif.lst_DCLS.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(c.Row, "W").Text
        ElseIf stat_red = False Then
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_DCLS.AddItem "Outdated:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If c.Row >= 10 And c.Row <= 20 And stat_black = True Then
            Notif.lst_DCLS.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(9, "W").Text & ": " & Cells(c.Row, "W").Text
        ElseIf c.Row >= 24 And c.Row <= 46 And stat_black = True Then
            Notif.lst_DCLS.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(23, "W").Text & ": " & Cells(c.Row, "W").Text
        ElseIf stat_black = True Then
            Notif.lst_DCLS.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(c.Row, "W").Text
        ElseIf stat_black = False Then
        End If
    Next
    
    'NGMMFxATOp
    Sheet4.Activate
    Set rng = Sheet4.Range("G3:G7,G9:G23,G25:G48")
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_NGMMFxATOp.AddItem "3 Months Before Deadline:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If c.Row >= 9 And c.Row <= 21 And stat_orange = True Then
            Notif.lst_NGMMFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(8, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 25 And c.Row <= 47 And stat_orange = True Then
            Notif.lst_NGMMFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(24, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf stat_orange = True Then
            Notif.lst_NGMMFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_orange = False Then
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_NGMMFxATOp.AddItem "For Urgent Update:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If c.Row >= 9 And c.Row <= 21 And stat_red = True Then
            Notif.lst_NGMMFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(8, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 25 And c.Row <= 47 And stat_red = True Then
            Notif.lst_NGMMFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(24, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf stat_red = True Then
            Notif.lst_NGMMFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_red = False Then
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_NGMMFxATOp.AddItem "Outdated:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If c.Row >= 9 And c.Row <= 21 And stat_black = True Then
            Notif.lst_NGMMFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(8, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 25 And c.Row <= 47 And stat_black = True Then
            Notif.lst_NGMMFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(24, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf stat_black = True Then
            Notif.lst_NGMMFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_black = False Then
        End If
    Next
    
    'FCNO FF2 QCY
    Set rng = Sheet4.Range("Q3:Q7,Q9:Q32,Q34:Q57")
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_FCNOD.AddItem "3 Months Before Deadline:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If c.Row >= 9 And c.Row <= 30 And stat_orange = True Then
            Notif.lst_FCNOD.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(8, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf c.Row >= 34 And c.Row <= 56 And stat_orange = True Then
            Notif.lst_FCNOD.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(33, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf stat_orange = True Then
            Notif.lst_FCNOD.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(c.Row, "M").Text
        ElseIf stat_orange = False Then
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_FCNOD.AddItem "For Urgent Update:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If c.Row >= 9 And c.Row <= 30 And stat_red = True Then
            Notif.lst_FCNOD.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(8, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf c.Row >= 34 And c.Row <= 56 And stat_red = True Then
            Notif.lst_FCNOD.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(33, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf stat_red = True Then
            Notif.lst_FCNOD.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(c.Row, "M").Text
        ElseIf stat_red = False Then
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_FCNOD.AddItem "Outdated:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If c.Row >= 9 And c.Row <= 30 And stat_black = True Then
            Notif.lst_FCNOD.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(8, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf c.Row >= 34 And c.Row <= 56 And stat_black = True Then
            Notif.lst_FCNOD.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(33, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf stat_black = True Then
            Notif.lst_FCNOD.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(c.Row, "M").Text
        ElseIf stat_black = False Then
        End If
    Next
    
    'IGO QC DFON Station
    Set rng = Sheet4.Range("AA3:AA7,AA9:AA19,AA21:AA44")
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_IGOD.AddItem "3 Months Before Deadline:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If c.Row >= 9 And c.Row <= 17 And stat_orange = True Then
            Notif.lst_IGOD.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(8, "W").Text & ": " & Cells(c.Row, "W").Text
        ElseIf c.Row >= 21 And c.Row <= 43 And stat_orange = True Then
            Notif.lst_IGOD.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(20, "W").Text & ": " & Cells(c.Row, "W").Text
        ElseIf stat_orange = True Then
            Notif.lst_IGOD.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(c.Row, "W").Text
        ElseIf stat_orange = False Then
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_IGOD.AddItem "For Urgent Update:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If c.Row >= 9 And c.Row <= 17 And stat_red = True Then
            Notif.lst_IGOD.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(8, "W").Text & ": " & Cells(c.Row, "W").Text
        ElseIf c.Row >= 21 And c.Row <= 43 And stat_red = True Then
            Notif.lst_IGOD.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(20, "W").Text & ": " & Cells(c.Row, "W").Text
        ElseIf stat_red = True Then
            Notif.lst_IGOD.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(c.Row, "W").Text
        ElseIf stat_red = False Then
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_IGOD.AddItem "Outdated:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If c.Row >= 9 And c.Row <= 17 And stat_black = True Then
            Notif.lst_IGOD.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(8, "W").Text & ": " & Cells(c.Row, "W").Text
        ElseIf c.Row >= 21 And c.Row <= 43 And stat_black = True Then
            Notif.lst_IGOD.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(20, "W").Text & ": " & Cells(c.Row, "W").Text
        ElseIf stat_black = True Then
            Notif.lst_IGOD.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(c.Row, "W").Text
        ElseIf stat_black = False Then
        End If
    Next
    
    'WGMMFxATOp
    Sheet6.Activate
    Set rng = Sheet6.Range("G3:G7,G9:G45,G47:G70")
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_WGMMFxATOp.AddItem "3 Months Before Deadline:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If c.Row >= 9 And c.Row <= 43 And stat_orange = True Then
            Notif.lst_WGMMFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(8, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 47 And c.Row <= 69 And stat_orange = True Then
            Notif.lst_WGMMFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(46, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf stat_orange = True Then
            Notif.lst_WGMMFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_orange = False Then
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_WGMMFxATOp.AddItem "For Urgent Update:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If c.Row >= 9 And c.Row <= 43 And stat_red = True Then
            Notif.lst_WGMMFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(8, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 47 And c.Row <= 69 And stat_red = True Then
            Notif.lst_WGMMFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(46, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf stat_red = True Then
            Notif.lst_WGMMFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_red = False Then
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_WGMMFxATOp.AddItem "Outdated:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If c.Row >= 9 And c.Row <= 43 And stat_black = True Then
            Notif.lst_WGMMFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(8, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 47 And c.Row <= 69 And stat_black = True Then
            Notif.lst_WGMMFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(46, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf stat_black = True Then
            Notif.lst_WGMMFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_black = False Then
        End If
    Next
    
    'FCNO FF2 SPC
    Set rng = Sheets("PLDT Sampaloc").Range("Q3:Q7,Q9:Q32,Q34:Q57")
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_FCNOS.AddItem "3 Months Before Deadline:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If c.Row >= 9 And c.Row <= 30 And stat_orange = True Then
            Notif.lst_FCNOS.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(8, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf c.Row >= 34 And c.Row <= 56 And stat_orange = True Then
            Notif.lst_FCNOS.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(33, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf stat_orange = True Then
            Notif.lst_FCNOS.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(c.Row, "M").Text
        ElseIf stat_orange = False Then
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_FCNOS.AddItem "For Urgent Update:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If c.Row >= 9 And c.Row <= 30 And stat_red = True Then
            Notif.lst_FCNOS.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(8, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf c.Row >= 34 And c.Row <= 56 And stat_red = True Then
            Notif.lst_FCNOS.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(33, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf stat_red = True Then
            Notif.lst_FCNOS.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(c.Row, "M").Text
        ElseIf stat_red = False Then
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_FCNOS.AddItem "Outdated:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If c.Row >= 9 And c.Row <= 30 And stat_black = True Then
            Notif.lst_FCNOS.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(8, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf c.Row >= 34 And c.Row <= 56 And stat_black = True Then
            Notif.lst_FCNOS.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(33, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf stat_black = True Then
            Notif.lst_FCNOS.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(c.Row, "M").Text
        ElseIf stat_black = False Then
        End If
    Next
    
    'Manila IGO
    Set rng = Sheet6.Range("AA3:AA7,AA9:AA25,AA27:AA50")
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_IGOS.AddItem "3 Months Before Deadline:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If c.Row >= 9 And c.Row <= 23 And stat_orange = True Then
            Notif.lst_IGOS.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(8, "W").Text & ": " & Cells(c.Row, "W").Text
        ElseIf c.Row >= 27 And c.Row <= 49 And stat_orange = True Then
            Notif.lst_IGOS.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(26, "W").Text & ": " & Cells(c.Row, "W").Text
        ElseIf stat_orange = True Then
            Notif.lst_IGOS.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(c.Row, "W").Text
        ElseIf stat_orange = False Then
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_IGOS.AddItem "For Urgent Update:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If c.Row >= 9 And c.Row <= 23 And stat_red = True Then
            Notif.lst_IGOS.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(8, "W").Text & ": " & Cells(c.Row, "W").Text
        ElseIf c.Row >= 27 And c.Row <= 49 And stat_red = True Then
            Notif.lst_IGOS.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(26, "W").Text & ": " & Cells(c.Row, "W").Text
        ElseIf stat_red = True Then
            Notif.lst_IGOS.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(c.Row, "W").Text
        ElseIf stat_red = False Then
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_IGOS.AddItem "Outdated:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If c.Row >= 9 And c.Row <= 23 And stat_black = True Then
            Notif.lst_IGOS.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(8, "W").Text & ": " & Cells(c.Row, "W").Text
        ElseIf c.Row >= 27 And c.Row <= 49 And stat_black = True Then
            Notif.lst_IGOS.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(26, "W").Text & ": " & Cells(c.Row, "W").Text
        ElseIf stat_black = True Then
            Notif.lst_IGOS.AddItem "    " & Cells(c.Row, "Y").Value & "    " & Cells(c.Row, "Z").Value & "    " & "REV " & Cells(c.Row, "X").Text & "    " & Cells(c.Row, "W").Text
        ElseIf stat_black = False Then
        End If
    Next
    
    'EGMMFxATOp GHL
    Sheet5.Activate
    Set rng = Sheet5.Range("G3:G7,G9:G34,G36:G59")
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_EGMMGHL.AddItem "3 Months Before Deadline:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If c.Row >= 9 And c.Row <= 32 And stat_orange = True Then
            Notif.lst_EGMMGHL.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(8, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 36 And c.Row <= 58 And stat_orange = True Then
            Notif.lst_EGMMGHL.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(35, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf stat_orange = True Then
            Notif.lst_EGMMGHL.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_orange = False Then
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_EGMMGHL.AddItem "For Urgent Update:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If c.Row >= 9 And c.Row <= 32 And stat_red = True Then
            Notif.lst_EGMMGHL.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(8, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 36 And c.Row <= 58 And stat_red = True Then
            Notif.lst_EGMMGHL.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(35, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf stat_red = True Then
            Notif.lst_EGMMGHL.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_red = False Then
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_EGMMGHL.AddItem "Outdated:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If c.Row >= 9 And c.Row <= 32 And stat_black = True Then
            Notif.lst_EGMMGHL.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(8, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 36 And c.Row <= 58 And stat_black = True Then
            Notif.lst_EGMMGHL.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(35, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf stat_black = True Then
            Notif.lst_EGMMGHL.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_black = False Then
        End If
    Next
    
    'FCNO GHL
    Set rng = Sheet5.Range("Q3:Q7,Q9:Q22,Q24:Q47")
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_FCNOGH.AddItem "3 Months Before Deadline:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If c.Row >= 9 And c.Row <= 20 And stat_orange = True Then
            Notif.lst_FCNOGH.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(8, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf c.Row >= 24 And c.Row <= 46 And stat_orange = True Then
            Notif.lst_FCNOGH.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(23, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf stat_orange = True Then
            Notif.lst_FCNOGH.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(c.Row, "M").Text
        ElseIf stat_orange = False Then
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_FCNOGH.AddItem "For Urgent Update:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If c.Row >= 9 And c.Row <= 20 And stat_red = True Then
            Notif.lst_FCNOGH.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(8, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf c.Row >= 24 And c.Row <= 46 And stat_red = True Then
            Notif.lst_FCNOGH.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(23, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf stat_red = True Then
            Notif.lst_FCNOGH.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(c.Row, "M").Text
        ElseIf stat_red = False Then
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_FCNOGH.AddItem "Outdated:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If c.Row >= 9 And c.Row <= 20 And stat_black = True Then
            Notif.lst_FCNOGH.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(8, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf c.Row >= 24 And c.Row <= 46 And stat_black = True Then
            Notif.lst_FCNOGH.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(23, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf stat_black = True Then
            Notif.lst_FCNOGH.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(c.Row, "M").Text
        ElseIf stat_black = False Then
        End If
    Next
    
    Call Init2
    Call Init3
End Sub

Sub Init2()
    Dim c, rng As Range
    Dim stat_black, stat_red, stat_orange As Boolean
    
    'EGMMFxATOp GNT
    Sheet11.Activate
    Set rng = Sheet11.Range("G3:G7,G9:G34,G36:G59")
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_EGMMGNT.AddItem "3 Months Before Deadline:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If c.Row >= 9 And c.Row <= 32 And stat_orange = True Then
            Notif.lst_EGMMGNT.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(8, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 36 And c.Row <= 58 And stat_orange = True Then
            Notif.lst_EGMMGNT.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(35, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf stat_orange = True Then
            Notif.lst_EGMMGNT.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_orange = False Then
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_EGMMGNT.AddItem "For Urgent Update:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If c.Row >= 9 And c.Row <= 32 And stat_red = True Then
            Notif.lst_EGMMGNT.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(8, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 36 And c.Row <= 58 And stat_red = True Then
            Notif.lst_EGMMGNT.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(35, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf stat_red = True Then
            Notif.lst_EGMMGNT.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_red = False Then
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_EGMMGNT.AddItem "Outdated:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If c.Row >= 9 And c.Row <= 32 And stat_black = True Then
            Notif.lst_EGMMGNT.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(8, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 36 And c.Row <= 58 And stat_black = True Then
            Notif.lst_EGMMGNT.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(35, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf stat_black = True Then
            Notif.lst_EGMMGNT.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_black = False Then
        End If
    Next
    
    'FCNO GHL
    Set rng = Sheet12.Range("Q3:Q7,Q9:Q22,Q24:Q47")
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_FCNOGNT.AddItem "3 Months Before Deadline:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If c.Row >= 9 And c.Row <= 20 And stat_orange = True Then
            Notif.lst_FCNOGNT.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(8, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf c.Row >= 24 And c.Row <= 46 And stat_orange = True Then
            Notif.lst_FCNOGNT.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(23, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf stat_orange = True Then
            Notif.lst_FCNOGNT.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(c.Row, "M").Text
        ElseIf stat_orange = False Then
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_FCNOGNT.AddItem "For Urgent Update:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If c.Row >= 9 And c.Row <= 20 And stat_red = True Then
            Notif.lst_FCNOGNT.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(8, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf c.Row >= 24 And c.Row <= 46 And stat_red = True Then
            Notif.lst_FCNOGNT.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(23, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf stat_red = True Then
            Notif.lst_FCNOGNT.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(c.Row, "M").Text
        ElseIf stat_red = False Then
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_FCNOGNT.AddItem "Outdated:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If c.Row >= 9 And c.Row <= 20 And stat_black = True Then
            Notif.lst_FCNOGNT.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(8, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf c.Row >= 24 And c.Row <= 46 And stat_black = True Then
            Notif.lst_FCNOGNT.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(23, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf stat_black = True Then
            Notif.lst_FCNOGNT.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(c.Row, "M").Text
        ElseIf stat_black = False Then
        End If
    Next
    
    'VisFxATOp
    Sheet12.Activate
    Set rng = Sheet12.Range("G3:G7,G9:G34,G36:G59")
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_VisFxATOp.AddItem "3 Months Before Deadline:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If c.Row >= 9 And c.Row <= 32 And stat_orange = True Then
            Notif.lst_VisFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(8, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 36 And c.Row <= 58 And stat_orange = True Then
            Notif.lst_VisFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(35, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf stat_orange = True Then
            Notif.lst_VisFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_orange = False Then
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_VisFxATOp.AddItem "For Urgent Update:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If c.Row >= 9 And c.Row <= 32 And stat_red = True Then
            Notif.lst_VisFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(8, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 36 And c.Row <= 58 And stat_red = True Then
            Notif.lst_VisFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(35, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf stat_red = True Then
            Notif.lst_VisFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_red = False Then
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_VisFxATOp.AddItem "Outdated:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If c.Row >= 9 And c.Row <= 32 And stat_black = True Then
            Notif.lst_VisFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(8, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 36 And c.Row <= 58 And stat_black = True Then
            Notif.lst_VisFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(35, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf stat_black = True Then
            Notif.lst_VisFxATOp.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_black = False Then
        End If
    Next
    
    'FCNO JNE
    Set rng = Sheet12.Range("Q3:Q7,Q9:Q22,Q24:Q47")
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_FCNOJNE.AddItem "3 Months Before Deadline:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If c.Row >= 9 And c.Row <= 20 And stat_orange = True Then
            Notif.lst_FCNOJNE.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(8, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf c.Row >= 24 And c.Row <= 46 And stat_orange = True Then
            Notif.lst_FCNOJNE.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(23, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf stat_orange = True Then
            Notif.lst_FCNOJNE.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(c.Row, "M").Text
        ElseIf stat_orange = False Then
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_FCNOJNE.AddItem "For Urgent Update:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If c.Row >= 9 And c.Row <= 20 And stat_red = True Then
            Notif.lst_FCNOJNE.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(8, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf c.Row >= 24 And c.Row <= 46 And stat_red = True Then
            Notif.lst_FCNOJNE.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(23, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf stat_red = True Then
            Notif.lst_FCNOJNE.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(c.Row, "M").Text
        ElseIf stat_red = False Then
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_FCNOJNE.AddItem "Outdated:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If c.Row >= 9 And c.Row <= 20 And stat_black = True Then
            Notif.lst_FCNOJNE.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(8, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf c.Row >= 24 And c.Row <= 46 And stat_black = True Then
            Notif.lst_FCNOJNE.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(23, "M").Text & ": " & Cells(c.Row, "M").Text
        ElseIf stat_black = True Then
            Notif.lst_FCNOJNE.AddItem "    " & Cells(c.Row, "O").Value & "    " & Cells(c.Row, "P").Value & "    " & "REV " & Cells(c.Row, "N").Text & "    " & Cells(c.Row, "M").Text
        ElseIf stat_black = False Then
        End If
    Next
    
    'CONSOLIDATED STRAT
    Sheet7.Activate
    Set rng = Sheet7.Range("G3:G4,G28:G29,G53:G54,G78:G79")
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_STRATC.AddItem "3 Months Before Deadline:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_STRATC.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_orange = False Then
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_STRATC.AddItem "For Urgent Update:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_STRATC.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_red = False Then
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_STRATC.AddItem "Outdated:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_STRATC.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(c.Row, "C").Text
        ElseIf stat_black = False Then
        End If
    Next
    
    'CONSOLIDATED ENG
    Set rng = Sheet7.Range("G6:G27,G31:G52,G56:G77")
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If stat_orange = True Then
            Notif.lst_ENGC.AddItem "3 Months Before Deadline:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_orange = Evaluate(c.FormatConditions.Item(4).Formula1)
        If c.Row >= 6 And c.Row <= 27 And stat_orange = True Then
            Notif.lst_ENGC.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(5, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 31 And c.Row <= 52 And stat_orange = True Then
            Notif.lst_ENGC.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(30, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 56 And c.Row <= 77 And stat_orange = True Then
            Notif.lst_ENGC.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(55, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf stat_orange = False Then
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If stat_red = True Then
            Notif.lst_ENGC.AddItem "For Urgent Update:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_red = Evaluate(c.FormatConditions.Item(3).Formula1)
        If c.Row >= 6 And c.Row <= 27 And stat_red = True Then
            Notif.lst_ENGC.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(8, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 31 And c.Row <= 52 And stat_red = True Then
            Notif.lst_ENGC.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(35, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 56 And c.Row <= 77 And stat_red = True Then
            Notif.lst_ENGC.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(55, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf stat_red = False Then
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If stat_black = True Then
            Notif.lst_ENGC.AddItem "Outdated:"
            Exit For
        End If
    Next
    For Each c In rng
        stat_black = Evaluate(c.FormatConditions.Item(2).Formula1)
        If c.Row >= 6 And c.Row <= 27 And stat_black = True Then
            Notif.lst_ENGC.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(8, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 31 And c.Row <= 52 And stat_black = True Then
            Notif.lst_ENGC.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(35, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf c.Row >= 56 And c.Row <= 77 And stat_black = True Then
            Notif.lst_ENGC.AddItem "    " & Cells(c.Row, "E").Value & "    " & Cells(c.Row, "F").Value & "    " & "REV " & Cells(c.Row, "D").Text & "    " & Cells(55, "C").Text & ": " & Cells(c.Row, "C").Text
        ElseIf stat_black = False Then
        End If
    Next
End Sub

Sub Init3()
    If Notif.lst_GOV.ListCount >= 2 Or Notif.lst_STRAT.ListCount >= 2 Or Notif.lst_ENG.ListCount >= 2 Then
        Notif.MultiPage1.Pages(0).Caption = "*" + Notif.MultiPage1.Pages(0).Caption
        If Notif.lst_GOV.ListCount >= 2 Then
            Notif.MultiPage2.Pages(0).Caption = "*" + Notif.MultiPage2.Pages(0).Caption
            Notif.Image1.Visible = True
        End If
        If Notif.lst_STRAT.ListCount >= 2 Then
            Notif.MultiPage2.Pages(1).Caption = "*" + Notif.MultiPage2.Pages(1).Caption
            Notif.Image13.Visible = True
        End If
        If Notif.lst_ENG.ListCount >= 2 Then
            Notif.MultiPage2.Pages(2).Caption = "*" + Notif.MultiPage2.Pages(2).Caption
            Notif.Image14.Visible = True
        End If
        
    End If
    If Notif.lst_LUCLS.ListCount >= 2 Or Notif.lst_BCLS.ListCount >= 2 Or Notif.lst_DCLS.ListCount >= 2 Then
        Notif.MultiPage1.Pages(1).Caption = "*" + Notif.MultiPage1.Pages(1).Caption
        If Notif.lst_LUCLS.ListCount >= 2 Then
            Notif.MultiPage3.Pages(0).Caption = "*" + Notif.MultiPage3.Pages(0).Caption
            Notif.Image2.Visible = True
        End If
        If Notif.lst_BCLS.ListCount >= 2 Then
            Notif.MultiPage3.Pages(1).Caption = "*" + Notif.MultiPage3.Pages(1).Caption
            Notif.Image3.Visible = True
        End If
        If Notif.lst_DCLS.ListCount >= 2 Then
            Notif.MultiPage3.Pages(2).Caption = "*" + Notif.MultiPage3.Pages(2).Caption
            Notif.Image4.Visible = True
        End If
    End If
    If Notif.lst_NGMMFxATOp.ListCount >= 2 Or Notif.lst_FCNOD.ListCount >= 2 Or Notif.lst_IGOD.ListCount >= 2 Then
        Notif.MultiPage1.Pages(2).Caption = "*" + Notif.MultiPage1.Pages(2).Caption
        If Notif.lst_NGMMFxATOp.ListCount >= 2 Then
            Notif.MultiPage4.Pages(0).Caption = "*" + Notif.MultiPage4.Pages(0).Caption
            Notif.Image5.Visible = True
        End If
        If Notif.lst_FCNOD.ListCount >= 2 Then
            Notif.MultiPage4.Pages(1).Caption = "*" + Notif.MultiPage4.Pages(1).Caption
            Notif.Image6.Visible = True
        End If
        If Notif.lst_IGOD.ListCount >= 2 Then
            Notif.MultiPage4.Pages(2).Caption = "*" + Notif.MultiPage4.Pages(2).Caption
            Notif.Image7.Visible = True
        End If
    End If
    If Notif.lst_WGMMFxATOp.ListCount >= 2 Or Notif.lst_FCNOS.ListCount >= 2 Or Notif.lst_IGOS.ListCount >= 2 Then
        Notif.MultiPage1.Pages(3).Caption = "*" + Notif.MultiPage1.Pages(3).Caption
        If Notif.lst_WGMMFxATOp.ListCount >= 2 Then
            Notif.MultiPage5.Pages(0).Caption = "*" + Notif.MultiPage5.Pages(0).Caption
            Notif.Image8.Visible = True
        End If
        If Notif.lst_FCNOS.ListCount >= 2 Then
            Notif.MultiPage5.Pages(1).Caption = "*" + Notif.MultiPage5.Pages(1).Caption
            Notif.Image9.Visible = True
        End If
        If Notif.lst_IGOS.ListCount >= 2 Then
            Notif.MultiPage5.Pages(2).Caption = "*" + Notif.MultiPage5.Pages(2).Caption
            Notif.Image10.Visible = True
        End If
    End If
    If Notif.lst_EGMMGHL.ListCount >= 2 Or Notif.lst_FCNOGH.ListCount >= 2 Then
        Notif.MultiPage1.Pages(4).Caption = "*" + Notif.MultiPage1.Pages(4).Caption
        If Notif.lst_EGMMGHL.ListCount >= 2 Then
            Notif.MultiPage6.Pages(0).Caption = "*" + Notif.MultiPage6.Pages(0).Caption
            Notif.Image11.Visible = True
        End If
        If Notif.lst_FCNOGH.ListCount >= 2 Then
            Notif.MultiPage6.Pages(1).Caption = "*" + Notif.MultiPage6.Pages(1).Caption
            Notif.Image12.Visible = True
        End If
    End If
    If Notif.lst_EGMMGNT.ListCount >= 2 Or Notif.lst_FCNOGNT.ListCount >= 2 Then
        Notif.MultiPage1.Pages(5).Caption = "*" + Notif.MultiPage1.Pages(5).Caption
        If Notif.lst_EGMMGNT.ListCount >= 2 Then
            Notif.MultiPage8.Pages(0).Caption = "*" + Notif.MultiPage8.Pages(0).Caption
            Notif.Image17.Visible = True
        End If
        If Notif.lst_FCNOGNT.ListCount >= 2 Then
            Notif.MultiPage8.Pages(1).Caption = "*" + Notif.MultiPage8.Pages(1).Caption
            Notif.Image18.Visible = True
        End If
    End If
    If Notif.lst_VisFxATOp.ListCount >= 2 Or Notif.lst_FCNOJNE.ListCount >= 2 Then
        Notif.MultiPage1.Pages(6).Caption = "*" + Notif.MultiPage1.Pages(6).Caption
        If Notif.lst_VisFxATOp.ListCount >= 2 Then
            Notif.MultiPage9.Pages(0).Caption = "*" + Notif.MultiPage9.Pages(0).Caption
            Notif.Image19.Visible = True
        End If
        If Notif.lst_FCNOJNE.ListCount >= 2 Then
            Notif.MultiPage9.Pages(1).Caption = "*" + Notif.MultiPage9.Pages(1).Caption
            Notif.Image20.Visible = True
        End If
    End If
    If Notif.lst_STRATC.ListCount >= 2 Or Notif.lst_ENGC.ListCount >= 2 Then
        Notif.MultiPage1.Pages(7).Caption = "*" + Notif.MultiPage1.Pages(7).Caption
        If Notif.lst_STRATC.ListCount >= 2 Then
            Notif.MultiPage7.Pages(0).Caption = "*" + Notif.MultiPage7.Pages(0).Caption
            Notif.Image15.Visible = True
        End If
        If Notif.lst_ENGC.ListCount >= 2 Then
            Notif.MultiPage7.Pages(1).Caption = "*" + Notif.MultiPage7.Pages(1).Caption
            Notif.Image16.Visible = True
        End If
    End If
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Notif.Show
End Sub
