Attribute VB_Name = "Module2"
'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-'
'    Title: BCMS Project Monitoring (All Sites)     '
'    Author: Engr. Kenneth Caro Karamihan           '
'    Company: PLDT Inc.                             '
'    Division: BC Governance and Reporting          '
'    Date: May 24, 2018                             '
'    Code version: 2.0                              '
'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-'

Sub ReadDataFromCloseFile()
    On Error GoTo ErrHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Dim src As Workbook
    Dim WBD As String
    Dim Col As String
    Dim ColSrc As String
    Dim ErrMsg As String
    Dim iTotalRows As Integer
    Dim iSheetNum As Integer
    ErrMsg = ""
    ColSrc = "O"
    
    WBD = Sheet1.Range("AH16").Text & "\"
        
    ' OPEN BCLS EXCEL WORKBOOK
    If Not Sheet1.Range("BCLS").Text = "" Then
        Set src = Workbooks.Open(WBD & Sheet1.Range("BCLS").Text, True, True)
        iSheetNum = 1
        iTotalRows = src.Worksheets(iSheetNum).Range("B1:B" & Cells(Rows.Count, "B").End(xlUp).Row).Rows.Count
        Col = "O"
        Call WriteData(src, iTotalRows, Col, iSheetNum, ColSrc)
        src.Close False
        Set src = Nothing
    Else
        ErrMsg = "BCLS BCMS Monitoring xlsm file not found!"
        GoTo ErrHandler
    End If
    
    ' OPEN DCLS EXCEL WORKBOOK
    If Not Sheet1.Range("DCLS").Text = "" Then
        Set src = Workbooks.Open(WBD & Sheet1.Range("DCLS").Text, True, True)
        iSheetNum = 1
        iTotalRows = src.Worksheets(iSheetNum).Range("B1:B" & Cells(Rows.Count, "B").End(xlUp).Row).Rows.Count
        Col = "P"
        Call WriteData(src, iTotalRows, Col, iSheetNum, ColSrc)
        src.Close False
        Set src = Nothing
    Else
        ErrMsg = "DCLS BCMS Monitoring xlsm file not found!"
        GoTo ErrHandler
    End If
    
    ' OPEN LUCLS EXCEL WORKBOOK
    If Not Sheet1.Range("LUCLS").Text = "" Then
        Set src = Workbooks.Open(WBD & Sheet1.Range("LUCLS").Text, True, True)
        iSheetNum = 1
        iTotalRows = src.Worksheets(iSheetNum).Range("B1:B" & Cells(Rows.Count, "B").End(xlUp).Row).Rows.Count
        Col = "Q"
        Call WriteData(src, iTotalRows, Col, iSheetNum, ColSrc)
        src.Close False
        Set src = Nothing
    Else
        ErrMsg = "LUCLS BCMS Monitoring xlsm file not found!"
        GoTo ErrHandler
    End If
    
    ' OPEN FCNO FF1 GNT EXCEL WORKBOOK
    If Not Sheet1.Range("FCNO_FF1").Text = "" Then
        Set src = Workbooks.Open(WBD & Sheet1.Range("FCNO_FF1").Text, True, True)
        iSheetNum = 1
        iTotalRows = src.Worksheets(iSheetNum).Range("B1:B" & Cells(Rows.Count, "B").End(xlUp).Row).Rows.Count
        Col = "R"
        Call WriteData(src, iTotalRows, Col, iSheetNum, ColSrc)
        src.Close False
        Set src = Nothing
    Else
        ErrMsg = "FCNO FF1 GNT BCMS Monitoring xlsm file not found!"
        GoTo ErrHandler
    End If
    
    ' OPEN FCNO FF2 QCY EXCEL WORKBOOK
    If Not Sheet1.Range("FCNO_FF2").Text = "" Then
        Set src = Workbooks.Open(WBD & Sheet1.Range("FCNO_FF2").Text, True, True)
        iSheetNum = 1
        iTotalRows = src.Worksheets(iSheetNum).Range("B1:B" & Cells(Rows.Count, "B").End(xlUp).Row).Rows.Count
        Col = "S"
        Call WriteData(src, iTotalRows, Col, iSheetNum, ColSrc)
        src.Close False
        Set src = Nothing
    Else
        ErrMsg = "FCNO FF2 BCMS Monitoring xlsm file not found!"
        GoTo ErrHandler
    End If
    
    ' OPEN FCNO FF2 SPC EXCEL WORKBOOK
    If Not Sheet1.Range("FCNO_FF2").Text = "" Then
        Set src = Workbooks.Open(WBD & Sheet1.Range("FCNO_FF2").Text, True, True)
        iSheetNum = 2
        iTotalRows = src.Worksheets(iSheetNum).Range("B1:B" & Cells(Rows.Count, "B").End(xlUp).Row).Rows.Count
        Col = "T"
        Call WriteData(src, iTotalRows, Col, iSheetNum, ColSrc)
        src.Close False
        Set src = Nothing
    Else
        ErrMsg = "FCNO FF2 BCMS Monitoring xlsm file not found!"
        GoTo ErrHandler
    End If
    
    ' OPEN FCNO FF2 GHL EXCEL WORKBOOK
    If Not Sheet1.Range("FCNO_FF2").Text = "" Then
        Set src = Workbooks.Open(WBD & Sheet1.Range("FCNO_FF2").Text, True, True)
        iSheetNum = 3
        iTotalRows = src.Worksheets(iSheetNum).Range("B1:B" & Cells(Rows.Count, "B").End(xlUp).Row).Rows.Count
        Col = "U"
        Call WriteData(src, iTotalRows, Col, iSheetNum, ColSrc)
        src.Close False
        Set src = Nothing
    Else
        ErrMsg = "FCNO FF2 BCMS Monitoring xlsm file not found!"
        GoTo ErrHandler
    End If
    
    ' OPEN FCNO FF5 JNE EXCEL WORKBOOK
    If Not Sheet1.Range("FCNO_FF5").Text = "" Then
        Set src = Workbooks.Open(WBD & Sheet1.Range("FCNO_FF5").Text, True, True)
        iSheetNum = 1
        iTotalRows = src.Worksheets(iSheetNum).Range("B1:B" & Cells(Rows.Count, "B").End(xlUp).Row).Rows.Count
        Col = "V"
        Call WriteData(src, iTotalRows, Col, iSheetNum, ColSrc)
        src.Close False
        Set src = Nothing
    Else
        ErrMsg = "FCNO FF5 Cebu BCMS Monitoring xlsm file not found!"
        GoTo ErrHandler
    End If
    
    ' OPEN IGO QC DFON EXCEL WORKBOOK
    If Not Sheet1.Range("IGO").Text = "" Then
        Set src = Workbooks.Open(WBD & Sheet1.Range("IGO").Text, True, True)
        iSheetNum = 1
        iTotalRows = src.Worksheets(iSheetNum).Range("B1:B" & Cells(Rows.Count, "B").End(xlUp).Row).Rows.Count
        Col = "W"
        Call WriteData(src, iTotalRows, Col, iSheetNum, ColSrc)
        src.Close False
        Set src = Nothing
    Else
        ErrMsg = "IGO BCMS Monitoring xlsm file not found!"
        GoTo ErrHandler
    End If
    
    ' OPEN Manila IGO EXCEL WORKBOOK
    If Not Sheet1.Range("IGO").Text = "" Then
        Set src = Workbooks.Open(WBD & Sheet1.Range("IGO").Text, True, True)
        iSheetNum = 2
        iTotalRows = src.Worksheets(iSheetNum).Range("B1:B" & Cells(Rows.Count, "B").End(xlUp).Row).Rows.Count
        Col = "X"
        ColSrc = "P"
        Call WriteData(src, iTotalRows, Col, iSheetNum, ColSrc)
        src.Close False
        Set src = Nothing
        ColSrc = "O"
    Else
        ErrMsg = "IGO BCMS Monitoring xlsm file not found!"
        GoTo ErrHandler
    End If
    
    ' OPEN NGMMFxATOp FF1 EXCEL WORKBOOK
    If Not Sheet1.Range("NGMMFxATOp_FF1").Text = "" Then
        Set src = Workbooks.Open(WBD & Sheet1.Range("NGMMFxATOp_FF1").Text, True, True)
        iSheetNum = 1
        iTotalRows = src.Worksheets(iSheetNum).Range("B1:B" & Cells(Rows.Count, "B").End(xlUp).Row).Rows.Count
        Col = "Y"
        Call WriteData(src, iTotalRows, Col, iSheetNum, ColSrc)
        src.Close False
        Set src = Nothing
    Else
        ErrMsg = "NGMMFxATOp FF1 BCMS Monitoring xlsm file not found!"
        GoTo ErrHandler
    End If
    
    ' OPEN WGMMFxATOp FF1 EXCEL WORKBOOK
    If Not Sheet1.Range("WGMMFxATOp_FF1").Text = "" Then
        Set src = Workbooks.Open(WBD & Sheet1.Range("WGMMFxATOp_FF1").Text, True, True)
        iSheetNum = 1
        iTotalRows = src.Worksheets(iSheetNum).Range("B1:B" & Cells(Rows.Count, "B").End(xlUp).Row).Rows.Count
        Col = "Z"
        Call WriteData(src, iTotalRows, Col, iSheetNum, ColSrc)
        src.Close False
        Set src = Nothing
    Else
        ErrMsg = "WGMMFxATOp FF1 BCMS Monitoring xlsm file not found!"
        GoTo ErrHandler
    End If
    
    ' OPEN EGMMFxATOp FF1 GHL EXCEL WORKBOOK
    If Not Sheet1.Range("EGMMFxATOp_FF1").Text = "" Then
        Set src = Workbooks.Open(WBD & Sheet1.Range("EGMMFxATOp_FF1").Text, True, True)
        iSheetNum = 1
        iTotalRows = src.Worksheets(iSheetNum).Range("B1:B" & Cells(Rows.Count, "B").End(xlUp).Row).Rows.Count
        Col = "AA"
        Call WriteData(src, iTotalRows, Col, iSheetNum, ColSrc)
        src.Close False
        Set src = Nothing
    Else
        ErrMsg = "EGMMFxATOp FF1 BCMS Monitoring xlsm file not found!"
        GoTo ErrHandler
    End If
    
    ' OPEN EGMMFxATOp FF1 GNT EXCEL WORKBOOK
    If Not Sheet1.Range("EGMMFxATOp_FF1").Text = "" Then
        Set src = Workbooks.Open(WBD & Sheet1.Range("EGMMFxATOp_FF1").Text, True, True)
        iSheetNum = 2
        iTotalRows = src.Worksheets(iSheetNum).Range("B1:B" & Cells(Rows.Count, "B").End(xlUp).Row).Rows.Count
        Col = "AB"
        ColSrc = "P"
        Call WriteData(src, iTotalRows, Col, iSheetNum, ColSrc)
        src.Close False
        Set src = Nothing
        ColSrc = "O"
    Else
        ErrMsg = "EGMMFxATOp FF1 BCMS Monitoring xlsm file not found!"
        GoTo ErrHandler
    End If
    
    ' OPEN VisFxATOp FF5 EXCEL WORKBOOK
    If Not Sheet1.Range("VisFxATOp_FF5").Text = "" Then
        Set src = Workbooks.Open(WBD & Sheet1.Range("VisFxATOp_FF5").Text, True, True)
        iSheetNum = 1
        iTotalRows = src.Worksheets(iSheetNum).Range("B1:B" & Cells(Rows.Count, "B").End(xlUp).Row).Rows.Count
        Col = "AC"
        Call WriteData(src, iTotalRows, Col, iSheetNum, ColSrc)
        src.Close False
        Set src = Nothing
    Else
        ErrMsg = "VisFxATOp FF5 BCMS Monitoring xlsm file not found!"
        GoTo ErrHandler
    End If
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    Call ErrorHandler(ErrMsg)
End Sub

Sub WriteData(src As Excel.Workbook, iTotalRows As Integer, Col As String, iSheetNum As Integer, ColSrc As String)
    Dim iCnt As Integer
    Dim jCnt As Integer
    jCnt = 2
    For iCnt = 2 To iTotalRows
        If Sheet1.Range(Col & jCnt).HasFormula() = True Then
            jCnt = jCnt + 1
            GoTo ContinueLoop
        ElseIf src.Worksheets(iSheetNum).Range("B" & iCnt).Value Like Sheet1.Range("B" & jCnt).Value Then
            Sheet1.Range(Col & jCnt).Value = src.Worksheets(iSheetNum).Range(ColSrc & iCnt).Value
            jCnt = jCnt + 1
        End If
ContinueLoop:
    Next iCnt
    
End Sub

Sub ErrorHandler(ErrMsg As String)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox (ErrMsg)
End Sub

