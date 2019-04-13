VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Notif 
   Caption         =   "Status (Document Monitoring)"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11265
   OleObjectBlob   =   "Notif.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Notif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+'
'    Title: BCMS Docs Monitoring Automatic Email and Deadline Notifier     '
'    Author: Engr. Kenneth Caro Karamihan                                  '
'    Company: PLDT Inc.                                                    '
'    Division: BC Governance and Reporting                                 '
'    Date: May 22, 2018                                                    '
'    Code version: 1.0                                                     '
'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+'

Private Sub btn_Send_Click()
    Dim strbody As String
    Dim strtable As String
    Dim strto As String
    Dim strcc As String
    Dim s As String
    Dim i As Integer
    Dim signature As String
    
    If Sheet8.opt_ToGOV.Value = True Then
        strto = strto + Range("rdvillasenor").Value + "; "
    ElseIf Sheet8.opt_CCGOV.Value = True Then
        strcc = strcc + Range("rdvillasenor").Value + "; "
    ElseIf Sheet8.opt_NGOV.Value = True Then
    End If
          
    'GPs
    If Notif.MultiPage1.Value = 0 Then
    'GOV
        If Notif.MultiPage2.Value = 0 And Notif.MultiPage2.Pages(0).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B13").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            If Sheet8.opt_ToGOV2.Value = True Or Sheet8.opt_CCGOV2.Value = True Then
                strto = strto + Range("hclim").Value + "; "
            ElseIf Sheet8.opt_NGOV2.Value = True Then
            End If
            For i = 1 To Notif.lst_GOV.ListCount - 1
                s = Notif.lst_GOV.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B13").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'STRAT
        ElseIf Notif.MultiPage2.Value = 1 And Notif.MultiPage2.Pages(1).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B14").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToSTRAT.Value = True Then
                strto = strto + Range("aiseeco").Value + "; "
            ElseIf Sheet8.opt_CCSTRAT.Value = True Then
                strcc = strcc + Range("aiseeco").Value + "; "
            ElseIf Sheet8.opt_NSTRAT.Value = True Then
            End If
            
            If Sheet8.opt_ToSTRAT1.Value = True Then
                strto = strto + Range("bccordova").Value + "; "
            ElseIf Sheet8.opt_CCSTRAT1.Value = True Then
                strcc = strcc + Range("bccordova").Value + "; "
            ElseIf Sheet8.opt_NSTRAT1.Value = True Then
            End If
            
            If Sheet8.opt_ToSTRAT2.Value = True Then
                strto = strto + Range("jpjandayan").Value + "; "
            ElseIf Sheet8.opt_CCSTRAT2.Value = True Then
                strcc = strcc + Range("jpjandayan").Value + "; "
            ElseIf Sheet8.opt_NSTRAT2.Value = True Then
            End If
            
            For i = 1 To Notif.lst_STRAT.ListCount - 1
                s = Notif.lst_STRAT.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B14").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'ENG
        ElseIf Notif.MultiPage2.Value = 2 And Notif.MultiPage2.Pages(2).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B15").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToENG.Value = True Then
                strto = strto + Range("megutierrez").Value + "; "
            ElseIf Sheet8.opt_CCENG.Value = True Then
                strcc = strcc + Range("megutierrez").Value + "; "
            ElseIf Sheet8.opt_NENG.Value = True Then
            End If
            
            If Sheet8.opt_ToENG1.Value = True Then
                strto = strto + Range("amarsua").Value + "; "
            ElseIf Sheet8.opt_CCENG1.Value = True Then
                strcc = strcc + Range("amarsua").Value + "; "
            ElseIf Sheet8.opt_NENG1.Value = True Then
            End If
            
            If Sheet8.opt_ToENG2.Value = True Then
                strto = strto + Range("istolentino").Value + "; "
            ElseIf Sheet8.opt_CCENG2.Value = True Then
                strcc = strcc + Range("istolentino").Value + "; "
            ElseIf Sheet8.opt_NENG2.Value = True Then
            End If
            
            For i = 1 To Notif.lst_ENG.ListCount - 1
                s = Notif.lst_ENG.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B15").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
        End If
    End If
    
    'CLS
    If Notif.MultiPage1.Value = 1 And Notif.MultiPage1.Pages(1).Caption Like "*" Then
        If Sheet8.opt_ToCLS.Value = True Then
            strto = strto + Range("emgacayan").Value + "; "
        ElseIf Sheet8.opt_CCCLS.Value = True Then
            strcc = strcc + Range("emgacayan").Value + "; "
        ElseIf Sheet8.opt_NCLS.Value = True Then
        End If
    'LUCLS
        If Notif.MultiPage3.Value = 0 And Notif.MultiPage3.Pages(0).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B16").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToLUCLS1.Value = True Then
                strto = strto + Range("jajacinto").Value + "; "
            ElseIf Sheet8.opt_CCLUCLS1.Value = True Then
                strcc = strcc + Range("jajacinto").Value + "; "
            ElseIf Sheet8.opt_NLUCLS1.Value = True Then
            End If
            
            If Sheet8.opt_ToLUCLS2.Value = True Then
                strto = strto + Range("ljechave").Value + "; "
            ElseIf Sheet8.opt_CCLUCLS2.Value = True Then
                strcc = strcc + Range("ljechave").Value + "; "
            ElseIf Sheet8.opt_NLUCLS2.Value = True Then
            End If
            
            If Sheet8.opt_ToLUCLS3.Value = True Then
                strto = strto + Range("ecganuelas").Value + "; "
            ElseIf Sheet8.opt_CCLUCLS3.Value = True Then
                strcc = strcc + Range("ecganuelas").Value + "; "
            ElseIf Sheet8.opt_NLUCLS3.Value = True Then
            End If
            
            If Sheet8.opt_ToLUCLS4.Value = True Then
                strto = strto + Range("mmginete").Value + "; "
            ElseIf Sheet8.opt_CCLUCLS4.Value = True Then
                strcc = strcc + Range("mmginete").Value + "; "
            ElseIf Sheet8.opt_NLUCLS4.Value = True Then
            End If
            
            If Sheet8.opt_ToLUCLS5.Value = True Then
                strto = strto + Range("jogutierrez").Value + "; "
            ElseIf Sheet8.opt_CCLUCLS5.Value = True Then
                strcc = strcc + Range("jogutierrez").Value + "; "
            ElseIf Sheet8.opt_NLUCLS5.Value = True Then
            End If
            
            If Sheet8.opt_ToLUCLS6.Value = True Then
                strto = strto + Range("crcmendoza").Value + "; "
            ElseIf Sheet8.opt_CCLUCLS6.Value = True Then
                strcc = strcc + Range("crcmendoza").Value + "; "
            ElseIf Sheet8.opt_NLUCLS6.Value = True Then
            End If
            
            If Sheet8.opt_ToLUCLS7.Value = True Then
                strto = strto + Range("jlmoldez").Value + "; "
            ElseIf Sheet8.opt_CCLUCLS7.Value = True Then
                strcc = strcc + Range("jlmoldez").Value + "; "
            ElseIf Sheet8.opt_NLUCLS7.Value = True Then
            End If
            
            If Sheet8.opt_ToLUCLS8.Value = True Then
                strto = strto + Range("mmmones").Value + "; "
            ElseIf Sheet8.opt_CCLUCLS8.Value = True Then
                strcc = strcc + Range("mmmones").Value + "; "
            ElseIf Sheet8.opt_NLUCLS8.Value = True Then
            End If
            
            If Sheet8.opt_ToLUCLS9.Value = True Then
                strto = strto + Range("vprodriguez").Value + "; "
            ElseIf Sheet8.opt_CCLUCLS9.Value = True Then
                strcc = strcc + Range("vprodriguez").Value + "; "
            ElseIf Sheet8.opt_NLUCLS9.Value = True Then
            End If
                        
            For i = 1 To Notif.lst_LUCLS.ListCount - 1
                s = Notif.lst_LUCLS.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B16").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'BCLS
        ElseIf Notif.MultiPage3.Value = 1 And Notif.MultiPage3.Pages(1).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B17").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToBCLS1.Value = True Then
                strto = strto + Range("rjenriquez").Value + "; "
            ElseIf Sheet8.opt_CCBCLS1.Value = True Then
                strcc = strcc + Range("rjenriquez").Value + "; "
            ElseIf Sheet8.opt_NBCLS1.Value = True Then
            End If
            
            If Sheet8.opt_ToBCLS2.Value = True Then
                strto = strto + Range("jdbustamante").Value + "; "
            ElseIf Sheet8.opt_CCBCLS2.Value = True Then
                strcc = strcc + Range("jdbustamante").Value + "; "
            ElseIf Sheet8.opt_NBCLS2.Value = True Then
            End If
            
            If Sheet8.opt_ToBCLS3.Value = True Then
                strto = strto + Range("apcaringal").Value + "; "
            ElseIf Sheet8.opt_CCBCLS3.Value = True Then
                strcc = strcc + Range("apcaringal").Value + "; "
            ElseIf Sheet8.opt_NBCLS3.Value = True Then
            End If
            
            If Sheet8.opt_ToBCLS4.Value = True Then
                strto = strto + Range("jacatibog").Value + "; "
            ElseIf Sheet8.opt_CCBCLS4.Value = True Then
                strcc = strcc + Range("jacatibog").Value + "; "
            ElseIf Sheet8.opt_NBCLS4.Value = True Then
            End If
            
            If Sheet8.opt_ToBCLS5.Value = True Then
                strto = strto + Range("ebdeleon").Value + "; "
            ElseIf Sheet8.opt_CCBCLS5.Value = True Then
                strcc = strcc + Range("ebdeleon").Value + "; "
            ElseIf Sheet8.opt_NBCLS5.Value = True Then
            End If
            
            If Sheet8.opt_ToBCLS6.Value = True Then
                strto = strto + Range("pretcobanez").Value + "; "
            ElseIf Sheet8.opt_CCBCLS6.Value = True Then
                strcc = strcc + Range("pretcobanez").Value + "; "
            ElseIf Sheet8.opt_NBCLS6.Value = True Then
            End If
            
            If Sheet8.opt_ToBCLS7.Value = True Then
                strto = strto + Range("rpmanago").Value + "; "
            ElseIf Sheet8.opt_CCBCLS7.Value = True Then
                strcc = strcc + Range("rpmanago").Value + "; "
            ElseIf Sheet8.opt_NBCLS7.Value = True Then
            End If
            
            If Sheet8.opt_ToBCLS8.Value = True Then
                strto = strto + Range("hdpilar").Value + "; "
            ElseIf Sheet8.opt_CCBCLS8.Value = True Then
                strcc = strcc + Range("hdpilar").Value + "; "
            ElseIf Sheet8.opt_NBCLS8.Value = True Then
            End If
            
            If Sheet8.opt_ToBCLS9.Value = True Then
                strto = strto + Range("acreyes").Value + "; "
            ElseIf Sheet8.opt_CCBCLS9.Value = True Then
                strcc = strcc + Range("acreyes").Value + "; "
            ElseIf Sheet8.opt_NBCLS9.Value = True Then
            End If
            
            If Sheet8.opt_ToBCLS10.Value = True Then
                strto = strto + Range("aevizconde").Value + "; "
            ElseIf Sheet8.opt_CCBCLS10.Value = True Then
                strcc = strcc + Range("aevizconde").Value + "; "
            ElseIf Sheet8.opt_NBCLS10.Value = True Then
            End If
            
            For i = 1 To Notif.lst_BCLS.ListCount - 1
                s = Notif.lst_BCLS.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B17").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'DCLS
        ElseIf Notif.MultiPage3.Value = 2 And Notif.MultiPage3.Pages(2).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B18").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToDCLS1.Value = True Then
                strto = strto + Range("vvpunzalan").Value + "; "
            ElseIf Sheet8.opt_CCDCLS1.Value = True Then
                strcc = strcc + Range("vvpunzalan").Value + "; "
            ElseIf Sheet8.opt_NDCLS1.Value = True Then
            End If
            
            If Sheet8.opt_ToDCLS4.Value = True Then
                strto = strto + Range("rspanta").Value + "; "
            ElseIf Sheet8.opt_CCDCLS4.Value = True Then
                strcc = strcc + Range("rspanta").Value + "; "
            ElseIf Sheet8.opt_NDCLS4.Value = True Then
            End If
            
            If Sheet8.opt_ToDCLS2.Value = True Then
                strto = strto + Range("jlandicoy").Value + "; "
            ElseIf Sheet8.opt_CCDCLS2.Value = True Then
                strcc = strcc + Range("jlandicoy").Value + "; "
            ElseIf Sheet8.opt_NDCLS2.Value = True Then
            End If
            
            If Sheet8.opt_ToDCLS3.Value = True Then
                strto = strto + Range("mbdevilla").Value + "; "
            ElseIf Sheet8.opt_CCDCLS3.Value = True Then
                strcc = strcc + Range("mbdevilla").Value + "; "
            ElseIf Sheet8.opt_NDCLS3.Value = True Then
            End If
            
            If Sheet8.opt_ToDCLS5.Value = True Then
                strto = strto + Range("neromero").Value + "; "
            ElseIf Sheet8.opt_CCDCLS5.Value = True Then
                strcc = strcc + Range("neromero").Value + "; "
            ElseIf Sheet8.opt_NDCLS5.Value = True Then
            End If
            
            If Sheet8.opt_ToDCLS6.Value = True Then
                strto = strto + Range("josalvador").Value + "; "
            ElseIf Sheet8.opt_CCDCLS6.Value = True Then
                strcc = strcc + Range("josalvador").Value + "; "
            ElseIf Sheet8.opt_NDCLS6.Value = True Then
            End If
            
            If Sheet8.opt_ToDCLS7.Value = True Then
                strto = strto + Range("dltatad").Value + "; "
            ElseIf Sheet8.opt_CCDCLS7.Value = True Then
                strcc = strcc + Range("dltatad").Value + "; "
            ElseIf Sheet8.opt_NDCLS7.Value = True Then
            End If
            
            For i = 1 To Notif.lst_DCLS.ListCount - 1
                s = Notif.lst_DCLS.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B18").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
        End If
    End If
    
    'DILIMAN
    If Notif.MultiPage1.Value = 2 Then
    'NGMMFxATOp
        If Notif.MultiPage4.Value = 0 And Notif.MultiPage4.Pages(0).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B19").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToNGMM.Value = True Then
                strto = strto + Range("dcroque").Value + "; "
            ElseIf Sheet8.opt_CCNGMM.Value = True Then
                strcc = strcc + Range("dcroque").Value + "; "
            ElseIf Sheet8.opt_NNGMM.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM1.Value = True Then
                strto = strto + Range("agsoro").Value + "; "
            ElseIf Sheet8.opt_CCNGMM1.Value = True Then
                strcc = strcc + Range("agsoro").Value + "; "
            ElseIf Sheet8.opt_NNGMM1.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM2.Value = True Then
                strto = strto + Range("aoabary").Value + "; "
            ElseIf Sheet8.opt_CCNGMM2.Value = True Then
                strcc = strcc + Range("aoabary").Value + "; "
            ElseIf Sheet8.opt_NNGMM2.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM3.Value = True Then
                strto = strto + Range("rogbanez").Value + "; "
            ElseIf Sheet8.opt_CCNGMM3.Value = True Then
                strcc = strcc + Range("rogbanez").Value + "; "
            ElseIf Sheet8.opt_NNGMM3.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM4.Value = True Then
                strto = strto + Range("blbautista").Value + "; "
            ElseIf Sheet8.opt_CCNGMM4.Value = True Then
                strcc = strcc + Range("blbautista").Value + "; "
            ElseIf Sheet8.opt_NNGMM4.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM5.Value = True Then
                strto = strto + Range("ebbayle").Value + "; "
            ElseIf Sheet8.opt_CCNGMM5.Value = True Then
                strcc = strcc + Range("ebbayle").Value + "; "
            ElseIf Sheet8.opt_NNGMM5.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM6.Value = True Then
                strto = strto + Range("accruz").Value + "; "
            ElseIf Sheet8.opt_CCNGMM6.Value = True Then
                strcc = strcc + Range("accruz").Value + "; "
            ElseIf Sheet8.opt_NNGMM6.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM7.Value = True Then
                strto = strto + Range("mgdioso").Value + "; "
            ElseIf Sheet8.opt_CCNGMM7.Value = True Then
                strcc = strcc + Range("mgdioso").Value + "; "
            ElseIf Sheet8.opt_NNGMM7.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM8.Value = True Then
                strto = strto + Range("amestrella").Value + "; "
            ElseIf Sheet8.opt_CCNGMM8.Value = True Then
                strcc = strcc + Range("amestrella").Value + "; "
            ElseIf Sheet8.opt_NNGMM8.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM9.Value = True Then
                strto = strto + Range("drfonacier").Value + "; "
            ElseIf Sheet8.opt_CCNGMM9.Value = True Then
                strcc = strcc + Range("drfonacier").Value + "; "
            ElseIf Sheet8.opt_NNGMM9.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM10.Value = True Then
                strto = strto + Range("rjlim").Value + "; "
            ElseIf Sheet8.opt_CCNGMM10.Value = True Then
                strcc = strcc + Range("rjlim").Value + "; "
            ElseIf Sheet8.opt_NNGMM10.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM11.Value = True Then
                strto = strto + Range("ecmadrilejo").Value + "; "
            ElseIf Sheet8.opt_CCNGMM11.Value = True Then
                strcc = strcc + Range("ecmadrilejo").Value + "; "
            ElseIf Sheet8.opt_NNGMM11.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM12.Value = True Then
                strto = strto + Range("jvnaval").Value + "; "
            ElseIf Sheet8.opt_CCNGMM12.Value = True Then
                strcc = strcc + Range("jvnaval").Value + "; "
            ElseIf Sheet8.opt_NNGMM12.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM13.Value = True Then
                strto = strto + Range("wmsabile").Value + "; "
            ElseIf Sheet8.opt_CCNGMM13.Value = True Then
                strcc = strcc + Range("wmsabile").Value + "; "
            ElseIf Sheet8.opt_NNGMM13.Value = True Then
            End If
            
            For i = 1 To Notif.lst_NGMMFxATOp.ListCount - 1
                s = Notif.lst_NGMMFxATOp.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B19").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'FCNO FF2 QCY
        ElseIf Notif.MultiPage4.Value = 1 And Notif.MultiPage4.Pages(1).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B20").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToFCNO.Value = True Then
                strto = strto + Range("afcapiral").Value + "; "
            ElseIf Sheet8.opt_CCFCNO.Value = True Then
                strcc = strcc + Range("afcapiral").Value + "; "
            ElseIf Sheet8.opt_NFCNO.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO1.Value = True Then
                strto = strto + Range("picables").Value + "; "
            ElseIf Sheet8.opt_CCFCNO1.Value = True Then
                strcc = strcc + Range("picables").Value + "; "
            ElseIf Sheet8.opt_NFCNO1.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO2.Value = True Then
                strto = strto + Range("fcbaul").Value + "; "
            ElseIf Sheet8.opt_CCFCNO2.Value = True Then
                strcc = strcc + Range("fcbaul").Value + "; "
            ElseIf Sheet8.opt_NFCNO2.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO3.Value = True Then
                strto = strto + Range("aljimenez").Value + "; "
            ElseIf Sheet8.opt_CCFCNO3.Value = True Then
                strcc = strcc + Range("aljimenez").Value + "; "
            ElseIf Sheet8.opt_NFCNO3.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO4.Value = True Then
                strto = strto + Range("rtlampa").Value + "; "
            ElseIf Sheet8.opt_CCFCNO4.Value = True Then
                strcc = strcc + Range("rtlampa").Value + "; "
            ElseIf Sheet8.opt_NFCNO4.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO5.Value = True Then
                strto = strto + Range("oalinco").Value + "; "
            ElseIf Sheet8.opt_CCFCNO5.Value = True Then
                strcc = strcc + Range("oalinco").Value + "; "
            ElseIf Sheet8.opt_NFCNO5.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO6.Value = True Then
                strto = strto + Range("mfsantos").Value + "; "
            ElseIf Sheet8.opt_CCFCNO6.Value = True Then
                strcc = strcc + Range("mfsantos").Value + "; "
            ElseIf Sheet8.opt_NFCNO6.Value = True Then
            End If
            
            For i = 1 To Notif.lst_FCNOD.ListCount - 1
                s = Notif.lst_FCNOD.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B20").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'IGO QC DFON Station
        ElseIf Notif.MultiPage4.Value = 2 And Notif.MultiPage4.Pages(2).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B21").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToIGO.Value = True Then
                strto = strto + Range("asgaba").Value + "; "
            ElseIf Sheet8.opt_CCIGO.Value = True Then
                strcc = strcc + Range("asgaba").Value + "; "
            ElseIf Sheet8.opt_NIGO.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO1.Value = True Then
                strto = strto + Range("rdvizmanos").Value + "; "
            ElseIf Sheet8.opt_CCIGO1.Value = True Then
                strcc = strcc + Range("rdvizmanos").Value + "; "
            ElseIf Sheet8.opt_NIGO1.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO2.Value = True Then
                strto = strto + Range("csalejo").Value + "; "
            ElseIf Sheet8.opt_CCIGO2.Value = True Then
                strcc = strcc + Range("csalejo").Value + "; "
            ElseIf Sheet8.opt_NIGO2.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO3.Value = True Then
                strto = strto + Range("ecbuera").Value + "; "
            ElseIf Sheet8.opt_CCIGO3.Value = True Then
                strcc = strcc + Range("ecbuera").Value + "; "
            ElseIf Sheet8.opt_NIGO3.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO4.Value = True Then
                strto = strto + Range("kdcruz").Value + "; "
            ElseIf Sheet8.opt_CCIGO4.Value = True Then
                strcc = strcc + Range("kdcruz").Value + "; "
            ElseIf Sheet8.opt_NIGO4.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO5.Value = True Then
                strto = strto + Range("wdgdebelen").Value + "; "
            ElseIf Sheet8.opt_CCIGO5.Value = True Then
                strcc = strcc + Range("wdgdebelen").Value + "; "
            ElseIf Sheet8.opt_NIGO5.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO6.Value = True Then
                strto = strto + Range("nmestevez").Value + "; "
            ElseIf Sheet8.opt_CCIGO6.Value = True Then
                strcc = strcc + Range("nmestevez").Value + "; "
            ElseIf Sheet8.opt_NIGO6.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO7.Value = True Then
                strto = strto + Range("vdgrino").Value + "; "
            ElseIf Sheet8.opt_CCIGO7.Value = True Then
                strcc = strcc + Range("vdgrino").Value + "; "
            ElseIf Sheet8.opt_NIGO7.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO8.Value = True Then
                strto = strto + Range("jmhernandez").Value + "; "
            ElseIf Sheet8.opt_CCIGO8.Value = True Then
                strcc = strcc + Range("jmhernandez").Value + "; "
            ElseIf Sheet8.opt_NIGO8.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO9.Value = True Then
                strto = strto + Range("anmantes").Value + "; "
            ElseIf Sheet8.opt_CCIGO9.Value = True Then
                strcc = strcc + Range("anmantes").Value + "; "
            ElseIf Sheet8.opt_NIGO9.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO10.Value = True Then
                strto = strto + Range("lgpagay").Value + "; "
            ElseIf Sheet8.opt_CCIGO10.Value = True Then
                strcc = strcc + Range("lgpagay").Value + "; "
            ElseIf Sheet8.opt_NIGO10.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO11.Value = True Then
                strto = strto + Range("rjtorio").Value + "; "
            ElseIf Sheet8.opt_CCIGO11.Value = True Then
                strcc = strcc + Range("rjtorio").Value + "; "
            ElseIf Sheet8.opt_NIGO11.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO12.Value = True Then
                strto = strto + Range("vauypala").Value + "; "
            ElseIf Sheet8.opt_CCIGO12.Value = True Then
                strcc = strcc + Range("vauypala").Value + "; "
            ElseIf Sheet8.opt_NIGO12.Value = True Then
            End If
            
            For i = 1 To Notif.lst_IGOD.ListCount - 1
                s = Notif.lst_IGOD.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B21").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
        End If
    End If
    
    'SAMPALOC
    If Notif.MultiPage1.Value = 3 Then
    'WGMMFxATOp
        If Notif.MultiPage5.Value = 0 And Notif.MultiPage5.Pages(0).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B22").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToWGMM.Value = True Then
                strto = strto + Range("emalsol").Value + "; "
            ElseIf Sheet8.opt_CCWGMM.Value = True Then
                strcc = strcc + Range("emalsol").Value + "; "
            ElseIf Sheet8.opt_NWGMM.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM1.Value = True Then
                strto = strto + Range("nnelegado").Value + "; "
            ElseIf Sheet8.opt_CCWGMM1.Value = True Then
                strcc = strcc + Range("nnelegado").Value + "; "
            ElseIf Sheet8.opt_NWGMM1.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM2.Value = True Then
                strto = strto + Range("jnadia").Value + "; "
            ElseIf Sheet8.opt_CCWGMM2.Value = True Then
                strcc = strcc + Range("jnadia").Value + "; "
            ElseIf Sheet8.opt_NWGMM2.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM3.Value = True Then
                strto = strto + Range("aaagbayani").Value + "; "
            ElseIf Sheet8.opt_CCWGMM3.Value = True Then
                strcc = strcc + Range("aaagbayani").Value + "; "
            ElseIf Sheet8.opt_NWGMM3.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM4.Value = True Then
                strto = strto + Range("gpaquino").Value + "; "
            ElseIf Sheet8.opt_CCWGMM4.Value = True Then
                strcc = strcc + Range("gpaquino").Value + "; "
            ElseIf Sheet8.opt_NWGMM4.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM5.Value = True Then
                strto = strto + Range("rhatendido").Value + "; "
            ElseIf Sheet8.opt_CCWGMM5.Value = True Then
                strcc = strcc + Range("rhatendido").Value + "; "
            ElseIf Sheet8.opt_NWGMM5.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM6.Value = True Then
                strto = strto + Range("almgonzales").Value + "; "
            ElseIf Sheet8.opt_CCWGMM6.Value = True Then
                strcc = strcc + Range("almgonzales").Value + "; "
            ElseIf Sheet8.opt_NWGMM6.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM7.Value = True Then
                strto = strto + Range("rvhundana").Value + "; "
            ElseIf Sheet8.opt_CCWGMM7.Value = True Then
                strcc = strcc + Range("rvhundana").Value + "; "
            ElseIf Sheet8.opt_NWGMM7.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM8.Value = True Then
                strto = strto + Range("famacadaeg").Value + "; "
            ElseIf Sheet8.opt_CCWGMM8.Value = True Then
                strcc = strcc + Range("famacadaeg").Value + "; "
            ElseIf Sheet8.opt_NWGMM8.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM9.Value = True Then
                strto = strto + Range("rsmariano").Value + "; "
            ElseIf Sheet8.opt_CCWGMM9.Value = True Then
                strcc = strcc + Range("rsmariano").Value + "; "
            ElseIf Sheet8.opt_NWGMM9.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM10.Value = True Then
                strto = strto + Range("aenito").Value + "; "
            ElseIf Sheet8.opt_CCWGMM10.Value = True Then
                strcc = strcc + Range("aenito").Value + "; "
            ElseIf Sheet8.opt_NWGMM10.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM11.Value = True Then
                strto = strto + Range("papagtalunan").Value + "; "
            ElseIf Sheet8.opt_CCWGMM11.Value = True Then
                strcc = strcc + Range("papagtalunan").Value + "; "
            ElseIf Sheet8.opt_NWGMM11.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM12.Value = True Then
                strto = strto + Range("lpparadero").Value + "; "
            ElseIf Sheet8.opt_CCWGMM12.Value = True Then
                strcc = strcc + Range("lpparadero").Value + "; "
            ElseIf Sheet8.opt_NWGMM12.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM13.Value = True Then
                strto = strto + Range("basoriano").Value + "; "
            ElseIf Sheet8.opt_CCWGMM13.Value = True Then
                strcc = strcc + Range("basoriano").Value + "; "
            ElseIf Sheet8.opt_NWGMM13.Value = True Then
            End If
            
            For i = 1 To Notif.lst_WGMMFxATOp.ListCount - 1
                s = Notif.lst_WGMMFxATOp.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B22").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'FCNO FF2 SPC
        ElseIf Notif.MultiPage5.Value = 1 And Notif.MultiPage5.Pages(1).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B23").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToFCNO.Value = True Then
                strto = strto + Range("afcapiral").Value + "; "
            ElseIf Sheet8.opt_CCFCNO.Value = True Then
                strcc = strcc + Range("afcapiral").Value + "; "
            ElseIf Sheet8.opt_NFCNO.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO1.Value = True Then
                strto = strto + Range("picables").Value + "; "
            ElseIf Sheet8.opt_CCFCNO1.Value = True Then
                strcc = strcc + Range("picables").Value + "; "
            ElseIf Sheet8.opt_NFCNO1.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO2.Value = True Then
                strto = strto + Range("fcbaul").Value + "; "
            ElseIf Sheet8.opt_CCFCNO2.Value = True Then
                strcc = strcc + Range("fcbaul").Value + "; "
            ElseIf Sheet8.opt_NFCNO2.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO3.Value = True Then
                strto = strto + Range("aljimenez").Value + "; "
            ElseIf Sheet8.opt_CCFCNO3.Value = True Then
                strcc = strcc + Range("aljimenez").Value + "; "
            ElseIf Sheet8.opt_NFCNO3.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO4.Value = True Then
                strto = strto + Range("rtlampa").Value + "; "
            ElseIf Sheet8.opt_CCFCNO4.Value = True Then
                strcc = strcc + Range("rtlampa").Value + "; "
            ElseIf Sheet8.opt_NFCNO4.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO5.Value = True Then
                strto = strto + Range("oalinco").Value + "; "
            ElseIf Sheet8.opt_CCFCNO5.Value = True Then
                strcc = strcc + Range("oalinco").Value + "; "
            ElseIf Sheet8.opt_NFCNO5.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO6.Value = True Then
                strto = strto + Range("mfsantos").Value + "; "
            ElseIf Sheet8.opt_CCFCNO6.Value = True Then
                strcc = strcc + Range("mfsantos").Value + "; "
            ElseIf Sheet8.opt_NFCNO6.Value = True Then
            End If
            
            For i = 1 To Notif.lst_FCNOS.ListCount - 1
                s = Notif.lst_FCNOS.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B23").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'Manila IGO
        ElseIf Notif.MultiPage5.Value = 2 And Notif.MultiPage5.Pages(2).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B24").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToIGO.Value = True Then
                strto = strto + Range("asgaba").Value + "; "
            ElseIf Sheet8.opt_CCIGO.Value = True Then
                strcc = strcc + Range("asgaba").Value + "; "
            ElseIf Sheet8.opt_NIGO.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO1.Value = True Then
                strto = strto + Range("rdvizmanos").Value + "; "
            ElseIf Sheet8.opt_CCIGO1.Value = True Then
                strcc = strcc + Range("rdvizmanos").Value + "; "
            ElseIf Sheet8.opt_NIGO1.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO2.Value = True Then
                strto = strto + Range("csalejo").Value + "; "
            ElseIf Sheet8.opt_CCIGO2.Value = True Then
                strcc = strcc + Range("csalejo").Value + "; "
            ElseIf Sheet8.opt_NIGO2.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO3.Value = True Then
                strto = strto + Range("ecbuera").Value + "; "
            ElseIf Sheet8.opt_CCIGO3.Value = True Then
                strcc = strcc + Range("ecbuera").Value + "; "
            ElseIf Sheet8.opt_NIGO3.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO4.Value = True Then
                strto = strto + Range("kdcruz").Value + "; "
            ElseIf Sheet8.opt_CCIGO4.Value = True Then
                strcc = strcc + Range("kdcruz").Value + "; "
            ElseIf Sheet8.opt_NIGO4.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO5.Value = True Then
                strto = strto + Range("wdgdebelen").Value + "; "
            ElseIf Sheet8.opt_CCIGO5.Value = True Then
                strcc = strcc + Range("wdgdebelen").Value + "; "
            ElseIf Sheet8.opt_NIGO5.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO6.Value = True Then
                strto = strto + Range("nmestevez").Value + "; "
            ElseIf Sheet8.opt_CCIGO6.Value = True Then
                strcc = strcc + Range("nmestevez").Value + "; "
            ElseIf Sheet8.opt_NIGO6.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO7.Value = True Then
                strto = strto + Range("vdgrino").Value + "; "
            ElseIf Sheet8.opt_CCIGO7.Value = True Then
                strcc = strcc + Range("vdgrino").Value + "; "
            ElseIf Sheet8.opt_NIGO7.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO8.Value = True Then
                strto = strto + Range("jmhernandez").Value + "; "
            ElseIf Sheet8.opt_CCIGO8.Value = True Then
                strcc = strcc + Range("jmhernandez").Value + "; "
            ElseIf Sheet8.opt_NIGO8.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO9.Value = True Then
                strto = strto + Range("anmantes").Value + "; "
            ElseIf Sheet8.opt_CCIGO9.Value = True Then
                strcc = strcc + Range("anmantes").Value + "; "
            ElseIf Sheet8.opt_NIGO9.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO10.Value = True Then
                strto = strto + Range("lgpagay").Value + "; "
            ElseIf Sheet8.opt_CCIGO10.Value = True Then
                strcc = strcc + Range("lgpagay").Value + "; "
            ElseIf Sheet8.opt_NIGO10.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO11.Value = True Then
                strto = strto + Range("rjtorio").Value + "; "
            ElseIf Sheet8.opt_CCIGO11.Value = True Then
                strcc = strcc + Range("rjtorio").Value + "; "
            ElseIf Sheet8.opt_NIGO11.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO12.Value = True Then
                strto = strto + Range("vauypala").Value + "; "
            ElseIf Sheet8.opt_CCIGO12.Value = True Then
                strcc = strcc + Range("vauypala").Value + "; "
            ElseIf Sheet8.opt_NIGO12.Value = True Then
            End If
            
            For i = 1 To Notif.lst_IGOS.ListCount - 1
                s = Notif.lst_IGOS.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B24").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
        End If
    End If
    
    'GREENHILLS
    If Notif.MultiPage1.Value = 4 Then
    'EGMMFxATOp GHL
        If Notif.MultiPage6.Value = 0 And Notif.MultiPage6.Pages(0).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B25").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToEGMM.Value = True Then
                strto = strto + Range("ntalcantara").Value + "; "
            ElseIf Sheet8.opt_CCEGMM.Value = True Then
                strcc = strcc + Range("ntalcantara").Value + "; "
            ElseIf Sheet8.opt_NEGMM.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM1.Value = True Then
                strto = strto + Range("armadrelino").Value + "; "
            ElseIf Sheet8.opt_CCEGMM1.Value = True Then
                strcc = strcc + Range("armadrelino").Value + "; "
            ElseIf Sheet8.opt_NEGMM1.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM2.Value = True Then
                strto = strto + Range("lpabapo").Value + "; "
            ElseIf Sheet8.opt_CCEGMM2.Value = True Then
                strcc = strcc + Range("lpabapo").Value + "; "
            ElseIf Sheet8.opt_NEGMM2.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM3.Value = True Then
                strto = strto + Range("raaquino").Value + "; "
            ElseIf Sheet8.opt_CCEGMM3.Value = True Then
                strcc = strcc + Range("raaquino").Value + "; "
            ElseIf Sheet8.opt_NEGMM3.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM4.Value = True Then
                strto = strto + Range("rrcerdon").Value + "; "
            ElseIf Sheet8.opt_CCEGMM4.Value = True Then
                strcc = strcc + Range("rrcerdon").Value + "; "
            ElseIf Sheet8.opt_NEGMM4.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM5.Value = True Then
                strto = strto + Range("lmcoma").Value + "; "
            ElseIf Sheet8.opt_CCEGMM5.Value = True Then
                strcc = strcc + Range("lmcoma").Value + "; "
            ElseIf Sheet8.opt_NEGMM5.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM6.Value = True Then
                strto = strto + Range("radeguzman").Value + "; "
            ElseIf Sheet8.opt_CCEGMM6.Value = True Then
                strcc = strcc + Range("radeguzman").Value + "; "
            ElseIf Sheet8.opt_NEGMM6.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM7.Value = True Then
                strto = strto + Range("jugabriel").Value + "; "
            ElseIf Sheet8.opt_CCEGMM7.Value = True Then
                strcc = strcc + Range("jugabriel").Value + "; "
            ElseIf Sheet8.opt_NEGMM7.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM8.Value = True Then
                strto = strto + Range("mbgaspay").Value + "; "
            ElseIf Sheet8.opt_CCEGMM8.Value = True Then
                strcc = strcc + Range("mbgaspay").Value + "; "
            ElseIf Sheet8.opt_NEGMM8.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM9.Value = True Then
                strto = strto + Range("cgkatalbas").Value + "; "
            ElseIf Sheet8.opt_CCEGMM9.Value = True Then
                strcc = strcc + Range("cgkatalbas").Value + "; "
            ElseIf Sheet8.opt_NEGMM9.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM10.Value = True Then
                strto = strto + Range("nrlapus").Value + "; "
            ElseIf Sheet8.opt_CCEGMM10.Value = True Then
                strcc = strcc + Range("nrlapus").Value + "; "
            ElseIf Sheet8.opt_NEGMM10.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM11.Value = True Then
                strto = strto + Range("fslizardo").Value + "; "
            ElseIf Sheet8.opt_CCEGMM11.Value = True Then
                strcc = strcc + Range("fslizardo").Value + "; "
            ElseIf Sheet8.opt_NEGMM11.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM12.Value = True Then
                strto = strto + Range("drparrocha").Value + "; "
            ElseIf Sheet8.opt_CCEGMM12.Value = True Then
                strcc = strcc + Range("drparrocha").Value + "; "
            ElseIf Sheet8.opt_NEGMM12.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM13.Value = True Then
                strto = strto + Range("nsreas").Value + "; "
            ElseIf Sheet8.opt_CCEGMM13.Value = True Then
                strcc = strcc + Range("nsreas").Value + "; "
            ElseIf Sheet8.opt_NEGMM13.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM14.Value = True Then
                strto = strto + Range("jdsarinas").Value + "; "
            ElseIf Sheet8.opt_CCEGMM14.Value = True Then
                strcc = strcc + Range("jdsarinas").Value + "; "
            ElseIf Sheet8.opt_NEGMM14.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM15.Value = True Then
                strto = strto + Range("rjteves").Value + "; "
            ElseIf Sheet8.opt_CCEGMM15.Value = True Then
                strcc = strcc + Range("rjteves").Value + "; "
            ElseIf Sheet8.opt_NEGMM15.Value = True Then
            End If
            
            For i = 1 To Notif.lst_EGMMGHL.ListCount - 1
                s = Notif.lst_EGMMGHL.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B25").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'FCNO GHL
        ElseIf Notif.MultiPage6.Value = 1 And Notif.MultiPage6.Pages(1).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B26").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToFCNO.Value = True Then
                strto = strto + Range("afcapiral").Value + "; "
            ElseIf Sheet8.opt_CCFCNO.Value = True Then
                strcc = strcc + Range("afcapiral").Value + "; "
            ElseIf Sheet8.opt_NFCNO.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO1.Value = True Then
                strto = strto + Range("picables").Value + "; "
            ElseIf Sheet8.opt_CCFCNO1.Value = True Then
                strcc = strcc + Range("picables").Value + "; "
            ElseIf Sheet8.opt_NFCNO1.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO2.Value = True Then
                strto = strto + Range("fcbaul").Value + "; "
            ElseIf Sheet8.opt_CCFCNO2.Value = True Then
                strcc = strcc + Range("fcbaul").Value + "; "
            ElseIf Sheet8.opt_NFCNO2.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO3.Value = True Then
                strto = strto + Range("aljimenez").Value + "; "
            ElseIf Sheet8.opt_CCFCNO3.Value = True Then
                strcc = strcc + Range("aljimenez").Value + "; "
            ElseIf Sheet8.opt_NFCNO3.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO4.Value = True Then
                strto = strto + Range("rtlampa").Value + "; "
            ElseIf Sheet8.opt_CCFCNO4.Value = True Then
                strcc = strcc + Range("rtlampa").Value + "; "
            ElseIf Sheet8.opt_NFCNO4.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO5.Value = True Then
                strto = strto + Range("oalinco").Value + "; "
            ElseIf Sheet8.opt_CCFCNO5.Value = True Then
                strcc = strcc + Range("oalinco").Value + "; "
            ElseIf Sheet8.opt_NFCNO5.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO6.Value = True Then
                strto = strto + Range("mfsantos").Value + "; "
            ElseIf Sheet8.opt_CCFCNO6.Value = True Then
                strcc = strcc + Range("mfsantos").Value + "; "
            ElseIf Sheet8.opt_NFCNO6.Value = True Then
            End If
            
            For i = 1 To Notif.lst_FCNOGH.ListCount - 1
                s = Notif.lst_FCNOGH.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B26").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
        End If
    End If
    
    Call Send2(strbody, strtable, strto, strcc, s, signature)
    
End Sub

Sub Send2(strbody As String, strtable As String, strto As String, strcc As String, s As String, signature As String)
    Dim OutApp As Object
    Dim OutMail As Object
    
    'GARNET
    If Notif.MultiPage1.Value = 5 Then
    'EGMMFxATOp GNT
        If Notif.MultiPage8.Value = 0 And Notif.MultiPage8.Pages(0).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B27").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToEGMM.Value = True Then
                strto = strto + Range("ntalcantara").Value + "; "
            ElseIf Sheet8.opt_CCEGMM.Value = True Then
                strcc = strcc + Range("ntalcantara").Value + "; "
            ElseIf Sheet8.opt_NEGMM.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM1.Value = True Then
                strto = strto + Range("armadrelino").Value + "; "
            ElseIf Sheet8.opt_CCEGMM1.Value = True Then
                strcc = strcc + Range("armadrelino").Value + "; "
            ElseIf Sheet8.opt_NEGMM1.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM2.Value = True Then
                strto = strto + Range("lpabapo").Value + "; "
            ElseIf Sheet8.opt_CCEGMM2.Value = True Then
                strcc = strcc + Range("lpabapo").Value + "; "
            ElseIf Sheet8.opt_NEGMM2.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM3.Value = True Then
                strto = strto + Range("raaquino").Value + "; "
            ElseIf Sheet8.opt_CCEGMM3.Value = True Then
                strcc = strcc + Range("raaquino").Value + "; "
            ElseIf Sheet8.opt_NEGMM3.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM4.Value = True Then
                strto = strto + Range("rrcerdon").Value + "; "
            ElseIf Sheet8.opt_CCEGMM4.Value = True Then
                strcc = strcc + Range("rrcerdon").Value + "; "
            ElseIf Sheet8.opt_NEGMM4.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM5.Value = True Then
                strto = strto + Range("lmcoma").Value + "; "
            ElseIf Sheet8.opt_CCEGMM5.Value = True Then
                strcc = strcc + Range("lmcoma").Value + "; "
            ElseIf Sheet8.opt_NEGMM5.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM6.Value = True Then
                strto = strto + Range("radeguzman").Value + "; "
            ElseIf Sheet8.opt_CCEGMM6.Value = True Then
                strcc = strcc + Range("radeguzman").Value + "; "
            ElseIf Sheet8.opt_NEGMM6.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM7.Value = True Then
                strto = strto + Range("jugabriel").Value + "; "
            ElseIf Sheet8.opt_CCEGMM7.Value = True Then
                strcc = strcc + Range("jugabriel").Value + "; "
            ElseIf Sheet8.opt_NEGMM7.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM8.Value = True Then
                strto = strto + Range("mbgaspay").Value + "; "
            ElseIf Sheet8.opt_CCEGMM8.Value = True Then
                strcc = strcc + Range("mbgaspay").Value + "; "
            ElseIf Sheet8.opt_NEGMM8.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM9.Value = True Then
                strto = strto + Range("cgkatalbas").Value + "; "
            ElseIf Sheet8.opt_CCEGMM9.Value = True Then
                strcc = strcc + Range("cgkatalbas").Value + "; "
            ElseIf Sheet8.opt_NEGMM9.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM10.Value = True Then
                strto = strto + Range("nrlapus").Value + "; "
            ElseIf Sheet8.opt_CCEGMM10.Value = True Then
                strcc = strcc + Range("nrlapus").Value + "; "
            ElseIf Sheet8.opt_NEGMM10.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM11.Value = True Then
                strto = strto + Range("fslizardo").Value + "; "
            ElseIf Sheet8.opt_CCEGMM11.Value = True Then
                strcc = strcc + Range("fslizardo").Value + "; "
            ElseIf Sheet8.opt_NEGMM11.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM12.Value = True Then
                strto = strto + Range("drparrocha").Value + "; "
            ElseIf Sheet8.opt_CCEGMM12.Value = True Then
                strcc = strcc + Range("drparrocha").Value + "; "
            ElseIf Sheet8.opt_NEGMM12.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM13.Value = True Then
                strto = strto + Range("nsreas").Value + "; "
            ElseIf Sheet8.opt_CCEGMM13.Value = True Then
                strcc = strcc + Range("nsreas").Value + "; "
            ElseIf Sheet8.opt_NEGMM13.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM14.Value = True Then
                strto = strto + Range("jdsarinas").Value + "; "
            ElseIf Sheet8.opt_CCEGMM14.Value = True Then
                strcc = strcc + Range("jdsarinas").Value + "; "
            ElseIf Sheet8.opt_NEGMM14.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM15.Value = True Then
                strto = strto + Range("rjteves").Value + "; "
            ElseIf Sheet8.opt_CCEGMM15.Value = True Then
                strcc = strcc + Range("rjteves").Value + "; "
            ElseIf Sheet8.opt_NEGMM15.Value = True Then
            End If
            
            
            For i = 1 To Notif.lst_EGMMGNT.ListCount - 1
                s = Notif.lst_EGMMGNT.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B27").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'FCNO FF1 GNT
        ElseIf Notif.MultiPage8.Value = 1 And Notif.MultiPage8.Pages(1).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B28").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToFCNO.Value = True Then
                strto = strto + Range("afcapiral").Value + "; "
            ElseIf Sheet8.opt_CCFCNO.Value = True Then
                strcc = strcc + Range("afcapiral").Value + "; "
            ElseIf Sheet8.opt_NFCNO.Value = True Then
            End If
            
            If Sheet8.opt_ToGARNET1.Value = True Then
                strto = strto + Range("eanieva").Value + "; "
            ElseIf Sheet8.opt_CCGARNET1.Value = True Then
                strcc = strcc + Range("eanieva").Value + "; "
            ElseIf Sheet8.opt_NGARNET1.Value = True Then
            End If
            
            If Sheet8.opt_ToGARNET2.Value = True Then
                strto = strto + Range("aeenano").Value + "; "
            ElseIf Sheet8.opt_CCGARNET2.Value = True Then
                strcc = strcc + Range("aeenano").Value + "; "
            ElseIf Sheet8.opt_NGARNET2.Value = True Then
            End If
            
            If Sheet8.opt_ToGARNET3.Value = True Then
                strto = strto + Range("apinciong").Value + "; "
            ElseIf Sheet8.opt_CCGARNET3.Value = True Then
                strcc = strcc + Range("apinciong").Value + "; "
            ElseIf Sheet8.opt_NGARNET3.Value = True Then
            End If
            
            If Sheet8.opt_ToGARNET4.Value = True Then
                strto = strto + Range("rcroque").Value + "; "
            ElseIf Sheet8.opt_CCGARNET4.Value = True Then
                strcc = strcc + Range("rcroque").Value + "; "
            ElseIf Sheet8.opt_NGARNET4.Value = True Then
            End If
            
            If Sheet8.opt_ToGARNET5.Value = True Then
                strto = strto + Range("mosena").Value + "; "
            ElseIf Sheet8.opt_CCGARNET5.Value = True Then
                strcc = strcc + Range("mosena").Value + "; "
            ElseIf Sheet8.opt_NGARNET5.Value = True Then
            End If
            
            If Sheet8.opt_ToGARNET6.Value = True Then
                strto = strto + Range("jutamondong").Value + "; "
            ElseIf Sheet8.opt_CCGARNET6.Value = True Then
                strcc = strcc + Range("jutamondong").Value + "; "
            ElseIf Sheet8.opt_NGARNET6.Value = True Then
            End If
            
            For i = 1 To Notif.lst_FCNOGNT.ListCount - 1
                s = Notif.lst_FCNOGNT.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B28").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
        End If
    End If
    
    'CEBU
    If Notif.MultiPage1.Value = 6 Then
    'VisFxATOp
        If Notif.MultiPage9.Value = 0 And Notif.MultiPage9.Pages(0).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B29").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToVis.Value = True Then
                strto = strto + Range("rasalas").Value + "; "
            ElseIf Sheet8.opt_CCVis.Value = True Then
                strcc = strcc + Range("rasalas").Value + "; "
            ElseIf Sheet8.opt_NVis.Value = True Then
            End If
            
            If Sheet8.opt_ToVis1.Value = True Then
                strto = strto + Range("mpoplas").Value + "; "
            ElseIf Sheet8.opt_CCVis1.Value = True Then
                strcc = strcc + Range("mpoplas").Value + "; "
            ElseIf Sheet8.opt_NVis1.Value = True Then
            End If
            
            If Sheet8.opt_ToVis2.Value = True Then
                strto = strto + Range("emallosada").Value + "; "
            ElseIf Sheet8.opt_CCVis2.Value = True Then
                strcc = strcc + Range("emallosada").Value + "; "
            ElseIf Sheet8.opt_NVis2.Value = True Then
            End If
            
            If Sheet8.opt_ToVis3.Value = True Then
                strto = strto + Range("rbarias").Value + "; "
            ElseIf Sheet8.opt_CCVis3.Value = True Then
                strcc = strcc + Range("rbarias").Value + "; "
            ElseIf Sheet8.opt_NVis3.Value = True Then
            End If
            
            If Sheet8.opt_ToVis4.Value = True Then
                strto = strto + Range("aabaguio").Value + "; "
            ElseIf Sheet8.opt_CCVis4.Value = True Then
                strcc = strcc + Range("aabaguio").Value + "; "
            ElseIf Sheet8.opt_NVis4.Value = True Then
            End If
            
            If Sheet8.opt_ToVis5.Value = True Then
                strto = strto + Range("lacabigas").Value + "; "
            ElseIf Sheet8.opt_CCVis5.Value = True Then
                strcc = strcc + Range("lacabigas").Value + "; "
            ElseIf Sheet8.opt_NVis5.Value = True Then
            End If
            
            If Sheet8.opt_ToVis6.Value = True Then
                strto = strto + Range("jjcarumba").Value + "; "
            ElseIf Sheet8.opt_CCVis6.Value = True Then
                strcc = strcc + Range("jjcarumba").Value + "; "
            ElseIf Sheet8.opt_NVis6.Value = True Then
            End If
            
            If Sheet8.opt_ToVis7.Value = True Then
                strto = strto + Range("rgconchas").Value + "; "
            ElseIf Sheet8.opt_CCVis7.Value = True Then
                strcc = strcc + Range("rgconchas").Value + "; "
            ElseIf Sheet8.opt_NVis7.Value = True Then
            End If
            
            If Sheet8.opt_ToVis8.Value = True Then
                strto = strto + Range("vbcuevas").Value + "; "
            ElseIf Sheet8.opt_CCVis8.Value = True Then
                strcc = strcc + Range("vbcuevas").Value + "; "
            ElseIf Sheet8.opt_NVis8.Value = True Then
            End If
            
            If Sheet8.opt_ToVis9.Value = True Then
                strto = strto + Range("jbdesamparado").Value + "; "
            ElseIf Sheet8.opt_CCVis9.Value = True Then
                strcc = strcc + Range("jbdesamparado").Value + "; "
            ElseIf Sheet8.opt_NVis9.Value = True Then
            End If
            
            If Sheet8.opt_ToVis10.Value = True Then
                strto = strto + Range("rddesquitado").Value + "; "
            ElseIf Sheet8.opt_CCVis10.Value = True Then
                strcc = strcc + Range("rddesquitado").Value + "; "
            ElseIf Sheet8.opt_NVis10.Value = True Then
            End If
            
            If Sheet8.opt_ToVis11.Value = True Then
                strto = strto + Range("mmdevero").Value + "; "
            ElseIf Sheet8.opt_CCVis11.Value = True Then
                strcc = strcc + Range("mmdevero").Value + "; "
            ElseIf Sheet8.opt_NVis11.Value = True Then
            End If
            
            If Sheet8.opt_ToVis12.Value = True Then
                strto = strto + Range("jcfelisarta").Value + "; "
            ElseIf Sheet8.opt_CCVis12.Value = True Then
                strcc = strcc + Range("jcfelisarta").Value + "; "
            ElseIf Sheet8.opt_NVis12.Value = True Then
            End If
            
            If Sheet8.opt_ToVis13.Value = True Then
                strto = strto + Range("rdflores").Value + "; "
            ElseIf Sheet8.opt_CCVis13.Value = True Then
                strcc = strcc + Range("rdflores").Value + "; "
            ElseIf Sheet8.opt_NVis13.Value = True Then
            End If
            
            If Sheet8.opt_ToVis14.Value = True Then
                strto = strto + Range("gpintes").Value + "; "
            ElseIf Sheet8.opt_CCVis14.Value = True Then
                strcc = strcc + Range("gpintes").Value + "; "
            ElseIf Sheet8.opt_NVis14.Value = True Then
            End If
            
            If Sheet8.opt_ToVis15.Value = True Then
                strto = strto + Range("dsisleta").Value + "; "
            ElseIf Sheet8.opt_CCVis15.Value = True Then
                strcc = strcc + Range("dsisleta").Value + "; "
            ElseIf Sheet8.opt_NVis15.Value = True Then
            End If
            
            If Sheet8.opt_ToVis16.Value = True Then
                strto = strto + Range("rmlocaylocay").Value + "; "
            ElseIf Sheet8.opt_CCVis16.Value = True Then
                strcc = strcc + Range("rmlocaylocay").Value + "; "
            ElseIf Sheet8.opt_NVis16.Value = True Then
            End If
            
            If Sheet8.opt_ToVis17.Value = True Then
                strto = strto + Range("mmnadal").Value + "; "
            ElseIf Sheet8.opt_CCVis17.Value = True Then
                strcc = strcc + Range("mmnadal").Value + "; "
            ElseIf Sheet8.opt_NVis17.Value = True Then
            End If
            
            If Sheet8.opt_ToVis18.Value = True Then
                strto = strto + Range("npompad").Value + "; "
            ElseIf Sheet8.opt_CCVis18.Value = True Then
                strcc = strcc + Range("npompad").Value + "; "
            ElseIf Sheet8.opt_NVis18.Value = True Then
            End If
            
            If Sheet8.opt_ToVis19.Value = True Then
                strto = strto + Range("dbpepito").Value + "; "
            ElseIf Sheet8.opt_CCVis19.Value = True Then
                strcc = strcc + Range("dbpepito").Value + "; "
            ElseIf Sheet8.opt_NVis19.Value = True Then
            End If
            
            If Sheet8.opt_ToVis20.Value = True Then
                strto = strto + Range("izpono").Value + "; "
            ElseIf Sheet8.opt_CCVis20.Value = True Then
                strcc = strcc + Range("izpono").Value + "; "
            ElseIf Sheet8.opt_NVis20.Value = True Then
            End If
            
            If Sheet8.opt_ToVis21.Value = True Then
                strto = strto + Range("clrosales").Value + "; "
            ElseIf Sheet8.opt_CCVis21.Value = True Then
                strcc = strcc + Range("clrosales").Value + "; "
            ElseIf Sheet8.opt_NVis21.Value = True Then
            End If
            
            If Sheet8.opt_ToVis22.Value = True Then
                strto = strto + Range("mdsarcauga").Value + "; "
            ElseIf Sheet8.opt_CCVis22.Value = True Then
                strcc = strcc + Range("mdsarcauga").Value + "; "
            ElseIf Sheet8.opt_NVis22.Value = True Then
            End If
            
            If Sheet8.opt_ToVis23.Value = True Then
                strto = strto + Range("mlsarmiento").Value + "; "
            ElseIf Sheet8.opt_CCVis23.Value = True Then
                strcc = strcc + Range("mlsarmiento").Value + "; "
            ElseIf Sheet8.opt_NVis23.Value = True Then
            End If
            
            If Sheet8.opt_ToVis24.Value = True Then
                strto = strto + Range("vjson").Value + "; "
            ElseIf Sheet8.opt_CCVis24.Value = True Then
                strcc = strcc + Range("vjson").Value + "; "
            ElseIf Sheet8.opt_NVis24.Value = True Then
            End If
            
            If Sheet8.opt_ToVis25.Value = True Then
                strto = strto + Range("jltacan").Value + "; "
            ElseIf Sheet8.opt_CCVis25.Value = True Then
                strcc = strcc + Range("jltacan").Value + "; "
            ElseIf Sheet8.opt_NVis25.Value = True Then
            End If
            
            If Sheet8.opt_ToVis26.Value = True Then
                strto = strto + Range("fotamarra").Value + "; "
            ElseIf Sheet8.opt_CCVis26.Value = True Then
                strcc = strcc + Range("fotamarra").Value + "; "
            ElseIf Sheet8.opt_NVis26.Value = True Then
            End If
            
            If Sheet8.opt_ToVis27.Value = True Then
                strto = strto + Range("ertejano").Value + "; "
            ElseIf Sheet8.opt_CCVis27.Value = True Then
                strcc = strcc + Range("ertejano").Value + "; "
            ElseIf Sheet8.opt_NVis27.Value = True Then
            End If
            
            If Sheet8.opt_ToVis28.Value = True Then
                strto = strto + Range("blynot").Value + "; "
            ElseIf Sheet8.opt_CCVis28.Value = True Then
                strcc = strcc + Range("blynot").Value + "; "
            ElseIf Sheet8.opt_NVis28.Value = True Then
            End If
            
            For i = 1 To Notif.lst_VisFxATOp.ListCount - 1
                s = Notif.lst_VisFxATOp.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B29").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'FCNO FF5 JNE
        ElseIf Notif.MultiPage9.Value = 1 And Notif.MultiPage9.Pages(1).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B30").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToFCNO.Value = True Then
                strto = strto + Range("afcapiral").Value + "; "
            ElseIf Sheet8.opt_CCFCNO.Value = True Then
                strcc = strcc + Range("afcapiral").Value + "; "
            ElseIf Sheet8.opt_NFCNO.Value = True Then
            End If
            
            If Sheet8.opt_ToCEBU1.Value = True Then
                strto = strto + Range("eddivinagracia").Value + "; "
            ElseIf Sheet8.opt_CCCEBU1.Value = True Then
                strcc = strcc + Range("eddivinagracia").Value + "; "
            ElseIf Sheet8.opt_NCEBU1.Value = True Then
            End If
            
            If Sheet8.opt_ToCEBU2.Value = True Then
                strto = strto + Range("rngloria").Value + "; "
            ElseIf Sheet8.opt_CCCEBU2.Value = True Then
                strcc = strcc + Range("rngloria").Value + "; "
            ElseIf Sheet8.opt_NCEBU2.Value = True Then
            End If
            
            If Sheet8.opt_ToCEBU3.Value = True Then
                strto = strto + Range("cminoferio").Value + "; "
            ElseIf Sheet8.opt_CCCEBU3.Value = True Then
                strcc = strcc + Range("cminoferio").Value + "; "
            ElseIf Sheet8.opt_NCEBU3.Value = True Then
            End If
            
            If Sheet8.opt_ToCEBU4.Value = True Then
                strto = strto + Range("jrmaninang").Value + "; "
            ElseIf Sheet8.opt_CCCEBU4.Value = True Then
                strcc = strcc + Range("jrmaninang").Value + "; "
            ElseIf Sheet8.opt_NCEBU4.Value = True Then
            End If
            
            If Sheet8.opt_ToCEBU5.Value = True Then
                strto = strto + Range("rrson").Value + "; "
            ElseIf Sheet8.opt_CCCEBU5.Value = True Then
                strcc = strcc + Range("rrson").Value + "; "
            ElseIf Sheet8.opt_NCEBU5.Value = True Then
            End If
            
            If Sheet8.opt_ToCEBU6.Value = True Then
                strto = strto + Range("mjvendiola").Value + "; "
            ElseIf Sheet8.opt_CCCEBU6.Value = True Then
                strcc = strcc + Range("mjvendiola").Value + "; "
            ElseIf Sheet8.opt_NCEBU6.Value = True Then
            End If
            
            For i = 1 To Notif.lst_FCNOJNE.ListCount - 1
                s = Notif.lst_FCNOJNE.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>F" & Sheet1.Range("B30").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
        End If
    End If
    
    'CONSOLIDATED
    If Notif.MultiPage1.Value = 7 Then
    'STRAT
        If Notif.MultiPage7.Value = 0 And Notif.MultiPage7.Pages(0).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B14").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToSTRAT.Value = True Then
                strto = strto + Range("aiseeco").Value + "; "
            ElseIf Sheet8.opt_CCSTRAT.Value = True Then
                strcc = strcc + Range("aiseeco").Value + "; "
            ElseIf Sheet8.opt_NSTRAT.Value = True Then
            End If
            
            If Sheet8.opt_ToSTRAT1.Value = True Then
                strto = strto + Range("bccordova").Value + "; "
            ElseIf Sheet8.opt_CCSTRAT1.Value = True Then
                strcc = strcc + Range("bccordova").Value + "; "
            ElseIf Sheet8.opt_NSTRAT1.Value = True Then
            End If
            
            If Sheet8.opt_ToSTRAT2.Value = True Then
                strto = strto + Range("jpjandayan").Value + "; "
            ElseIf Sheet8.opt_CCSTRAT2.Value = True Then
                strcc = strcc + Range("jpjandayan").Value + "; "
            ElseIf Sheet8.opt_NSTRAT2.Value = True Then
            End If
            
            For i = 1 To Notif.lst_STRATC.ListCount - 1
                s = Notif.lst_STRATC.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B14").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'ENG
        ElseIf Notif.MultiPage7.Value = 1 And Notif.MultiPage7.Pages(1).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B15").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToENG.Value = True Then
                strto = strto + Range("megutierrez").Value + "; "
            ElseIf Sheet8.opt_CCENG.Value = True Then
                strcc = strcc + Range("megutierrez").Value + "; "
            ElseIf Sheet8.opt_NENG.Value = True Then
            End If
            
            If Sheet8.opt_ToENG1.Value = True Then
                strto = strto + Range("amarsua").Value + "; "
            ElseIf Sheet8.opt_CCENG1.Value = True Then
                strcc = strcc + Range("amarsua").Value + "; "
            ElseIf Sheet8.opt_NENG1.Value = True Then
            End If
            
            If Sheet8.opt_ToENG2.Value = True Then
                strto = strto + Range("istolentino").Value + "; "
            ElseIf Sheet8.opt_CCENG2.Value = True Then
                strcc = strcc + Range("istolentino").Value + "; "
            ElseIf Sheet8.opt_NENG2.Value = True Then
            End If
            
            For i = 1 To Notif.lst_ENGC.ListCount - 1
                s = Notif.lst_ENGC.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B15").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
        End If
    End If
    
    If Not strto Like "*bccordova*" And strtable Like "*BC Scope (Specific for BU)*" _
        Or strtable Like "*BC Objectives*" _
        Or strtable Like "*Identification of Organization and its Context (4.1)*" _
        Or strtable Like "*Identification of Interested Parties and Their Needs (4.2)*" _
        Or strtable Like "*Business Impact Analysis Summary Report*" _
        Or strtable Like "*BIA Questionnaires*" _
        Or strtable Like "*Risk Assesssment*" _
        Or strtable Like "*BC Strategy*" Then
        If Sheet8.opt_ToSTRAT1.Value = True Or Sheet8.opt_CCSTRAT1.Value = True Then
            strcc = strcc + Range("bccordova").Value + "; "
        ElseIf Sheet8.opt_NSTRAT1.Value = True Then
        End If
    End If
    
    If Not strto Like "*jpjandayan*" And strtable Like "*BC Scope (Specific for BU)*" _
        Or strtable Like "*BC Objectives*" _
        Or strtable Like "*Identification of Organization and its Context (4.1)*" _
        Or strtable Like "*Identification of Interested Parties and Their Needs (4.2)*" _
        Or strtable Like "*Business Impact Analysis Summary Report*" _
        Or strtable Like "*BIA Questionnaires*" _
        Or strtable Like "*Risk Assesssment*" _
        Or strtable Like "*BC Strategy*" Then
        If Sheet8.opt_ToSTRAT2.Value = True Or Sheet8.opt_CCSTRAT2.Value = True Then
            strcc = strcc + Range("jpjandayan").Value + "; "
        ElseIf Sheet8.opt_NSTRAT2.Value = True Then
        End If
    End If
    
    If Not strto Like "*amarsua*" And strtable Like "*BC Plan-Business Unit*" Then
        If Sheet8.opt_ToENG1.Value = True Or Sheet8.opt_CCENG1.Value = True Then
            strcc = strcc + Range("amarsua").Value + "; "
        ElseIf Sheet8.opt_NENG1.Value = True Then
        End If
    End If
    
    If Not strto Like "*istolentino*" And strtable Like "*BC Plan-Business Unit*" Then
        If Sheet8.opt_ToENG2.Value = True Or Sheet8.opt_CCENG2.Value = True Then
            strcc = strcc + Range("istolentino").Value + "; "
        ElseIf Sheet8.opt_NENG2.Value = True Then
        End If
    End If
    
    If Sheet8.opt_ToGOV1.Value = True Then
        strto = strto + Range("janacpil").Value + "; "
    ElseIf Sheet8.opt_CCGOV1.Value = True Then
        strcc = strcc + Range("janacpil").Value + "; "
    ElseIf Sheet8.opt_NGOV1.Value = True Then
    End If
    
    If Not strto Like "*hclim*" Then
        If Sheet8.opt_ToGOV2.Value = True Then
            strto = strto + Range("hclim").Value + "; "
        ElseIf Sheet8.opt_CCGOV2.Value = True Then
            strcc = strcc + Range("hclim").Value + "; "
        ElseIf Sheet8.opt_NGOV2.Value = True Then
        End If
    End If
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error Resume Next
    
    signature = Environ("appdata") & "\Microsoft\Signatures\"
    If Dir(signature, vbDirectory) <> vbNullString Then
        signature = signature & Dir$(signature & "*.htm")
    Else:
        signature = ""
    End If
    signature = CreateObject("Scripting.FileSystemObject").GetFile(signature).OpenAsTextStream(1, -2).ReadAll
    OutMail.HTMLBody = signature
    
    With OutMail
        .To = strto
        .CC = strcc
        .BCC = ""
        .Subject = "BCMS Documents Update"
        .Body = strbody & vbCrLf & Range("AddNote")
        .Attachments.Add (Range("FileSource").Value)
        .HTMLBody = .HTMLBody & "<table border='1' cellspacing='0' cellpadding='0' style='width:100%; border:1px solid black; border-collapse:collapse; text-align:center; font-family:calibri; font-size:14.5px'>" _
            & "<tbody><tr><td rowspan='2' ><p><b>BUSINESS UNIT/Document Owner</b></p></td>" _
            & "<td rowspan='2' ><p><b>DOCUMENT TITLE</b></p></td>" _
            & "<td rowspan='2' ><p><b>EFFECTIVITY DATE</b></p></td>" _
            & "<td rowspan='2' ><p><b>EXPIRATION DATE</b></p></td>" _
            & "<td colspan='2' style='background-color: yellow'><p><b>NEED UPDATE?</b></p></td>" _
            & "<td rowspan='2' style='background-color: yellow'><p><b>REMARKS</b><br><i style='font-size:13.5'>(e.g. reason why updating is not needed)</i></p></td></tr>" _
            & "<tr><td valign='top' style='background-color: yellow'><p><i>YES</i></p></td>" _
            & "<td valign='top' style='background-color: yellow'><p><i>NO</i></p></td></tr>" _
            & strtable & "</tbody></table>" & "<br>Thank You.<br>Regards,<br>" & signature
        .Send
    End With
    On Error GoTo 0
    
    Set OutMail = Nothing
    Set OutApp = Nothing
    MsgBox "Email Sent", vbInformation, "Success"
End Sub


Private Sub MultiPage1_Change()
    If Notif.MultiPage1.Value = 0 Then
        Sheet2.Activate
    ElseIf Notif.MultiPage1.Value = 1 Then
        Sheet3.Activate
    ElseIf Notif.MultiPage1.Value = 2 Then
        Sheet4.Activate
    ElseIf Notif.MultiPage1.Value = 3 Then
        Sheet6.Activate
    ElseIf Notif.MultiPage1.Value = 4 Then
        Sheet5.Activate
    ElseIf Notif.MultiPage1.Value = 5 Then
        Sheet11.Activate
    ElseIf Notif.MultiPage1.Value = 6 Then
        Sheet12.Activate
    ElseIf Notif.MultiPage1.Value = 7 Then
        Sheet7.Activate
    End If
End Sub

Private Sub btn_Directory_Click()
    Sheet8.Select
    End
End Sub

Private Sub btn_Source_Click()
    Sheet10.Select
    End
End Sub

Private Sub btn_View_Click()
    Dim strbody As String
    Dim strtable As String
    Dim strto As String
    Dim strcc As String
    Dim s As String
    Dim i As Integer
    Dim signature As String
    
    If Sheet8.opt_ToGOV.Value = True Then
        strto = strto + Range("rdvillasenor").Value + "; "
    ElseIf Sheet8.opt_CCGOV.Value = True Then
        strcc = strcc + Range("rdvillasenor").Value + "; "
    ElseIf Sheet8.opt_NGOV.Value = True Then
    End If
          
    'GPs
    If Notif.MultiPage1.Value = 0 Then
    'GOV
        If Notif.MultiPage2.Value = 0 And Notif.MultiPage2.Pages(0).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B13").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            If Sheet8.opt_ToGOV2.Value = True Or Sheet8.opt_CCGOV2.Value = True Then
                strto = strto + Range("hclim").Value + "; "
            ElseIf Sheet8.opt_NGOV2.Value = True Then
            End If
            For i = 1 To Notif.lst_GOV.ListCount - 1
                s = Notif.lst_GOV.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B13").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'STRAT
        ElseIf Notif.MultiPage2.Value = 1 And Notif.MultiPage2.Pages(1).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B14").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToSTRAT.Value = True Then
                strto = strto + Range("aiseeco").Value + "; "
            ElseIf Sheet8.opt_CCSTRAT.Value = True Then
                strcc = strcc + Range("aiseeco").Value + "; "
            ElseIf Sheet8.opt_NSTRAT.Value = True Then
            End If
            
            If Sheet8.opt_ToSTRAT1.Value = True Then
                strto = strto + Range("bccordova").Value + "; "
            ElseIf Sheet8.opt_CCSTRAT1.Value = True Then
                strcc = strcc + Range("bccordova").Value + "; "
            ElseIf Sheet8.opt_NSTRAT1.Value = True Then
            End If
            
            If Sheet8.opt_ToSTRAT2.Value = True Then
                strto = strto + Range("jpjandayan").Value + "; "
            ElseIf Sheet8.opt_CCSTRAT2.Value = True Then
                strcc = strcc + Range("jpjandayan").Value + "; "
            ElseIf Sheet8.opt_NSTRAT2.Value = True Then
            End If
            
            For i = 1 To Notif.lst_STRAT.ListCount - 1
                s = Notif.lst_STRAT.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B14").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'ENG
        ElseIf Notif.MultiPage2.Value = 2 And Notif.MultiPage2.Pages(2).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B15").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToENG.Value = True Then
                strto = strto + Range("megutierrez").Value + "; "
            ElseIf Sheet8.opt_CCENG.Value = True Then
                strcc = strcc + Range("megutierrez").Value + "; "
            ElseIf Sheet8.opt_NENG.Value = True Then
            End If
            
            If Sheet8.opt_ToENG1.Value = True Then
                strto = strto + Range("amarsua").Value + "; "
            ElseIf Sheet8.opt_CCENG1.Value = True Then
                strcc = strcc + Range("amarsua").Value + "; "
            ElseIf Sheet8.opt_NENG1.Value = True Then
            End If
            
            If Sheet8.opt_ToENG2.Value = True Then
                strto = strto + Range("istolentino").Value + "; "
            ElseIf Sheet8.opt_CCENG2.Value = True Then
                strcc = strcc + Range("istolentino").Value + "; "
            ElseIf Sheet8.opt_NENG2.Value = True Then
            End If
            
            For i = 1 To Notif.lst_ENG.ListCount - 1
                s = Notif.lst_ENG.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B15").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
        End If
    End If
    
    'CLS
    If Notif.MultiPage1.Value = 1 And Notif.MultiPage1.Pages(1).Caption Like "*" Then
        If Sheet8.opt_ToCLS.Value = True Then
            strto = strto + Range("emgacayan").Value + "; "
        ElseIf Sheet8.opt_CCCLS.Value = True Then
            strcc = strcc + Range("emgacayan").Value + "; "
        ElseIf Sheet8.opt_NCLS.Value = True Then
        End If
    'LUCLS
        If Notif.MultiPage3.Value = 0 And Notif.MultiPage3.Pages(0).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B16").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToLUCLS1.Value = True Then
                strto = strto + Range("jajacinto").Value + "; "
            ElseIf Sheet8.opt_CCLUCLS1.Value = True Then
                strcc = strcc + Range("jajacinto").Value + "; "
            ElseIf Sheet8.opt_NLUCLS1.Value = True Then
            End If
            
            If Sheet8.opt_ToLUCLS2.Value = True Then
                strto = strto + Range("ljechave").Value + "; "
            ElseIf Sheet8.opt_CCLUCLS2.Value = True Then
                strcc = strcc + Range("ljechave").Value + "; "
            ElseIf Sheet8.opt_NLUCLS2.Value = True Then
            End If
            
            If Sheet8.opt_ToLUCLS3.Value = True Then
                strto = strto + Range("ecganuelas").Value + "; "
            ElseIf Sheet8.opt_CCLUCLS3.Value = True Then
                strcc = strcc + Range("ecganuelas").Value + "; "
            ElseIf Sheet8.opt_NLUCLS3.Value = True Then
            End If
            
            If Sheet8.opt_ToLUCLS4.Value = True Then
                strto = strto + Range("mmginete").Value + "; "
            ElseIf Sheet8.opt_CCLUCLS4.Value = True Then
                strcc = strcc + Range("mmginete").Value + "; "
            ElseIf Sheet8.opt_NLUCLS4.Value = True Then
            End If
            
            If Sheet8.opt_ToLUCLS5.Value = True Then
                strto = strto + Range("jogutierrez").Value + "; "
            ElseIf Sheet8.opt_CCLUCLS5.Value = True Then
                strcc = strcc + Range("jogutierrez").Value + "; "
            ElseIf Sheet8.opt_NLUCLS5.Value = True Then
            End If
            
            If Sheet8.opt_ToLUCLS6.Value = True Then
                strto = strto + Range("crcmendoza").Value + "; "
            ElseIf Sheet8.opt_CCLUCLS6.Value = True Then
                strcc = strcc + Range("crcmendoza").Value + "; "
            ElseIf Sheet8.opt_NLUCLS6.Value = True Then
            End If
            
            If Sheet8.opt_ToLUCLS7.Value = True Then
                strto = strto + Range("jlmoldez").Value + "; "
            ElseIf Sheet8.opt_CCLUCLS7.Value = True Then
                strcc = strcc + Range("jlmoldez").Value + "; "
            ElseIf Sheet8.opt_NLUCLS7.Value = True Then
            End If
            
            If Sheet8.opt_ToLUCLS8.Value = True Then
                strto = strto + Range("mmmones").Value + "; "
            ElseIf Sheet8.opt_CCLUCLS8.Value = True Then
                strcc = strcc + Range("mmmones").Value + "; "
            ElseIf Sheet8.opt_NLUCLS8.Value = True Then
            End If
            
            If Sheet8.opt_ToLUCLS9.Value = True Then
                strto = strto + Range("vprodriguez").Value + "; "
            ElseIf Sheet8.opt_CCLUCLS9.Value = True Then
                strcc = strcc + Range("vprodriguez").Value + "; "
            ElseIf Sheet8.opt_NLUCLS9.Value = True Then
            End If
                        
            For i = 1 To Notif.lst_LUCLS.ListCount - 1
                s = Notif.lst_LUCLS.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B16").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'BCLS
        ElseIf Notif.MultiPage3.Value = 1 And Notif.MultiPage3.Pages(1).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B17").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToBCLS1.Value = True Then
                strto = strto + Range("rjenriquez").Value + "; "
            ElseIf Sheet8.opt_CCBCLS1.Value = True Then
                strcc = strcc + Range("rjenriquez").Value + "; "
            ElseIf Sheet8.opt_NBCLS1.Value = True Then
            End If
            
            If Sheet8.opt_ToBCLS2.Value = True Then
                strto = strto + Range("jdbustamante").Value + "; "
            ElseIf Sheet8.opt_CCBCLS2.Value = True Then
                strcc = strcc + Range("jdbustamante").Value + "; "
            ElseIf Sheet8.opt_NBCLS2.Value = True Then
            End If
            
            If Sheet8.opt_ToBCLS3.Value = True Then
                strto = strto + Range("apcaringal").Value + "; "
            ElseIf Sheet8.opt_CCBCLS3.Value = True Then
                strcc = strcc + Range("apcaringal").Value + "; "
            ElseIf Sheet8.opt_NBCLS3.Value = True Then
            End If
            
            If Sheet8.opt_ToBCLS4.Value = True Then
                strto = strto + Range("jacatibog").Value + "; "
            ElseIf Sheet8.opt_CCBCLS4.Value = True Then
                strcc = strcc + Range("jacatibog").Value + "; "
            ElseIf Sheet8.opt_NBCLS4.Value = True Then
            End If
            
            If Sheet8.opt_ToBCLS5.Value = True Then
                strto = strto + Range("ebdeleon").Value + "; "
            ElseIf Sheet8.opt_CCBCLS5.Value = True Then
                strcc = strcc + Range("ebdeleon").Value + "; "
            ElseIf Sheet8.opt_NBCLS5.Value = True Then
            End If
            
            If Sheet8.opt_ToBCLS6.Value = True Then
                strto = strto + Range("pretcobanez").Value + "; "
            ElseIf Sheet8.opt_CCBCLS6.Value = True Then
                strcc = strcc + Range("pretcobanez").Value + "; "
            ElseIf Sheet8.opt_NBCLS6.Value = True Then
            End If
            
            If Sheet8.opt_ToBCLS7.Value = True Then
                strto = strto + Range("rpmanago").Value + "; "
            ElseIf Sheet8.opt_CCBCLS7.Value = True Then
                strcc = strcc + Range("rpmanago").Value + "; "
            ElseIf Sheet8.opt_NBCLS7.Value = True Then
            End If
            
            If Sheet8.opt_ToBCLS8.Value = True Then
                strto = strto + Range("hdpilar").Value + "; "
            ElseIf Sheet8.opt_CCBCLS8.Value = True Then
                strcc = strcc + Range("hdpilar").Value + "; "
            ElseIf Sheet8.opt_NBCLS8.Value = True Then
            End If
            
            If Sheet8.opt_ToBCLS9.Value = True Then
                strto = strto + Range("acreyes").Value + "; "
            ElseIf Sheet8.opt_CCBCLS9.Value = True Then
                strcc = strcc + Range("acreyes").Value + "; "
            ElseIf Sheet8.opt_NBCLS9.Value = True Then
            End If
            
            If Sheet8.opt_ToBCLS10.Value = True Then
                strto = strto + Range("aevizconde").Value + "; "
            ElseIf Sheet8.opt_CCBCLS10.Value = True Then
                strcc = strcc + Range("aevizconde").Value + "; "
            ElseIf Sheet8.opt_NBCLS10.Value = True Then
            End If
            
            For i = 1 To Notif.lst_BCLS.ListCount - 1
                s = Notif.lst_BCLS.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B17").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'DCLS
        ElseIf Notif.MultiPage3.Value = 2 And Notif.MultiPage3.Pages(2).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B18").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToDCLS1.Value = True Then
                strto = strto + Range("vvpunzalan").Value + "; "
            ElseIf Sheet8.opt_CCDCLS1.Value = True Then
                strcc = strcc + Range("vvpunzalan").Value + "; "
            ElseIf Sheet8.opt_NDCLS1.Value = True Then
            End If
            
            If Sheet8.opt_ToDCLS4.Value = True Then
                strto = strto + Range("rspanta").Value + "; "
            ElseIf Sheet8.opt_CCDCLS4.Value = True Then
                strcc = strcc + Range("rspanta").Value + "; "
            ElseIf Sheet8.opt_NDCLS4.Value = True Then
            End If
            
            If Sheet8.opt_ToDCLS2.Value = True Then
                strto = strto + Range("jlandicoy").Value + "; "
            ElseIf Sheet8.opt_CCDCLS2.Value = True Then
                strcc = strcc + Range("jlandicoy").Value + "; "
            ElseIf Sheet8.opt_NDCLS2.Value = True Then
            End If
            
            If Sheet8.opt_ToDCLS3.Value = True Then
                strto = strto + Range("mbdevilla").Value + "; "
            ElseIf Sheet8.opt_CCDCLS3.Value = True Then
                strcc = strcc + Range("mbdevilla").Value + "; "
            ElseIf Sheet8.opt_NDCLS3.Value = True Then
            End If
            
            If Sheet8.opt_ToDCLS5.Value = True Then
                strto = strto + Range("neromero").Value + "; "
            ElseIf Sheet8.opt_CCDCLS5.Value = True Then
                strcc = strcc + Range("neromero").Value + "; "
            ElseIf Sheet8.opt_NDCLS5.Value = True Then
            End If
            
            If Sheet8.opt_ToDCLS6.Value = True Then
                strto = strto + Range("josalvador").Value + "; "
            ElseIf Sheet8.opt_CCDCLS6.Value = True Then
                strcc = strcc + Range("josalvador").Value + "; "
            ElseIf Sheet8.opt_NDCLS6.Value = True Then
            End If
            
            If Sheet8.opt_ToDCLS7.Value = True Then
                strto = strto + Range("dltatad").Value + "; "
            ElseIf Sheet8.opt_CCDCLS7.Value = True Then
                strcc = strcc + Range("dltatad").Value + "; "
            ElseIf Sheet8.opt_NDCLS7.Value = True Then
            End If
            
            For i = 1 To Notif.lst_DCLS.ListCount - 1
                s = Notif.lst_DCLS.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B18").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
        End If
    End If
    
    'DILIMAN
    If Notif.MultiPage1.Value = 2 Then
    'NGMMFxATOp
        If Notif.MultiPage4.Value = 0 And Notif.MultiPage4.Pages(0).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B19").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToNGMM.Value = True Then
                strto = strto + Range("dcroque").Value + "; "
            ElseIf Sheet8.opt_CCNGMM.Value = True Then
                strcc = strcc + Range("dcroque").Value + "; "
            ElseIf Sheet8.opt_NNGMM.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM1.Value = True Then
                strto = strto + Range("agsoro").Value + "; "
            ElseIf Sheet8.opt_CCNGMM1.Value = True Then
                strcc = strcc + Range("agsoro").Value + "; "
            ElseIf Sheet8.opt_NNGMM1.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM2.Value = True Then
                strto = strto + Range("aoabary").Value + "; "
            ElseIf Sheet8.opt_CCNGMM2.Value = True Then
                strcc = strcc + Range("aoabary").Value + "; "
            ElseIf Sheet8.opt_NNGMM2.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM3.Value = True Then
                strto = strto + Range("rogbanez").Value + "; "
            ElseIf Sheet8.opt_CCNGMM3.Value = True Then
                strcc = strcc + Range("rogbanez").Value + "; "
            ElseIf Sheet8.opt_NNGMM3.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM4.Value = True Then
                strto = strto + Range("blbautista").Value + "; "
            ElseIf Sheet8.opt_CCNGMM4.Value = True Then
                strcc = strcc + Range("blbautista").Value + "; "
            ElseIf Sheet8.opt_NNGMM4.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM5.Value = True Then
                strto = strto + Range("ebbayle").Value + "; "
            ElseIf Sheet8.opt_CCNGMM5.Value = True Then
                strcc = strcc + Range("ebbayle").Value + "; "
            ElseIf Sheet8.opt_NNGMM5.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM6.Value = True Then
                strto = strto + Range("accruz").Value + "; "
            ElseIf Sheet8.opt_CCNGMM6.Value = True Then
                strcc = strcc + Range("accruz").Value + "; "
            ElseIf Sheet8.opt_NNGMM6.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM7.Value = True Then
                strto = strto + Range("mgdioso").Value + "; "
            ElseIf Sheet8.opt_CCNGMM7.Value = True Then
                strcc = strcc + Range("mgdioso").Value + "; "
            ElseIf Sheet8.opt_NNGMM7.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM8.Value = True Then
                strto = strto + Range("amestrella").Value + "; "
            ElseIf Sheet8.opt_CCNGMM8.Value = True Then
                strcc = strcc + Range("amestrella").Value + "; "
            ElseIf Sheet8.opt_NNGMM8.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM9.Value = True Then
                strto = strto + Range("drfonacier").Value + "; "
            ElseIf Sheet8.opt_CCNGMM9.Value = True Then
                strcc = strcc + Range("drfonacier").Value + "; "
            ElseIf Sheet8.opt_NNGMM9.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM10.Value = True Then
                strto = strto + Range("rjlim").Value + "; "
            ElseIf Sheet8.opt_CCNGMM10.Value = True Then
                strcc = strcc + Range("rjlim").Value + "; "
            ElseIf Sheet8.opt_NNGMM10.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM11.Value = True Then
                strto = strto + Range("ecmadrilejo").Value + "; "
            ElseIf Sheet8.opt_CCNGMM11.Value = True Then
                strcc = strcc + Range("ecmadrilejo").Value + "; "
            ElseIf Sheet8.opt_NNGMM11.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM12.Value = True Then
                strto = strto + Range("jvnaval").Value + "; "
            ElseIf Sheet8.opt_CCNGMM12.Value = True Then
                strcc = strcc + Range("jvnaval").Value + "; "
            ElseIf Sheet8.opt_NNGMM12.Value = True Then
            End If
            
            If Sheet8.opt_ToNGMM13.Value = True Then
                strto = strto + Range("wmsabile").Value + "; "
            ElseIf Sheet8.opt_CCNGMM13.Value = True Then
                strcc = strcc + Range("wmsabile").Value + "; "
            ElseIf Sheet8.opt_NNGMM13.Value = True Then
            End If
            
            For i = 1 To Notif.lst_NGMMFxATOp.ListCount - 1
                s = Notif.lst_NGMMFxATOp.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B19").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'FCNO FF2 QCY
        ElseIf Notif.MultiPage4.Value = 1 And Notif.MultiPage4.Pages(1).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B20").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToFCNO.Value = True Then
                strto = strto + Range("afcapiral").Value + "; "
            ElseIf Sheet8.opt_CCFCNO.Value = True Then
                strcc = strcc + Range("afcapiral").Value + "; "
            ElseIf Sheet8.opt_NFCNO.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO1.Value = True Then
                strto = strto + Range("picables").Value + "; "
            ElseIf Sheet8.opt_CCFCNO1.Value = True Then
                strcc = strcc + Range("picables").Value + "; "
            ElseIf Sheet8.opt_NFCNO1.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO2.Value = True Then
                strto = strto + Range("fcbaul").Value + "; "
            ElseIf Sheet8.opt_CCFCNO2.Value = True Then
                strcc = strcc + Range("fcbaul").Value + "; "
            ElseIf Sheet8.opt_NFCNO2.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO3.Value = True Then
                strto = strto + Range("aljimenez").Value + "; "
            ElseIf Sheet8.opt_CCFCNO3.Value = True Then
                strcc = strcc + Range("aljimenez").Value + "; "
            ElseIf Sheet8.opt_NFCNO3.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO4.Value = True Then
                strto = strto + Range("rtlampa").Value + "; "
            ElseIf Sheet8.opt_CCFCNO4.Value = True Then
                strcc = strcc + Range("rtlampa").Value + "; "
            ElseIf Sheet8.opt_NFCNO4.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO5.Value = True Then
                strto = strto + Range("oalinco").Value + "; "
            ElseIf Sheet8.opt_CCFCNO5.Value = True Then
                strcc = strcc + Range("oalinco").Value + "; "
            ElseIf Sheet8.opt_NFCNO5.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO6.Value = True Then
                strto = strto + Range("mfsantos").Value + "; "
            ElseIf Sheet8.opt_CCFCNO6.Value = True Then
                strcc = strcc + Range("mfsantos").Value + "; "
            ElseIf Sheet8.opt_NFCNO6.Value = True Then
            End If
            
            For i = 1 To Notif.lst_FCNOD.ListCount - 1
                s = Notif.lst_FCNOD.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B20").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'IGO QC DFON Station
        ElseIf Notif.MultiPage4.Value = 2 And Notif.MultiPage4.Pages(2).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B21").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToIGO.Value = True Then
                strto = strto + Range("asgaba").Value + "; "
            ElseIf Sheet8.opt_CCIGO.Value = True Then
                strcc = strcc + Range("asgaba").Value + "; "
            ElseIf Sheet8.opt_NIGO.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO1.Value = True Then
                strto = strto + Range("rdvizmanos").Value + "; "
            ElseIf Sheet8.opt_CCIGO1.Value = True Then
                strcc = strcc + Range("rdvizmanos").Value + "; "
            ElseIf Sheet8.opt_NIGO1.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO2.Value = True Then
                strto = strto + Range("csalejo").Value + "; "
            ElseIf Sheet8.opt_CCIGO2.Value = True Then
                strcc = strcc + Range("csalejo").Value + "; "
            ElseIf Sheet8.opt_NIGO2.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO3.Value = True Then
                strto = strto + Range("ecbuera").Value + "; "
            ElseIf Sheet8.opt_CCIGO3.Value = True Then
                strcc = strcc + Range("ecbuera").Value + "; "
            ElseIf Sheet8.opt_NIGO3.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO4.Value = True Then
                strto = strto + Range("kdcruz").Value + "; "
            ElseIf Sheet8.opt_CCIGO4.Value = True Then
                strcc = strcc + Range("kdcruz").Value + "; "
            ElseIf Sheet8.opt_NIGO4.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO5.Value = True Then
                strto = strto + Range("wdgdebelen").Value + "; "
            ElseIf Sheet8.opt_CCIGO5.Value = True Then
                strcc = strcc + Range("wdgdebelen").Value + "; "
            ElseIf Sheet8.opt_NIGO5.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO6.Value = True Then
                strto = strto + Range("nmestevez").Value + "; "
            ElseIf Sheet8.opt_CCIGO6.Value = True Then
                strcc = strcc + Range("nmestevez").Value + "; "
            ElseIf Sheet8.opt_NIGO6.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO7.Value = True Then
                strto = strto + Range("vdgrino").Value + "; "
            ElseIf Sheet8.opt_CCIGO7.Value = True Then
                strcc = strcc + Range("vdgrino").Value + "; "
            ElseIf Sheet8.opt_NIGO7.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO8.Value = True Then
                strto = strto + Range("jmhernandez").Value + "; "
            ElseIf Sheet8.opt_CCIGO8.Value = True Then
                strcc = strcc + Range("jmhernandez").Value + "; "
            ElseIf Sheet8.opt_NIGO8.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO9.Value = True Then
                strto = strto + Range("anmantes").Value + "; "
            ElseIf Sheet8.opt_CCIGO9.Value = True Then
                strcc = strcc + Range("anmantes").Value + "; "
            ElseIf Sheet8.opt_NIGO9.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO10.Value = True Then
                strto = strto + Range("lgpagay").Value + "; "
            ElseIf Sheet8.opt_CCIGO10.Value = True Then
                strcc = strcc + Range("lgpagay").Value + "; "
            ElseIf Sheet8.opt_NIGO10.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO11.Value = True Then
                strto = strto + Range("rjtorio").Value + "; "
            ElseIf Sheet8.opt_CCIGO11.Value = True Then
                strcc = strcc + Range("rjtorio").Value + "; "
            ElseIf Sheet8.opt_NIGO11.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO12.Value = True Then
                strto = strto + Range("vauypala").Value + "; "
            ElseIf Sheet8.opt_CCIGO12.Value = True Then
                strcc = strcc + Range("vauypala").Value + "; "
            ElseIf Sheet8.opt_NIGO12.Value = True Then
            End If
            
            For i = 1 To Notif.lst_IGOD.ListCount - 1
                s = Notif.lst_IGOD.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B21").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
        End If
    End If
    
    'SAMPALOC
    If Notif.MultiPage1.Value = 3 Then
    'WGMMFxATOp
        If Notif.MultiPage5.Value = 0 And Notif.MultiPage5.Pages(0).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B22").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToWGMM.Value = True Then
                strto = strto + Range("emalsol").Value + "; "
            ElseIf Sheet8.opt_CCWGMM.Value = True Then
                strcc = strcc + Range("emalsol").Value + "; "
            ElseIf Sheet8.opt_NWGMM.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM1.Value = True Then
                strto = strto + Range("nnelegado").Value + "; "
            ElseIf Sheet8.opt_CCWGMM1.Value = True Then
                strcc = strcc + Range("nnelegado").Value + "; "
            ElseIf Sheet8.opt_NWGMM1.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM2.Value = True Then
                strto = strto + Range("jnadia").Value + "; "
            ElseIf Sheet8.opt_CCWGMM2.Value = True Then
                strcc = strcc + Range("jnadia").Value + "; "
            ElseIf Sheet8.opt_NWGMM2.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM3.Value = True Then
                strto = strto + Range("aaagbayani").Value + "; "
            ElseIf Sheet8.opt_CCWGMM3.Value = True Then
                strcc = strcc + Range("aaagbayani").Value + "; "
            ElseIf Sheet8.opt_NWGMM3.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM4.Value = True Then
                strto = strto + Range("gpaquino").Value + "; "
            ElseIf Sheet8.opt_CCWGMM4.Value = True Then
                strcc = strcc + Range("gpaquino").Value + "; "
            ElseIf Sheet8.opt_NWGMM4.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM5.Value = True Then
                strto = strto + Range("rhatendido").Value + "; "
            ElseIf Sheet8.opt_CCWGMM5.Value = True Then
                strcc = strcc + Range("rhatendido").Value + "; "
            ElseIf Sheet8.opt_NWGMM5.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM6.Value = True Then
                strto = strto + Range("almgonzales").Value + "; "
            ElseIf Sheet8.opt_CCWGMM6.Value = True Then
                strcc = strcc + Range("almgonzales").Value + "; "
            ElseIf Sheet8.opt_NWGMM6.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM7.Value = True Then
                strto = strto + Range("rvhundana").Value + "; "
            ElseIf Sheet8.opt_CCWGMM7.Value = True Then
                strcc = strcc + Range("rvhundana").Value + "; "
            ElseIf Sheet8.opt_NWGMM7.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM8.Value = True Then
                strto = strto + Range("famacadaeg").Value + "; "
            ElseIf Sheet8.opt_CCWGMM8.Value = True Then
                strcc = strcc + Range("famacadaeg").Value + "; "
            ElseIf Sheet8.opt_NWGMM8.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM9.Value = True Then
                strto = strto + Range("rsmariano").Value + "; "
            ElseIf Sheet8.opt_CCWGMM9.Value = True Then
                strcc = strcc + Range("rsmariano").Value + "; "
            ElseIf Sheet8.opt_NWGMM9.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM10.Value = True Then
                strto = strto + Range("aenito").Value + "; "
            ElseIf Sheet8.opt_CCWGMM10.Value = True Then
                strcc = strcc + Range("aenito").Value + "; "
            ElseIf Sheet8.opt_NWGMM10.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM11.Value = True Then
                strto = strto + Range("papagtalunan").Value + "; "
            ElseIf Sheet8.opt_CCWGMM11.Value = True Then
                strcc = strcc + Range("papagtalunan").Value + "; "
            ElseIf Sheet8.opt_NWGMM11.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM12.Value = True Then
                strto = strto + Range("lpparadero").Value + "; "
            ElseIf Sheet8.opt_CCWGMM12.Value = True Then
                strcc = strcc + Range("lpparadero").Value + "; "
            ElseIf Sheet8.opt_NWGMM12.Value = True Then
            End If
            
            If Sheet8.opt_ToWGMM13.Value = True Then
                strto = strto + Range("basoriano").Value + "; "
            ElseIf Sheet8.opt_CCWGMM13.Value = True Then
                strcc = strcc + Range("basoriano").Value + "; "
            ElseIf Sheet8.opt_NWGMM13.Value = True Then
            End If
            
            For i = 1 To Notif.lst_WGMMFxATOp.ListCount - 1
                s = Notif.lst_WGMMFxATOp.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B22").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'FCNO FF2 SPC
        ElseIf Notif.MultiPage5.Value = 1 And Notif.MultiPage5.Pages(1).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B23").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToFCNO.Value = True Then
                strto = strto + Range("afcapiral").Value + "; "
            ElseIf Sheet8.opt_CCFCNO.Value = True Then
                strcc = strcc + Range("afcapiral").Value + "; "
            ElseIf Sheet8.opt_NFCNO.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO1.Value = True Then
                strto = strto + Range("picables").Value + "; "
            ElseIf Sheet8.opt_CCFCNO1.Value = True Then
                strcc = strcc + Range("picables").Value + "; "
            ElseIf Sheet8.opt_NFCNO1.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO2.Value = True Then
                strto = strto + Range("fcbaul").Value + "; "
            ElseIf Sheet8.opt_CCFCNO2.Value = True Then
                strcc = strcc + Range("fcbaul").Value + "; "
            ElseIf Sheet8.opt_NFCNO2.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO3.Value = True Then
                strto = strto + Range("aljimenez").Value + "; "
            ElseIf Sheet8.opt_CCFCNO3.Value = True Then
                strcc = strcc + Range("aljimenez").Value + "; "
            ElseIf Sheet8.opt_NFCNO3.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO4.Value = True Then
                strto = strto + Range("rtlampa").Value + "; "
            ElseIf Sheet8.opt_CCFCNO4.Value = True Then
                strcc = strcc + Range("rtlampa").Value + "; "
            ElseIf Sheet8.opt_NFCNO4.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO5.Value = True Then
                strto = strto + Range("oalinco").Value + "; "
            ElseIf Sheet8.opt_CCFCNO5.Value = True Then
                strcc = strcc + Range("oalinco").Value + "; "
            ElseIf Sheet8.opt_NFCNO5.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO6.Value = True Then
                strto = strto + Range("mfsantos").Value + "; "
            ElseIf Sheet8.opt_CCFCNO6.Value = True Then
                strcc = strcc + Range("mfsantos").Value + "; "
            ElseIf Sheet8.opt_NFCNO6.Value = True Then
            End If
            
            For i = 1 To Notif.lst_FCNOS.ListCount - 1
                s = Notif.lst_FCNOS.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B23").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'Manila IGO
        ElseIf Notif.MultiPage5.Value = 2 And Notif.MultiPage5.Pages(2).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B24").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToIGO.Value = True Then
                strto = strto + Range("asgaba").Value + "; "
            ElseIf Sheet8.opt_CCIGO.Value = True Then
                strcc = strcc + Range("asgaba").Value + "; "
            ElseIf Sheet8.opt_NIGO.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO1.Value = True Then
                strto = strto + Range("rdvizmanos").Value + "; "
            ElseIf Sheet8.opt_CCIGO1.Value = True Then
                strcc = strcc + Range("rdvizmanos").Value + "; "
            ElseIf Sheet8.opt_NIGO1.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO2.Value = True Then
                strto = strto + Range("csalejo").Value + "; "
            ElseIf Sheet8.opt_CCIGO2.Value = True Then
                strcc = strcc + Range("csalejo").Value + "; "
            ElseIf Sheet8.opt_NIGO2.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO3.Value = True Then
                strto = strto + Range("ecbuera").Value + "; "
            ElseIf Sheet8.opt_CCIGO3.Value = True Then
                strcc = strcc + Range("ecbuera").Value + "; "
            ElseIf Sheet8.opt_NIGO3.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO4.Value = True Then
                strto = strto + Range("kdcruz").Value + "; "
            ElseIf Sheet8.opt_CCIGO4.Value = True Then
                strcc = strcc + Range("kdcruz").Value + "; "
            ElseIf Sheet8.opt_NIGO4.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO5.Value = True Then
                strto = strto + Range("wdgdebelen").Value + "; "
            ElseIf Sheet8.opt_CCIGO5.Value = True Then
                strcc = strcc + Range("wdgdebelen").Value + "; "
            ElseIf Sheet8.opt_NIGO5.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO6.Value = True Then
                strto = strto + Range("nmestevez").Value + "; "
            ElseIf Sheet8.opt_CCIGO6.Value = True Then
                strcc = strcc + Range("nmestevez").Value + "; "
            ElseIf Sheet8.opt_NIGO6.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO7.Value = True Then
                strto = strto + Range("vdgrino").Value + "; "
            ElseIf Sheet8.opt_CCIGO7.Value = True Then
                strcc = strcc + Range("vdgrino").Value + "; "
            ElseIf Sheet8.opt_NIGO7.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO8.Value = True Then
                strto = strto + Range("jmhernandez").Value + "; "
            ElseIf Sheet8.opt_CCIGO8.Value = True Then
                strcc = strcc + Range("jmhernandez").Value + "; "
            ElseIf Sheet8.opt_NIGO8.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO9.Value = True Then
                strto = strto + Range("anmantes").Value + "; "
            ElseIf Sheet8.opt_CCIGO9.Value = True Then
                strcc = strcc + Range("anmantes").Value + "; "
            ElseIf Sheet8.opt_NIGO9.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO10.Value = True Then
                strto = strto + Range("lgpagay").Value + "; "
            ElseIf Sheet8.opt_CCIGO10.Value = True Then
                strcc = strcc + Range("lgpagay").Value + "; "
            ElseIf Sheet8.opt_NIGO10.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO11.Value = True Then
                strto = strto + Range("rjtorio").Value + "; "
            ElseIf Sheet8.opt_CCIGO11.Value = True Then
                strcc = strcc + Range("rjtorio").Value + "; "
            ElseIf Sheet8.opt_NIGO11.Value = True Then
            End If
            
            If Sheet8.opt_ToIGO12.Value = True Then
                strto = strto + Range("vauypala").Value + "; "
            ElseIf Sheet8.opt_CCIGO12.Value = True Then
                strcc = strcc + Range("vauypala").Value + "; "
            ElseIf Sheet8.opt_NIGO12.Value = True Then
            End If
            
            For i = 1 To Notif.lst_IGOS.ListCount - 1
                s = Notif.lst_IGOS.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B24").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
        End If
    End If
    
    'GREENHILLS
    If Notif.MultiPage1.Value = 4 Then
    'EGMMFxATOp GHL
        If Notif.MultiPage6.Value = 0 And Notif.MultiPage6.Pages(0).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B25").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToEGMM.Value = True Then
                strto = strto + Range("ntalcantara").Value + "; "
            ElseIf Sheet8.opt_CCEGMM.Value = True Then
                strcc = strcc + Range("ntalcantara").Value + "; "
            ElseIf Sheet8.opt_NEGMM.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM1.Value = True Then
                strto = strto + Range("armadrelino").Value + "; "
            ElseIf Sheet8.opt_CCEGMM1.Value = True Then
                strcc = strcc + Range("armadrelino").Value + "; "
            ElseIf Sheet8.opt_NEGMM1.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM2.Value = True Then
                strto = strto + Range("lpabapo").Value + "; "
            ElseIf Sheet8.opt_CCEGMM2.Value = True Then
                strcc = strcc + Range("lpabapo").Value + "; "
            ElseIf Sheet8.opt_NEGMM2.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM3.Value = True Then
                strto = strto + Range("raaquino").Value + "; "
            ElseIf Sheet8.opt_CCEGMM3.Value = True Then
                strcc = strcc + Range("raaquino").Value + "; "
            ElseIf Sheet8.opt_NEGMM3.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM4.Value = True Then
                strto = strto + Range("rrcerdon").Value + "; "
            ElseIf Sheet8.opt_CCEGMM4.Value = True Then
                strcc = strcc + Range("rrcerdon").Value + "; "
            ElseIf Sheet8.opt_NEGMM4.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM5.Value = True Then
                strto = strto + Range("lmcoma").Value + "; "
            ElseIf Sheet8.opt_CCEGMM5.Value = True Then
                strcc = strcc + Range("lmcoma").Value + "; "
            ElseIf Sheet8.opt_NEGMM5.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM6.Value = True Then
                strto = strto + Range("radeguzman").Value + "; "
            ElseIf Sheet8.opt_CCEGMM6.Value = True Then
                strcc = strcc + Range("radeguzman").Value + "; "
            ElseIf Sheet8.opt_NEGMM6.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM7.Value = True Then
                strto = strto + Range("jugabriel").Value + "; "
            ElseIf Sheet8.opt_CCEGMM7.Value = True Then
                strcc = strcc + Range("jugabriel").Value + "; "
            ElseIf Sheet8.opt_NEGMM7.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM8.Value = True Then
                strto = strto + Range("mbgaspay").Value + "; "
            ElseIf Sheet8.opt_CCEGMM8.Value = True Then
                strcc = strcc + Range("mbgaspay").Value + "; "
            ElseIf Sheet8.opt_NEGMM8.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM9.Value = True Then
                strto = strto + Range("cgkatalbas").Value + "; "
            ElseIf Sheet8.opt_CCEGMM9.Value = True Then
                strcc = strcc + Range("cgkatalbas").Value + "; "
            ElseIf Sheet8.opt_NEGMM9.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM10.Value = True Then
                strto = strto + Range("nrlapus").Value + "; "
            ElseIf Sheet8.opt_CCEGMM10.Value = True Then
                strcc = strcc + Range("nrlapus").Value + "; "
            ElseIf Sheet8.opt_NEGMM10.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM11.Value = True Then
                strto = strto + Range("fslizardo").Value + "; "
            ElseIf Sheet8.opt_CCEGMM11.Value = True Then
                strcc = strcc + Range("fslizardo").Value + "; "
            ElseIf Sheet8.opt_NEGMM11.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM12.Value = True Then
                strto = strto + Range("drparrocha").Value + "; "
            ElseIf Sheet8.opt_CCEGMM12.Value = True Then
                strcc = strcc + Range("drparrocha").Value + "; "
            ElseIf Sheet8.opt_NEGMM12.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM13.Value = True Then
                strto = strto + Range("nsreas").Value + "; "
            ElseIf Sheet8.opt_CCEGMM13.Value = True Then
                strcc = strcc + Range("nsreas").Value + "; "
            ElseIf Sheet8.opt_NEGMM13.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM14.Value = True Then
                strto = strto + Range("jdsarinas").Value + "; "
            ElseIf Sheet8.opt_CCEGMM14.Value = True Then
                strcc = strcc + Range("jdsarinas").Value + "; "
            ElseIf Sheet8.opt_NEGMM14.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM15.Value = True Then
                strto = strto + Range("rjteves").Value + "; "
            ElseIf Sheet8.opt_CCEGMM15.Value = True Then
                strcc = strcc + Range("rjteves").Value + "; "
            ElseIf Sheet8.opt_NEGMM15.Value = True Then
            End If
            
            For i = 1 To Notif.lst_EGMMGHL.ListCount - 1
                s = Notif.lst_EGMMGHL.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B25").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'FCNO GHL
        ElseIf Notif.MultiPage6.Value = 1 And Notif.MultiPage6.Pages(1).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B26").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToFCNO.Value = True Then
                strto = strto + Range("afcapiral").Value + "; "
            ElseIf Sheet8.opt_CCFCNO.Value = True Then
                strcc = strcc + Range("afcapiral").Value + "; "
            ElseIf Sheet8.opt_NFCNO.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO1.Value = True Then
                strto = strto + Range("picables").Value + "; "
            ElseIf Sheet8.opt_CCFCNO1.Value = True Then
                strcc = strcc + Range("picables").Value + "; "
            ElseIf Sheet8.opt_NFCNO1.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO2.Value = True Then
                strto = strto + Range("fcbaul").Value + "; "
            ElseIf Sheet8.opt_CCFCNO2.Value = True Then
                strcc = strcc + Range("fcbaul").Value + "; "
            ElseIf Sheet8.opt_NFCNO2.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO3.Value = True Then
                strto = strto + Range("aljimenez").Value + "; "
            ElseIf Sheet8.opt_CCFCNO3.Value = True Then
                strcc = strcc + Range("aljimenez").Value + "; "
            ElseIf Sheet8.opt_NFCNO3.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO4.Value = True Then
                strto = strto + Range("rtlampa").Value + "; "
            ElseIf Sheet8.opt_CCFCNO4.Value = True Then
                strcc = strcc + Range("rtlampa").Value + "; "
            ElseIf Sheet8.opt_NFCNO4.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO5.Value = True Then
                strto = strto + Range("oalinco").Value + "; "
            ElseIf Sheet8.opt_CCFCNO5.Value = True Then
                strcc = strcc + Range("oalinco").Value + "; "
            ElseIf Sheet8.opt_NFCNO5.Value = True Then
            End If
            
            If Sheet8.opt_ToFCNO6.Value = True Then
                strto = strto + Range("mfsantos").Value + "; "
            ElseIf Sheet8.opt_CCFCNO6.Value = True Then
                strcc = strcc + Range("mfsantos").Value + "; "
            ElseIf Sheet8.opt_NFCNO6.Value = True Then
            End If
            
            For i = 1 To Notif.lst_FCNOGH.ListCount - 1
                s = Notif.lst_FCNOGH.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B26").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
        End If
    End If
    
    Call View2(strbody, strtable, strto, strcc, s, signature)
    
End Sub

Sub View2(strbody As String, strtable As String, strto As String, strcc As String, s As String, signature As String)
    Dim OutApp As Object
    Dim OutMail As Object
    
    'GARNET
    If Notif.MultiPage1.Value = 5 Then
    'EGMMFxATOp GNT
        If Notif.MultiPage8.Value = 0 And Notif.MultiPage8.Pages(0).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B27").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToEGMM.Value = True Then
                strto = strto + Range("ntalcantara").Value + "; "
            ElseIf Sheet8.opt_CCEGMM.Value = True Then
                strcc = strcc + Range("ntalcantara").Value + "; "
            ElseIf Sheet8.opt_NEGMM.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM1.Value = True Then
                strto = strto + Range("armadrelino").Value + "; "
            ElseIf Sheet8.opt_CCEGMM1.Value = True Then
                strcc = strcc + Range("armadrelino").Value + "; "
            ElseIf Sheet8.opt_NEGMM1.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM2.Value = True Then
                strto = strto + Range("lpabapo").Value + "; "
            ElseIf Sheet8.opt_CCEGMM2.Value = True Then
                strcc = strcc + Range("lpabapo").Value + "; "
            ElseIf Sheet8.opt_NEGMM2.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM3.Value = True Then
                strto = strto + Range("raaquino").Value + "; "
            ElseIf Sheet8.opt_CCEGMM3.Value = True Then
                strcc = strcc + Range("raaquino").Value + "; "
            ElseIf Sheet8.opt_NEGMM3.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM4.Value = True Then
                strto = strto + Range("rrcerdon").Value + "; "
            ElseIf Sheet8.opt_CCEGMM4.Value = True Then
                strcc = strcc + Range("rrcerdon").Value + "; "
            ElseIf Sheet8.opt_NEGMM4.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM5.Value = True Then
                strto = strto + Range("lmcoma").Value + "; "
            ElseIf Sheet8.opt_CCEGMM5.Value = True Then
                strcc = strcc + Range("lmcoma").Value + "; "
            ElseIf Sheet8.opt_NEGMM5.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM6.Value = True Then
                strto = strto + Range("radeguzman").Value + "; "
            ElseIf Sheet8.opt_CCEGMM6.Value = True Then
                strcc = strcc + Range("radeguzman").Value + "; "
            ElseIf Sheet8.opt_NEGMM6.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM7.Value = True Then
                strto = strto + Range("jugabriel").Value + "; "
            ElseIf Sheet8.opt_CCEGMM7.Value = True Then
                strcc = strcc + Range("jugabriel").Value + "; "
            ElseIf Sheet8.opt_NEGMM7.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM8.Value = True Then
                strto = strto + Range("mbgaspay").Value + "; "
            ElseIf Sheet8.opt_CCEGMM8.Value = True Then
                strcc = strcc + Range("mbgaspay").Value + "; "
            ElseIf Sheet8.opt_NEGMM8.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM9.Value = True Then
                strto = strto + Range("cgkatalbas").Value + "; "
            ElseIf Sheet8.opt_CCEGMM9.Value = True Then
                strcc = strcc + Range("cgkatalbas").Value + "; "
            ElseIf Sheet8.opt_NEGMM9.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM10.Value = True Then
                strto = strto + Range("nrlapus").Value + "; "
            ElseIf Sheet8.opt_CCEGMM10.Value = True Then
                strcc = strcc + Range("nrlapus").Value + "; "
            ElseIf Sheet8.opt_NEGMM10.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM11.Value = True Then
                strto = strto + Range("fslizardo").Value + "; "
            ElseIf Sheet8.opt_CCEGMM11.Value = True Then
                strcc = strcc + Range("fslizardo").Value + "; "
            ElseIf Sheet8.opt_NEGMM11.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM12.Value = True Then
                strto = strto + Range("drparrocha").Value + "; "
            ElseIf Sheet8.opt_CCEGMM12.Value = True Then
                strcc = strcc + Range("drparrocha").Value + "; "
            ElseIf Sheet8.opt_NEGMM12.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM13.Value = True Then
                strto = strto + Range("nsreas").Value + "; "
            ElseIf Sheet8.opt_CCEGMM13.Value = True Then
                strcc = strcc + Range("nsreas").Value + "; "
            ElseIf Sheet8.opt_NEGMM13.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM14.Value = True Then
                strto = strto + Range("jdsarinas").Value + "; "
            ElseIf Sheet8.opt_CCEGMM14.Value = True Then
                strcc = strcc + Range("jdsarinas").Value + "; "
            ElseIf Sheet8.opt_NEGMM14.Value = True Then
            End If
            
            If Sheet8.opt_ToEGMM15.Value = True Then
                strto = strto + Range("rjteves").Value + "; "
            ElseIf Sheet8.opt_CCEGMM15.Value = True Then
                strcc = strcc + Range("rjteves").Value + "; "
            ElseIf Sheet8.opt_NEGMM15.Value = True Then
            End If
            
            
            For i = 1 To Notif.lst_EGMMGNT.ListCount - 1
                s = Notif.lst_EGMMGNT.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B27").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'FCNO FF1 GNT
        ElseIf Notif.MultiPage8.Value = 1 And Notif.MultiPage8.Pages(1).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B28").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToFCNO.Value = True Then
                strto = strto + Range("afcapiral").Value + "; "
            ElseIf Sheet8.opt_CCFCNO.Value = True Then
                strcc = strcc + Range("afcapiral").Value + "; "
            ElseIf Sheet8.opt_NFCNO.Value = True Then
            End If
            
            If Sheet8.opt_ToGARNET1.Value = True Then
                strto = strto + Range("eanieva").Value + "; "
            ElseIf Sheet8.opt_CCGARNET1.Value = True Then
                strcc = strcc + Range("eanieva").Value + "; "
            ElseIf Sheet8.opt_NGARNET1.Value = True Then
            End If
            
            If Sheet8.opt_ToGARNET2.Value = True Then
                strto = strto + Range("aeenano").Value + "; "
            ElseIf Sheet8.opt_CCGARNET2.Value = True Then
                strcc = strcc + Range("aeenano").Value + "; "
            ElseIf Sheet8.opt_NGARNET2.Value = True Then
            End If
            
            If Sheet8.opt_ToGARNET3.Value = True Then
                strto = strto + Range("apinciong").Value + "; "
            ElseIf Sheet8.opt_CCGARNET3.Value = True Then
                strcc = strcc + Range("apinciong").Value + "; "
            ElseIf Sheet8.opt_NGARNET3.Value = True Then
            End If
            
            If Sheet8.opt_ToGARNET4.Value = True Then
                strto = strto + Range("rcroque").Value + "; "
            ElseIf Sheet8.opt_CCGARNET4.Value = True Then
                strcc = strcc + Range("rcroque").Value + "; "
            ElseIf Sheet8.opt_NGARNET4.Value = True Then
            End If
            
            If Sheet8.opt_ToGARNET5.Value = True Then
                strto = strto + Range("mosena").Value + "; "
            ElseIf Sheet8.opt_CCGARNET5.Value = True Then
                strcc = strcc + Range("mosena").Value + "; "
            ElseIf Sheet8.opt_NGARNET5.Value = True Then
            End If
            
            If Sheet8.opt_ToGARNET6.Value = True Then
                strto = strto + Range("jutamondong").Value + "; "
            ElseIf Sheet8.opt_CCGARNET6.Value = True Then
                strcc = strcc + Range("jutamondong").Value + "; "
            ElseIf Sheet8.opt_NGARNET6.Value = True Then
            End If
            
            For i = 1 To Notif.lst_FCNOGNT.ListCount - 1
                s = Notif.lst_FCNOGNT.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B28").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
        End If
    End If
    
    'CEBU
    If Notif.MultiPage1.Value = 6 Then
    'VisFxATOp
        If Notif.MultiPage9.Value = 0 And Notif.MultiPage9.Pages(0).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B29").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToVis.Value = True Then
                strto = strto + Range("rasalas").Value + "; "
            ElseIf Sheet8.opt_CCVis.Value = True Then
                strcc = strcc + Range("rasalas").Value + "; "
            ElseIf Sheet8.opt_NVis.Value = True Then
            End If
            
            If Sheet8.opt_ToVis1.Value = True Then
                strto = strto + Range("mpoplas").Value + "; "
            ElseIf Sheet8.opt_CCVis1.Value = True Then
                strcc = strcc + Range("mpoplas").Value + "; "
            ElseIf Sheet8.opt_NVis1.Value = True Then
            End If
            
            If Sheet8.opt_ToVis2.Value = True Then
                strto = strto + Range("emallosada").Value + "; "
            ElseIf Sheet8.opt_CCVis2.Value = True Then
                strcc = strcc + Range("emallosada").Value + "; "
            ElseIf Sheet8.opt_NVis2.Value = True Then
            End If
            
            If Sheet8.opt_ToVis3.Value = True Then
                strto = strto + Range("rbarias").Value + "; "
            ElseIf Sheet8.opt_CCVis3.Value = True Then
                strcc = strcc + Range("rbarias").Value + "; "
            ElseIf Sheet8.opt_NVis3.Value = True Then
            End If
            
            If Sheet8.opt_ToVis4.Value = True Then
                strto = strto + Range("aabaguio").Value + "; "
            ElseIf Sheet8.opt_CCVis4.Value = True Then
                strcc = strcc + Range("aabaguio").Value + "; "
            ElseIf Sheet8.opt_NVis4.Value = True Then
            End If
            
            If Sheet8.opt_ToVis5.Value = True Then
                strto = strto + Range("lacabigas").Value + "; "
            ElseIf Sheet8.opt_CCVis5.Value = True Then
                strcc = strcc + Range("lacabigas").Value + "; "
            ElseIf Sheet8.opt_NVis5.Value = True Then
            End If
            
            If Sheet8.opt_ToVis6.Value = True Then
                strto = strto + Range("jjcarumba").Value + "; "
            ElseIf Sheet8.opt_CCVis6.Value = True Then
                strcc = strcc + Range("jjcarumba").Value + "; "
            ElseIf Sheet8.opt_NVis6.Value = True Then
            End If
            
            If Sheet8.opt_ToVis7.Value = True Then
                strto = strto + Range("rgconchas").Value + "; "
            ElseIf Sheet8.opt_CCVis7.Value = True Then
                strcc = strcc + Range("rgconchas").Value + "; "
            ElseIf Sheet8.opt_NVis7.Value = True Then
            End If
            
            If Sheet8.opt_ToVis8.Value = True Then
                strto = strto + Range("vbcuevas").Value + "; "
            ElseIf Sheet8.opt_CCVis8.Value = True Then
                strcc = strcc + Range("vbcuevas").Value + "; "
            ElseIf Sheet8.opt_NVis8.Value = True Then
            End If
            
            If Sheet8.opt_ToVis9.Value = True Then
                strto = strto + Range("jbdesamparado").Value + "; "
            ElseIf Sheet8.opt_CCVis9.Value = True Then
                strcc = strcc + Range("jbdesamparado").Value + "; "
            ElseIf Sheet8.opt_NVis9.Value = True Then
            End If
            
            If Sheet8.opt_ToVis10.Value = True Then
                strto = strto + Range("rddesquitado").Value + "; "
            ElseIf Sheet8.opt_CCVis10.Value = True Then
                strcc = strcc + Range("rddesquitado").Value + "; "
            ElseIf Sheet8.opt_NVis10.Value = True Then
            End If
            
            If Sheet8.opt_ToVis11.Value = True Then
                strto = strto + Range("mmdevero").Value + "; "
            ElseIf Sheet8.opt_CCVis11.Value = True Then
                strcc = strcc + Range("mmdevero").Value + "; "
            ElseIf Sheet8.opt_NVis11.Value = True Then
            End If
            
            If Sheet8.opt_ToVis12.Value = True Then
                strto = strto + Range("jcfelisarta").Value + "; "
            ElseIf Sheet8.opt_CCVis12.Value = True Then
                strcc = strcc + Range("jcfelisarta").Value + "; "
            ElseIf Sheet8.opt_NVis12.Value = True Then
            End If
            
            If Sheet8.opt_ToVis13.Value = True Then
                strto = strto + Range("rdflores").Value + "; "
            ElseIf Sheet8.opt_CCVis13.Value = True Then
                strcc = strcc + Range("rdflores").Value + "; "
            ElseIf Sheet8.opt_NVis13.Value = True Then
            End If
            
            If Sheet8.opt_ToVis14.Value = True Then
                strto = strto + Range("gpintes").Value + "; "
            ElseIf Sheet8.opt_CCVis14.Value = True Then
                strcc = strcc + Range("gpintes").Value + "; "
            ElseIf Sheet8.opt_NVis14.Value = True Then
            End If
            
            If Sheet8.opt_ToVis15.Value = True Then
                strto = strto + Range("dsisleta").Value + "; "
            ElseIf Sheet8.opt_CCVis15.Value = True Then
                strcc = strcc + Range("dsisleta").Value + "; "
            ElseIf Sheet8.opt_NVis15.Value = True Then
            End If
            
            If Sheet8.opt_ToVis16.Value = True Then
                strto = strto + Range("rmlocaylocay").Value + "; "
            ElseIf Sheet8.opt_CCVis16.Value = True Then
                strcc = strcc + Range("rmlocaylocay").Value + "; "
            ElseIf Sheet8.opt_NVis16.Value = True Then
            End If
            
            If Sheet8.opt_ToVis17.Value = True Then
                strto = strto + Range("mmnadal").Value + "; "
            ElseIf Sheet8.opt_CCVis17.Value = True Then
                strcc = strcc + Range("mmnadal").Value + "; "
            ElseIf Sheet8.opt_NVis17.Value = True Then
            End If
            
            If Sheet8.opt_ToVis18.Value = True Then
                strto = strto + Range("npompad").Value + "; "
            ElseIf Sheet8.opt_CCVis18.Value = True Then
                strcc = strcc + Range("npompad").Value + "; "
            ElseIf Sheet8.opt_NVis18.Value = True Then
            End If
            
            If Sheet8.opt_ToVis19.Value = True Then
                strto = strto + Range("dbpepito").Value + "; "
            ElseIf Sheet8.opt_CCVis19.Value = True Then
                strcc = strcc + Range("dbpepito").Value + "; "
            ElseIf Sheet8.opt_NVis19.Value = True Then
            End If
            
            If Sheet8.opt_ToVis20.Value = True Then
                strto = strto + Range("izpono").Value + "; "
            ElseIf Sheet8.opt_CCVis20.Value = True Then
                strcc = strcc + Range("izpono").Value + "; "
            ElseIf Sheet8.opt_NVis20.Value = True Then
            End If
            
            If Sheet8.opt_ToVis21.Value = True Then
                strto = strto + Range("clrosales").Value + "; "
            ElseIf Sheet8.opt_CCVis21.Value = True Then
                strcc = strcc + Range("clrosales").Value + "; "
            ElseIf Sheet8.opt_NVis21.Value = True Then
            End If
            
            If Sheet8.opt_ToVis22.Value = True Then
                strto = strto + Range("mdsarcauga").Value + "; "
            ElseIf Sheet8.opt_CCVis22.Value = True Then
                strcc = strcc + Range("mdsarcauga").Value + "; "
            ElseIf Sheet8.opt_NVis22.Value = True Then
            End If
            
            If Sheet8.opt_ToVis23.Value = True Then
                strto = strto + Range("mlsarmiento").Value + "; "
            ElseIf Sheet8.opt_CCVis23.Value = True Then
                strcc = strcc + Range("mlsarmiento").Value + "; "
            ElseIf Sheet8.opt_NVis23.Value = True Then
            End If
            
            If Sheet8.opt_ToVis24.Value = True Then
                strto = strto + Range("vjson").Value + "; "
            ElseIf Sheet8.opt_CCVis24.Value = True Then
                strcc = strcc + Range("vjson").Value + "; "
            ElseIf Sheet8.opt_NVis24.Value = True Then
            End If
            
            If Sheet8.opt_ToVis25.Value = True Then
                strto = strto + Range("jltacan").Value + "; "
            ElseIf Sheet8.opt_CCVis25.Value = True Then
                strcc = strcc + Range("jltacan").Value + "; "
            ElseIf Sheet8.opt_NVis25.Value = True Then
            End If
            
            If Sheet8.opt_ToVis26.Value = True Then
                strto = strto + Range("fotamarra").Value + "; "
            ElseIf Sheet8.opt_CCVis26.Value = True Then
                strcc = strcc + Range("fotamarra").Value + "; "
            ElseIf Sheet8.opt_NVis26.Value = True Then
            End If
            
            If Sheet8.opt_ToVis27.Value = True Then
                strto = strto + Range("ertejano").Value + "; "
            ElseIf Sheet8.opt_CCVis27.Value = True Then
                strcc = strcc + Range("ertejano").Value + "; "
            ElseIf Sheet8.opt_NVis27.Value = True Then
            End If
            
            If Sheet8.opt_ToVis28.Value = True Then
                strto = strto + Range("blynot").Value + "; "
            ElseIf Sheet8.opt_CCVis28.Value = True Then
                strcc = strcc + Range("blynot").Value + "; "
            ElseIf Sheet8.opt_NVis28.Value = True Then
            End If
            
            For i = 1 To Notif.lst_VisFxATOp.ListCount - 1
                s = Notif.lst_VisFxATOp.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B29").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'FCNO FF5 JNE
        ElseIf Notif.MultiPage9.Value = 1 And Notif.MultiPage9.Pages(1).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B30").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToFCNO.Value = True Then
                strto = strto + Range("afcapiral").Value + "; "
            ElseIf Sheet8.opt_CCFCNO.Value = True Then
                strcc = strcc + Range("afcapiral").Value + "; "
            ElseIf Sheet8.opt_NFCNO.Value = True Then
            End If
            
            If Sheet8.opt_ToCEBU1.Value = True Then
                strto = strto + Range("eddivinagracia").Value + "; "
            ElseIf Sheet8.opt_CCCEBU1.Value = True Then
                strcc = strcc + Range("eddivinagracia").Value + "; "
            ElseIf Sheet8.opt_NCEBU1.Value = True Then
            End If
            
            If Sheet8.opt_ToCEBU2.Value = True Then
                strto = strto + Range("rngloria").Value + "; "
            ElseIf Sheet8.opt_CCCEBU2.Value = True Then
                strcc = strcc + Range("rngloria").Value + "; "
            ElseIf Sheet8.opt_NCEBU2.Value = True Then
            End If
            
            If Sheet8.opt_ToCEBU3.Value = True Then
                strto = strto + Range("cminoferio").Value + "; "
            ElseIf Sheet8.opt_CCCEBU3.Value = True Then
                strcc = strcc + Range("cminoferio").Value + "; "
            ElseIf Sheet8.opt_NCEBU3.Value = True Then
            End If
            
            If Sheet8.opt_ToCEBU4.Value = True Then
                strto = strto + Range("jrmaninang").Value + "; "
            ElseIf Sheet8.opt_CCCEBU4.Value = True Then
                strcc = strcc + Range("jrmaninang").Value + "; "
            ElseIf Sheet8.opt_NCEBU4.Value = True Then
            End If
            
            If Sheet8.opt_ToCEBU5.Value = True Then
                strto = strto + Range("rrson").Value + "; "
            ElseIf Sheet8.opt_CCCEBU5.Value = True Then
                strcc = strcc + Range("rrson").Value + "; "
            ElseIf Sheet8.opt_NCEBU5.Value = True Then
            End If
            
            If Sheet8.opt_ToCEBU6.Value = True Then
                strto = strto + Range("mjvendiola").Value + "; "
            ElseIf Sheet8.opt_CCCEBU6.Value = True Then
                strcc = strcc + Range("mjvendiola").Value + "; "
            ElseIf Sheet8.opt_NCEBU6.Value = True Then
            End If
            
            For i = 1 To Notif.lst_FCNOJNE.ListCount - 1
                s = Notif.lst_FCNOJNE.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>F" & Sheet1.Range("B30").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
        End If
    End If
    
    'CONSOLIDATED
    If Notif.MultiPage1.Value = 7 Then
    'STRAT
        If Notif.MultiPage7.Value = 0 And Notif.MultiPage7.Pages(0).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B14").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToSTRAT.Value = True Then
                strto = strto + Range("aiseeco").Value + "; "
            ElseIf Sheet8.opt_CCSTRAT.Value = True Then
                strcc = strcc + Range("aiseeco").Value + "; "
            ElseIf Sheet8.opt_NSTRAT.Value = True Then
            End If
            
            If Sheet8.opt_ToSTRAT1.Value = True Then
                strto = strto + Range("bccordova").Value + "; "
            ElseIf Sheet8.opt_CCSTRAT1.Value = True Then
                strcc = strcc + Range("bccordova").Value + "; "
            ElseIf Sheet8.opt_NSTRAT1.Value = True Then
            End If
            
            If Sheet8.opt_ToSTRAT2.Value = True Then
                strto = strto + Range("jpjandayan").Value + "; "
            ElseIf Sheet8.opt_CCSTRAT2.Value = True Then
                strcc = strcc + Range("jpjandayan").Value + "; "
            ElseIf Sheet8.opt_NSTRAT2.Value = True Then
            End If
            
            For i = 1 To Notif.lst_STRATC.ListCount - 1
                s = Notif.lst_STRATC.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B14").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
    'ENG
        ElseIf Notif.MultiPage7.Value = 1 And Notif.MultiPage7.Pages(1).Caption Like "*" Then
            strbody = "Dear " & Sheet1.Range("B15").Text & "," & vbCrLf & vbCrLf & "Please be informed that the following BCMS documents are already near expiration date. Please fill up the highlighted columns in the table below and make necessary updates, as applicable:"
            
            If Sheet8.opt_ToENG.Value = True Then
                strto = strto + Range("megutierrez").Value + "; "
            ElseIf Sheet8.opt_CCENG.Value = True Then
                strcc = strcc + Range("megutierrez").Value + "; "
            ElseIf Sheet8.opt_NENG.Value = True Then
            End If
            
            If Sheet8.opt_ToENG1.Value = True Then
                strto = strto + Range("amarsua").Value + "; "
            ElseIf Sheet8.opt_CCENG1.Value = True Then
                strcc = strcc + Range("amarsua").Value + "; "
            ElseIf Sheet8.opt_NENG1.Value = True Then
            End If
            
            If Sheet8.opt_ToENG2.Value = True Then
                strto = strto + Range("istolentino").Value + "; "
            ElseIf Sheet8.opt_CCENG2.Value = True Then
                strcc = strcc + Range("istolentino").Value + "; "
            ElseIf Sheet8.opt_NENG2.Value = True Then
            End If
            
            For i = 1 To Notif.lst_ENGC.ListCount - 1
                s = Notif.lst_ENGC.List(i)
                If s Like "*3 Months Before*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:orange'>" & s & "</td></tr>"
                ElseIf s Like "*For Urgent Update*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:red;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*Outdated*" Then
                    strtable = strtable & "<tr><td colspan='7' style='background-color:black;color:white'>" & s & "</td></tr>"
                ElseIf s Like "*REV*" Then
                    strtable = strtable & "<tr><td>" & Sheet1.Range("B15").Text _
                        & "</td> <td>" & Split(s, "    ")(4) _
                        & "</td> <td>" & Split(s, "    ")(1) _
                        & "</td> <td>" & Split(s, "    ")(2) _
                        & "</td> <td style='background-color:yellow'></td> <td style='background-color:yellow'></td> <td style='background-color: yellow'></td></tr>"
                End If
            Next i
        End If
    End If
    
    If Not strto Like "*bccordova*" And strtable Like "*BC Scope (Specific for BU)*" _
        Or strtable Like "*BC Objectives*" _
        Or strtable Like "*Identification of Organization and its Context (4.1)*" _
        Or strtable Like "*Identification of Interested Parties and Their Needs (4.2)*" _
        Or strtable Like "*Business Impact Analysis Summary Report*" _
        Or strtable Like "*BIA Questionnaires*" _
        Or strtable Like "*Risk Assesssment*" _
        Or strtable Like "*BC Strategy*" Then
        If Sheet8.opt_ToSTRAT1.Value = True Or Sheet8.opt_CCSTRAT1.Value = True Then
            strcc = strcc + Range("bccordova").Value + "; "
        ElseIf Sheet8.opt_NSTRAT1.Value = True Then
        End If
    End If
    
    If Not strto Like "*jpjandayan*" And strtable Like "*BC Scope (Specific for BU)*" _
        Or strtable Like "*BC Objectives*" _
        Or strtable Like "*Identification of Organization and its Context (4.1)*" _
        Or strtable Like "*Identification of Interested Parties and Their Needs (4.2)*" _
        Or strtable Like "*Business Impact Analysis Summary Report*" _
        Or strtable Like "*BIA Questionnaires*" _
        Or strtable Like "*Risk Assesssment*" _
        Or strtable Like "*BC Strategy*" Then
        If Sheet8.opt_ToSTRAT2.Value = True Or Sheet8.opt_CCSTRAT2.Value = True Then
            strcc = strcc + Range("jpjandayan").Value + "; "
        ElseIf Sheet8.opt_NSTRAT2.Value = True Then
        End If
    End If
    
    If Not strto Like "*amarsua*" And strtable Like "*BC Plan-Business Unit*" Then
        If Sheet8.opt_ToENG1.Value = True Or Sheet8.opt_CCENG1.Value = True Then
            strcc = strcc + Range("amarsua").Value + "; "
        ElseIf Sheet8.opt_NENG1.Value = True Then
        End If
    End If
    
    If Not strto Like "*istolentino*" And strtable Like "*BC Plan-Business Unit*" Then
        If Sheet8.opt_ToENG2.Value = True Or Sheet8.opt_CCENG2.Value = True Then
            strcc = strcc + Range("istolentino").Value + "; "
        ElseIf Sheet8.opt_NENG2.Value = True Then
        End If
    End If
    
    If Sheet8.opt_ToGOV1.Value = True Then
        strto = strto + Range("janacpil").Value + "; "
    ElseIf Sheet8.opt_CCGOV1.Value = True Then
        strcc = strcc + Range("janacpil").Value + "; "
    ElseIf Sheet8.opt_NGOV1.Value = True Then
    End If
    
    If Not strto Like "*hclim*" Then
        If Sheet8.opt_ToGOV2.Value = True Then
            strto = strto + Range("hclim").Value + "; "
        ElseIf Sheet8.opt_CCGOV2.Value = True Then
            strcc = strcc + Range("hclim").Value + "; "
        ElseIf Sheet8.opt_NGOV2.Value = True Then
        End If
    End If
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error Resume Next
    
    signature = Environ("appdata") & "\Microsoft\Signatures\"
    If Dir(signature, vbDirectory) <> vbNullString Then
        signature = signature & Dir$(signature & "*.htm")
    Else:
        signature = ""
    End If
    signature = CreateObject("Scripting.FileSystemObject").GetFile(signature).OpenAsTextStream(1, -2).ReadAll
    OutMail.HTMLBody = signature
    
    With OutMail
        .To = strto
        .CC = strcc
        .BCC = ""
        .Subject = "BCMS Documents Update"
        .Body = strbody & vbCrLf & Range("AddNote")
        .Attachments.Add (Range("FileSource").Value)
        .HTMLBody = .HTMLBody & "<table border='1' cellspacing='0' cellpadding='0' style='width:100%; border:1px solid black; border-collapse:collapse; text-align:center; font-family:calibri; font-size:14.5px'>" _
            & "<tbody><tr><td rowspan='2' ><p><b>BUSINESS UNIT/Document Owner</b></p></td>" _
            & "<td rowspan='2' ><p><b>DOCUMENT TITLE</b></p></td>" _
            & "<td rowspan='2' ><p><b>EFFECTIVITY DATE</b></p></td>" _
            & "<td rowspan='2' ><p><b>EXPIRATION DATE</b></p></td>" _
            & "<td colspan='2' style='background-color: yellow'><p><b>NEED UPDATE?</b></p></td>" _
            & "<td rowspan='2' style='background-color: yellow'><p><b>REMARKS</b><br><i style='font-size:13.5'>(e.g. reason why updating is not needed)</i></p></td></tr>" _
            & "<tr><td valign='top' style='background-color: yellow'><p><i>YES</i></p></td>" _
            & "<td valign='top' style='background-color: yellow'><p><i>NO</i></p></td></tr>" _
            & strtable & "</tbody></table>" & "<br>Thank You.<br>Regards,<br>" & signature
        .Display
    End With
    On Error GoTo 0
    
    Set OutMail = Nothing
    Set OutApp = Nothing
    MsgBox "Email Sent", vbInformation, "Success"
End Sub
