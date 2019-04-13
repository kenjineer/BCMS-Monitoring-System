VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Status 
   Caption         =   "Status"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7425
   OleObjectBlob   =   "Status.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+'
'    Title: BC Gov IA Calendar Automatic Email and Deadline Notifier     '
'    Author: Engr. Kenneth Caro Karamihan                                '
'    Company: PLDT Inc.                                                  '
'    Division: BC Governance and Reporting                               '
'    Date: May 7, 2018                                                   '
'    Code version: 1.0                                                   '
'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+'

Private Sub btn_Send_Click()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Dim strto As String
    Dim strcc As String
    Dim i As Integer
    
    If Sheet2.opt_To2.Value = True Then
        strto = strto + Sheet2.Range("rdvillasenor").Value + "; "
    ElseIf Sheet2.opt_CC2.Value = True Then
        strcc = strcc + Sheet2.Range("rdvillasenor").Value + "; "
    ElseIf Sheet2.opt_N2.Value = True Then
    End If
    
    If Sheet2.opt_To3.Value = True Then
        strto = strto + Sheet2.Range("aiseeco").Value + "; "
    ElseIf Sheet2.opt_CC3.Value = True Then
        strcc = strcc + Sheet2.Range("aiseeco").Value + "; "
    ElseIf Sheet2.opt_N3.Value = True Then
    End If
    
    If Sheet2.opt_To4.Value = True Then
        strto = strto + Sheet2.Range("megutierrez").Value + "; "
    ElseIf Sheet2.opt_CC4.Value = True Then
        strcc = strcc + Sheet2.Range("megutierrez").Value + "; "
    ElseIf Sheet2.opt_N4.Value = True Then
    End If
    
    If Sheet2.opt_To5.Value = True Then
        strto = strto + Sheet2.Range("janacpil").Value + "; "
    ElseIf Sheet2.opt_CC5.Value = True Then
        strcc = strcc + Sheet2.Range("janacpil").Value + "; "
    ElseIf Sheet2.opt_N5.Value = True Then
    End If
    
    If Sheet2.opt_To6.Value = True Then
        strto = strto + Sheet2.Range("hclim").Value + "; "
    ElseIf Sheet2.opt_CC6.Value = True Then
        strcc = strcc + Sheet2.Range("hclim").Value + "; "
    ElseIf Sheet2.opt_N6.Value = True Then
    End If
    
    If Sheet2.opt_To7.Value = True Then
        strto = strto + Sheet2.Range("bccordova").Value + "; "
    ElseIf Sheet2.opt_CC7.Value = True Then
        strcc = strcc + Sheet2.Range("bccordova").Value + "; "
    ElseIf Sheet2.opt_N7.Value = True Then
    End If
    
    If Sheet2.opt_To8.Value = True Then
        strto = strto + Sheet2.Range("jpjandayan").Value + "; "
    ElseIf Sheet2.opt_CC8.Value = True Then
        strcc = strcc + Sheet2.Range("jpjandayan").Value + "; "
    ElseIf Sheet2.opt_N8.Value = True Then
    End If
    
    If Sheet2.opt_To9.Value = True Then
        strto = strto + Sheet2.Range("amarsua").Value + "; "
    ElseIf Sheet2.opt_CC9.Value = True Then
        strcc = strcc + Sheet2.Range("amarsua").Value + "; "
    ElseIf Sheet2.opt_N9.Value = True Then
    End If
    
    If Sheet2.opt_To10.Value = True Then
        strto = strto + Sheet2.Range("istolentino").Value + "; "
    ElseIf Sheet2.opt_CC10.Value = True Then
        strcc = strcc + Sheet2.Range("istolentino").Value + "; "
    ElseIf Sheet2.opt_N10.Value = True Then
    End If

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    strbody = ""
    If Status.lst_L.ListCount >= 2 Then
        strbody = strbody + "LUCLS" + vbCrLf
        For i = 1 To Status.lst_L.ListCount - 1
           strbody = strbody + "   " + Status.lst_L.List(i) + vbCrLf
        Next i
    End If
    If Status.lst_B.ListCount >= 2 Then
        strbody = strbody + "BCLS" + vbCrLf
        For i = 1 To Status.lst_B.ListCount - 1
            strbody = strbody + "   " + Status.lst_B.List(i) + vbCrLf
        Next i
    End If
    If Status.lst_DC.ListCount >= 2 Then
        strbody = strbody + "DCLS" + vbCrLf
        For i = 1 To Status.lst_DC.ListCount - 1
            strbody = strbody + "   " + Status.lst_DC.List(i) + vbCrLf
        Next i
    End If
    If Status.lst_DL.ListCount >= 2 Then
        strbody = strbody + "DILIMAN" + vbCrLf
        For i = 1 To Status.lst_DL.ListCount - 1
            strbody = strbody + "   " + Status.lst_DL.List(i) + vbCrLf
        Next i
    End If
    If Status.lst_G.ListCount >= 2 Then
        strbody = strbody + "GARNET" + vbCrLf
        For i = 1 To Status.lst_G.ListCount - 1
            strbody = strbody + "   " + Status.lst_G.List(i) + vbCrLf
        Next i
    End If
    If Status.lst_S.ListCount >= 2 Then
        strbody = strbody + "SAMPALOC" + vbCrLf
        For i = 1 To Status.lst_S.ListCount - 1
            strbody = strbody + "   " + Status.lst_S.List(i) + vbCrLf
        Next i
    End If
    If Status.lst_GH.ListCount >= 2 Then
        strbody = strbody + "GREENHILLS" + vbCrLf
        For i = 1 To Status.lst_GH.ListCount - 1
            strbody = strbody + "   " + Status.lst_GH.List(i) + vbCrLf
        Next i
    End If
    If Status.lst_C.ListCount >= 2 Then
        strbody = strbody + "CEBU" + vbCrLf
        For i = 1 To Status.lst_C.ListCount - 1
            strbody = strbody + "   " + Status.lst_C.List(i) + vbCrLf
        Next i
    End If

    Dim signature As String
    
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
        .Subject = "BC Gov IA Deadline"
        .Body = strbody
        .Attachments.Add (Sheet3.Range("A11").Value)
        .Attachments.Add Sheet3.Range("A5").Value, olByValue, 0
        .HTMLBody = "<style>body{color:red;font-weight:bold}</style><h2>DUE DATE NEAR EXPIRATION:</h2>" & .HTMLBody & Sheet3.Range("A2").Value & "<br><br>" & "<img src='cid:" & Sheet3.Range("A8").Value & "'>" & signature
        .Send
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
    
    MsgBox "Email Sent!", vbOKOnly + vbInformation, "Alert"

End Sub

Private Sub btn_SendD_Click()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Dim strto As String
    Dim strcc As String
    Dim i As Integer
    
    If Sheet2.opt_To2.Value = True Then
        strto = strto + Sheet2.Range("rdvillasenor").Value + "; "
    ElseIf Sheet2.opt_CC2.Value = True Then
        strcc = strcc + Sheet2.Range("rdvillasenor").Value + "; "
    ElseIf Sheet2.opt_N2.Value = True Then
    End If
    
    If Sheet2.opt_To5.Value = True Then
        strto = strto + Sheet2.Range("janacpil").Value + "; "
    ElseIf Sheet2.opt_CC5.Value = True Then
        strcc = strcc + Sheet2.Range("janacpil").Value + "; "
    ElseIf Sheet2.opt_N5.Value = True Then
    End If
    
    If Sheet2.opt_To6.Value = True Then
        strto = strto + Sheet2.Range("hclim").Value + "; "
    ElseIf Sheet2.opt_CC6.Value = True Then
        strcc = strcc + Sheet2.Range("hclim").Value + "; "
    ElseIf Sheet2.opt_N6.Value = True Then
    End If
    
    If Status.MultiPage1.Value = 0 Then
        If Status.MultiPage1.Pages(0).Caption = "*LUCLS" Then
            If Sheet2.opt_ToL1.Value = True Then
                strto = strto + Sheet2.Range("mmginete").Value + "; "
            ElseIf Sheet2.opt_CCL1.Value = True Then
                strcc = strcc + Sheet2.Range("mmginete").Value + "; "
            ElseIf Sheet2.opt_NL1.Value = True Then
            End If
            
            If Sheet2.opt_ToL2.Value = True Then
                strto = strto + Sheet2.Range("jmjacinto").Value + "; "
            ElseIf Sheet2.opt_CCL2.Value = True Then
                strcc = strcc + Sheet2.Range("jmjacinto").Value + "; "
            ElseIf Sheet2.opt_NL2.Value = True Then
            End If
            
            If Sheet2.opt_ToL3.Value = True Then
                strto = strto + Sheet2.Range("vprodriguez").Value + "; "
            ElseIf Sheet2.opt_CCL3.Value = True Then
                strcc = strcc + Sheet2.Range("vprodriguez").Value + "; "
            ElseIf Sheet2.opt_NL3.Value = True Then
            End If
            
            If Sheet2.opt_ToL4.Value = True Then
                strto = strto + Sheet2.Range("ecganuelas").Value + "; "
            ElseIf Sheet2.opt_CCL4.Value = True Then
                strcc = strcc + Sheet2.Range("ecganuelas").Value + "; "
            ElseIf Sheet2.opt_NL4.Value = True Then
            End If
            
            If Sheet2.opt_ToL5.Value = True Then
                strto = strto + Sheet2.Range("mmmones").Value + "; "
            ElseIf Sheet2.opt_CCL5.Value = True Then
                strcc = strcc + Sheet2.Range("mmmones").Value + "; "
            ElseIf Sheet2.opt_NL5.Value = True Then
            End If
            
            If Sheet2.opt_ToL6.Value = True Then
                strto = strto + Sheet2.Range("ljechave").Value + "; "
            ElseIf Sheet2.opt_CCL6.Value = True Then
                strcc = strcc + Sheet2.Range("ljechave").Value + "; "
            ElseIf Sheet2.opt_NL6.Value = True Then
            End If
            
            If Sheet2.opt_ToL7.Value = True Then
                strto = strto + Sheet2.Range("ccmendoza").Value + "; "
            ElseIf Sheet2.opt_CCL7.Value = True Then
                strcc = strcc + Sheet2.Range("ccmendoza").Value + "; "
            ElseIf Sheet2.opt_NL7.Value = True Then
            End If
            
            If Sheet2.opt_ToL8.Value = True Then
                strto = strto + Sheet2.Range("jlmoldez").Value + "; "
            ElseIf Sheet2.opt_CCL8.Value = True Then
                strcc = strcc + Sheet2.Range("jlmoldez").Value + "; "
            ElseIf Sheet2.opt_NL8.Value = True Then
            End If
            
            If Sheet2.opt_ToL9.Value = True Then
                strto = strto + Sheet2.Range("jogutierrez").Value + "; "
            ElseIf Sheet2.opt_CCL9.Value = True Then
                strcc = strcc + Sheet2.Range("jogutierrez").Value + "; "
            ElseIf Sheet2.opt_NL9.Value = True Then
            End If
            
            If Sheet2.opt_ToCLS2.Value = True Then
                strto = strto + Sheet2.Range("emgacayan").Value + "; "
            ElseIf Sheet2.opt_CCCLS2.Value = True Then
                strcc = strcc + Sheet2.Range("emgacayan").Value + "; "
            ElseIf Sheet2.opt_NCLS2.Value = True Then
            End If
            
            strbody = ""
            strbody = strbody + "LUCLS" + vbCrLf
            For i = 1 To Status.lst_L.ListCount - 1
                If Status.lst_L.List(i) Like "*Off-site Documentation Audit*" Or Status.lst_L.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_L.List(i) + vbCrLf
                End If
            Next i
        End If
    End If
    
    If Status.MultiPage1.Value = 1 Then
        If Status.MultiPage1.Pages(1).Caption = "*BCLS" Then
            If Sheet2.opt_ToB1.Value = True Then
                strto = strto + Sheet2.Range("jdbustamante").Value + "; "
            ElseIf Sheet2.opt_CCB1.Value = True Then
                strcc = strcc + Sheet2.Range("jdbustamante").Value + "; "
            ElseIf Sheet2.opt_NB1.Value = True Then
            End If
            
            If Sheet2.opt_ToB2.Value = True Then
                strto = strto + Sheet2.Range("apcaringal").Value + "; "
            ElseIf Sheet2.opt_CCB2.Value = True Then
                strcc = strcc + Sheet2.Range("apcaringal").Value + "; "
            ElseIf Sheet2.opt_NB2.Value = True Then
            End If
            
            If Sheet2.opt_ToB3.Value = True Then
                strto = strto + Sheet2.Range("aacatibog").Value + "; "
            ElseIf Sheet2.opt_CCB3.Value = True Then
                strcc = strcc + Sheet2.Range("aacatibog").Value + "; "
            ElseIf Sheet2.opt_NB3.Value = True Then
            End If
            
            If Sheet2.opt_ToB4.Value = True Then
                strto = strto + Sheet2.Range("ebdeleon").Value + "; "
            ElseIf Sheet2.opt_CCB4.Value = True Then
                strcc = strcc + Sheet2.Range("ebdeleon").Value + "; "
            ElseIf Sheet2.opt_NB4.Value = True Then
            End If
            
            If Sheet2.opt_ToB5.Value = True Then
                strto = strto + Sheet2.Range("rjenriquez").Value + "; "
            ElseIf Sheet2.opt_CCB5.Value = True Then
                strcc = strcc + Sheet2.Range("rjenriquez").Value + "; "
            ElseIf Sheet2.opt_NB5.Value = True Then
            End If
            
            If Sheet2.opt_ToB6.Value = True Then
                strto = strto + Sheet2.Range("rjetcobanez").Value + "; "
            ElseIf Sheet2.opt_CCB6.Value = True Then
                strcc = strcc + Sheet2.Range("rjetcobanez").Value + "; "
            ElseIf Sheet2.opt_NB6.Value = True Then
            End If
            
            If Sheet2.opt_ToB7.Value = True Then
                strto = strto + Sheet2.Range("rpmanago").Value + "; "
            ElseIf Sheet2.opt_CCB7.Value = True Then
                strcc = strcc + Sheet2.Range("rpmanago").Value + "; "
            ElseIf Sheet2.opt_NB7.Value = True Then
            End If
            
            If Sheet2.opt_ToB8.Value = True Then
                strto = strto + Sheet2.Range("hdpilar").Value + "; "
            ElseIf Sheet2.opt_CCB8.Value = True Then
                strcc = strcc + Sheet2.Range("hdpilar").Value + "; "
            ElseIf Sheet2.opt_NB8.Value = True Then
            End If
            
            If Sheet2.opt_ToB9.Value = True Then
                strto = strto + Sheet2.Range("acreyes").Value + "; "
            ElseIf Sheet2.opt_CCB9.Value = True Then
                strcc = strcc + Sheet2.Range("acreyes").Value + "; "
            ElseIf Sheet2.opt_NB9.Value = True Then
            End If
            
            If Sheet2.opt_ToB10.Value = True Then
                strto = strto + Sheet2.Range("aevizconde").Value + "; "
            ElseIf Sheet2.opt_CCB10.Value = True Then
                strcc = strcc + Sheet2.Range("aevizconde").Value + "; "
            ElseIf Sheet2.opt_NB10.Value = True Then
            End If
            
            If Sheet2.opt_ToCLS2.Value = True Then
                strto = strto + Sheet2.Range("emgacayan").Value + "; "
            ElseIf Sheet2.opt_CCCLS2.Value = True Then
                strcc = strcc + Sheet2.Range("emgacayan").Value + "; "
            ElseIf Sheet2.opt_NCLS2.Value = True Then
            End If
            
            
            strbody = ""
            strbody = strbody + "BCLS" + vbCrLf
            For i = 1 To Status.lst_B.ListCount - 1
                If Status.lst_B.List(i) Like "*Off-site Documentation Audit*" Or Status.lst_B.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_B.List(i) + vbCrLf
                End If
            Next i
        End If
    End If
    
    If Status.MultiPage1.Value = 2 Then
        If Status.MultiPage1.Pages(2).Caption = "*DCLS" Then
            If Sheet2.opt_ToDC1.Value = True Then
                strto = strto + Sheet2.Range("vvpunzalan").Value + "; "
            ElseIf Sheet2.opt_CCDC1.Value = True Then
                strcc = strcc + Sheet2.Range("vvpunzalan").Value + "; "
            ElseIf Sheet2.opt_NDC1.Value = True Then
            End If
            
            If Sheet2.opt_ToDC2.Value = True Then
                strto = strto + Sheet2.Range("neromero").Value + "; "
            ElseIf Sheet2.opt_CCDC2.Value = True Then
                strcc = strcc + Sheet2.Range("neromero").Value + "; "
            ElseIf Sheet2.opt_NDC2.Value = True Then
            End If
            
            If Sheet2.opt_ToDC3.Value = True Then
                strto = strto + Sheet2.Range("dltatad").Value + "; "
            ElseIf Sheet2.opt_CCDC3.Value = True Then
                strcc = strcc + Sheet2.Range("dltatad").Value + "; "
            ElseIf Sheet2.opt_NDC3.Value = True Then
            End If
            
            If Sheet2.opt_ToDC4.Value = True Then
                strto = strto + Sheet2.Range("josalvador").Value + "; "
            ElseIf Sheet2.opt_CCDC4.Value = True Then
                strcc = strcc + Sheet2.Range("josalvador").Value + "; "
            ElseIf Sheet2.opt_NDC4.Value = True Then
            End If
            
            If Sheet2.opt_ToDC5.Value = True Then
                strto = strto + Sheet2.Range("mbdevilla").Value + "; "
            ElseIf Sheet2.opt_CCDC5.Value = True Then
                strcc = strcc + Sheet2.Range("mbdevilla").Value + "; "
            ElseIf Sheet2.opt_NDC5.Value = True Then
            End If
            
            If Sheet2.opt_ToDC6.Value = True Then
                strto = strto + Sheet2.Range("rspanta").Value + "; "
            ElseIf Sheet2.opt_CCDC6.Value = True Then
                strcc = strcc + Sheet2.Range("rspanta").Value + "; "
            ElseIf Sheet2.opt_NDC6.Value = True Then
            End If
            
            If Sheet2.opt_ToDC7.Value = True Then
                strto = strto + Sheet2.Range("jlandicoy").Value + "; "
            ElseIf Sheet2.opt_CCDC7.Value = True Then
                strcc = strcc + Sheet2.Range("jlandicoy").Value + "; "
            ElseIf Sheet2.opt_NDC7.Value = True Then
            End If
            
            If Sheet2.opt_ToCLS2.Value = True Then
                strto = strto + Sheet2.Range("emgacayan").Value + "; "
            ElseIf Sheet2.opt_CCCLS2.Value = True Then
                strcc = strcc + Sheet2.Range("emgacayan").Value + "; "
            ElseIf Sheet2.opt_NCLS2.Value = True Then
            End If
                      
            strbody = ""
            strbody = strbody + "DCLS" + vbCrLf
            For i = 1 To Status.lst_DC.ListCount - 1
                If Status.lst_DC.List(i) Like "*Off-site Documentation Audit*" Or Status.lst_DC.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_DC.List(i) + vbCrLf
                End If
            Next i
        End If
    End If
    
    If Status.MultiPage1.Value = 3 Then
        If Status.MultiPage1.Pages(3).Caption = "*DILIMAN" And chk_APFMDL.Value = False Then
            If Sheet2.opt_ToNGMM.Value = True Then
                strto = strto + Sheet2.Range("dcroque").Value + "; "
            ElseIf Sheet2.opt_CCNGMM.Value = True Then
                strcc = strcc + Sheet2.Range("dcroque").Value + "; "
            ElseIf Sheet2.opt_NNGMM.Value = True Then
            End If
        
            If Sheet2.opt_ToDL1.Value = True Then
                strto = strto + Sheet2.Range("aoabary").Value + "; "
            ElseIf Sheet2.opt_CCDL1.Value = True Then
                strcc = strcc + Sheet2.Range("aoabary").Value + "; "
            ElseIf Sheet2.opt_NDL1.Value = True Then
            End If
            
            If Sheet2.opt_ToDL2.Value = True Then
                strto = strto + Sheet2.Range("rogbanez").Value + "; "
            ElseIf Sheet2.opt_CCDL2.Value = True Then
                strcc = strcc + Sheet2.Range("rogbanez").Value + "; "
            ElseIf Sheet2.opt_NDL2.Value = True Then
            End If
            
            If Sheet2.opt_ToDL3.Value = True Then
                strto = strto + Sheet2.Range("blbautista").Value + "; "
            ElseIf Sheet2.opt_CCDL3.Value = True Then
                strcc = strcc + Sheet2.Range("blbautista").Value + "; "
            ElseIf Sheet2.opt_NDL3.Value = True Then
            End If
            
            If Sheet2.opt_ToDL4.Value = True Then
                strto = strto + Sheet2.Range("ebbayle").Value + "; "
            ElseIf Sheet2.opt_CCDL4.Value = True Then
                strcc = strcc + Sheet2.Range("ebbayle").Value + "; "
            ElseIf Sheet2.opt_NDL4.Value = True Then
            End If
            
            If Sheet2.opt_ToDL6.Value = True Then
                strto = strto + Sheet2.Range("accruz").Value + "; "
            ElseIf Sheet2.opt_CCDL6.Value = True Then
                strcc = strcc + Sheet2.Range("accruz").Value + "; "
            ElseIf Sheet2.opt_NDL6.Value = True Then
            End If
            
            If Sheet2.opt_ToDL7.Value = True Then
                strto = strto + Sheet2.Range("mgdioso").Value + "; "
            ElseIf Sheet2.opt_CCDL7.Value = True Then
                strcc = strcc + Sheet2.Range("mgdioso").Value + "; "
            ElseIf Sheet2.opt_NDL7.Value = True Then
            End If
            
            If Sheet2.opt_ToDL8.Value = True Then
                strto = strto + Sheet2.Range("amestrella").Value + "; "
            ElseIf Sheet2.opt_CCDL8.Value = True Then
                strcc = strcc + Sheet2.Range("amestrella").Value + "; "
            ElseIf Sheet2.opt_NDL8.Value = True Then
            End If
            
            If Sheet2.opt_ToDL9.Value = True Then
                strto = strto + Sheet2.Range("drfonacier").Value + "; "
            ElseIf Sheet2.opt_CCDL9.Value = True Then
                strcc = strcc + Sheet2.Range("drfonacier").Value + "; "
            ElseIf Sheet2.opt_NDL9.Value = True Then
            End If
            
            If Sheet2.opt_ToDL10.Value = True Then
                strto = strto + Sheet2.Range("rjlim").Value + "; "
            ElseIf Sheet2.opt_CCDL10.Value = True Then
                strcc = strcc + Sheet2.Range("rjlim").Value + "; "
            ElseIf Sheet2.opt_NDL10.Value = True Then
            End If
            
            If Sheet2.opt_ToDL11.Value = True Then
                strto = strto + Sheet2.Range("ecmadrilejo").Value + "; "
            ElseIf Sheet2.opt_CCDL11.Value = True Then
                strcc = strcc + Sheet2.Range("ecmadrilejo").Value + "; "
            ElseIf Sheet2.opt_NDL11.Value = True Then
            End If
            
            If Sheet2.opt_ToDL12.Value = True Then
                strto = strto + Sheet2.Range("jvnaval").Value + "; "
            ElseIf Sheet2.opt_CCDL12.Value = True Then
                strcc = strcc + Sheet2.Range("jvnaval").Value + "; "
            ElseIf Sheet2.opt_NDL12.Value = True Then
            End If
            
            If Sheet2.opt_ToDL13.Value = True Then
                strto = strto + Sheet2.Range("wmsabile").Value + "; "
            ElseIf Sheet2.opt_CCDL13.Value = True Then
                strcc = strcc + Sheet2.Range("wmsabile").Value + "; "
            ElseIf Sheet2.opt_NDL13.Value = True Then
            End If
            
            If Sheet2.opt_ToDL14.Value = True Then
                strto = strto + Sheet2.Range("agsoro").Value + "; "
            ElseIf Sheet2.opt_CCDL14.Value = True Then
                strcc = strcc + Sheet2.Range("agsoro").Value + "; "
            ElseIf Sheet2.opt_NDL14.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS1.Value = True Then
                strto = strto + Sheet2.Range("csalejo").Value + "; "
            ElseIf Sheet2.opt_CCDLS1.Value = True Then
                strcc = strcc + Sheet2.Range("csalejo").Value + "; "
            ElseIf Sheet2.opt_NDLS1.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS3.Value = True Then
                strto = strto + Sheet2.Range("ecbuera").Value + "; "
            ElseIf Sheet2.opt_CCDLS3.Value = True Then
                strcc = strcc + Sheet2.Range("ecbuera").Value + "; "
            ElseIf Sheet2.opt_NDLS3.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS4.Value = True Then
                strto = strto + Sheet2.Range("kdcruz").Value + "; "
            ElseIf Sheet2.opt_CCDLS4.Value = True Then
                strcc = strcc + Sheet2.Range("kdcruz").Value + "; "
            ElseIf Sheet2.opt_NDLS4.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS5.Value = True Then
                strto = strto + Sheet2.Range("wdgdebelen").Value + "; "
            ElseIf Sheet2.opt_CCDLS5.Value = True Then
                strcc = strcc + Sheet2.Range("wdgdebelen").Value + "; "
            ElseIf Sheet2.opt_NDLS5.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS6.Value = True Then
                strto = strto + Sheet2.Range("nmestevez").Value + "; "
            ElseIf Sheet2.opt_CCDLS6.Value = True Then
                strcc = strcc + Sheet2.Range("nmestevez").Value + "; "
            ElseIf Sheet2.opt_NDLS6.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS7.Value = True Then
                strto = strto + Sheet2.Range("vdgrino").Value + "; "
            ElseIf Sheet2.opt_CCDLS7.Value = True Then
                strcc = strcc + Sheet2.Range("vdgrino").Value + "; "
            ElseIf Sheet2.opt_NDLS7.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS8.Value = True Then
                strto = strto + Sheet2.Range("jmhernandez").Value + "; "
            ElseIf Sheet2.opt_CCDLS8.Value = True Then
                strcc = strcc + Sheet2.Range("jmhernandez").Value + "; "
            ElseIf Sheet2.opt_NDLS8.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS9.Value = True Then
                strto = strto + Sheet2.Range("anmantes").Value + "; "
            ElseIf Sheet2.opt_CCDLS9.Value = True Then
                strcc = strcc + Sheet2.Range("anmantes").Value + "; "
            ElseIf Sheet2.opt_NDLS9.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS10.Value = True Then
                strto = strto + Sheet2.Range("lgpagay").Value + "; "
            ElseIf Sheet2.opt_CCDLS10.Value = True Then
                strcc = strcc + Sheet2.Range("lgpagay").Value + "; "
            ElseIf Sheet2.opt_NDLS10.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS11.Value = True Then
                strto = strto + Sheet2.Range("rjtorio").Value + "; "
            ElseIf Sheet2.opt_CCDLS11.Value = True Then
                strcc = strcc + Sheet2.Range("rjtorio").Value + "; "
            ElseIf Sheet2.opt_NDLS11.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS12.Value = True Then
                strto = strto + Sheet2.Range("vauypala").Value + "; "
            ElseIf Sheet2.opt_CCDLS12.Value = True Then
                strcc = strcc + Sheet2.Range("vauypala").Value + "; "
            ElseIf Sheet2.opt_NDLS12.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS13.Value = True Then
                strto = strto + Sheet2.Range("rdvizmanos").Value + "; "
            ElseIf Sheet2.opt_CCDLS13.Value = True Then
                strcc = strcc + Sheet2.Range("rdvizmanos").Value + "; "
            ElseIf Sheet2.opt_NDLS13.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS.Value = True Then
                strto = strto + Sheet2.Range("asgaba").Value + "; "
            ElseIf Sheet2.opt_CCDLS.Value = True Then
                strcc = strcc + Sheet2.Range("asgaba").Value + "; "
            ElseIf Sheet2.opt_NDLS.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO1.Value = True Then
                strto = strto + Sheet2.Range("fcbaul").Value + "; "
            ElseIf Sheet2.opt_CCFCNO1.Value = True Then
                strcc = strcc + Sheet2.Range("fcbaul").Value + "; "
            ElseIf Sheet2.opt_NFCNO1.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO2.Value = True Then
                strto = strto + Sheet2.Range("picables").Value + "; "
            ElseIf Sheet2.opt_CCFCNO2.Value = True Then
                strcc = strcc + Sheet2.Range("picables").Value + "; "
            ElseIf Sheet2.opt_NFCNO2.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO3.Value = True Then
                strto = strto + Sheet2.Range("aljimenez").Value + "; "
            ElseIf Sheet2.opt_CCFCNO3.Value = True Then
                strcc = strcc + Sheet2.Range("aljimenez").Value + "; "
            ElseIf Sheet2.opt_NFCNO3.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO4.Value = True Then
                strto = strto + Sheet2.Range("rtlampa").Value + "; "
            ElseIf Sheet2.opt_CCFCNO4.Value = True Then
                strcc = strcc + Sheet2.Range("rtlampa").Value + "; "
            ElseIf Sheet2.opt_NFCNO4.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO5.Value = True Then
                strto = strto + Sheet2.Range("oalinco").Value + "; "
            ElseIf Sheet2.opt_CCFCNO5.Value = True Then
                strcc = strcc + Sheet2.Range("oalinco").Value + "; "
            ElseIf Sheet2.opt_NFCNO5.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO6.Value = True Then
                strto = strto + Sheet2.Range("mfsantos").Value + "; "
            ElseIf Sheet2.opt_CCFCNO6.Value = True Then
                strcc = strcc + Sheet2.Range("mfsantos").Value + "; "
            ElseIf Sheet2.opt_NFCNO6.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO.Value = True Then
                strto = strto + Sheet2.Range("afcapiral").Value + "; "
            ElseIf Sheet2.opt_CCFCNO.Value = True Then
                strcc = strcc + Sheet2.Range("afcapiral").Value + "; "
            ElseIf Sheet2.opt_NFCNO.Value = True Then
            End If
                      
            strbody = ""
            strbody = strbody + "DILIMAN" + vbCrLf
            For i = 1 To Status.lst_DL.ListCount - 1
                If Status.lst_DL.List(i) Like "*Off-site Documentation Audit*" Or Status.lst_DL.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_DL.List(i) + vbCrLf
                End If
            Next i
        End If
    End If
    
    If Status.MultiPage1.Value = 4 Or Status.MultiPage1.Value = 6 Then
        If Status.MultiPage1.Pages(4).Caption = "*GARNET" Or Status.MultiPage1.Pages(6).Caption = "*GREENHILLS" Then
            If Sheet2.opt_ToGGH1.Value = True Then
                strto = strto + Sheet2.Range("ntalcantara").Value + "; "
            ElseIf Sheet2.opt_CCGGH1.Value = True Then
                strcc = strcc + Sheet2.Range("ntalcantara").Value + "; "
            ElseIf Sheet2.opt_NGGH1.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH2.Value = True Then
                strto = strto + Sheet2.Range("lpabapo").Value + "; "
            ElseIf Sheet2.opt_CCGGH2.Value = True Then
                strcc = strcc + Sheet2.Range("lpabapo").Value + "; "
            ElseIf Sheet2.opt_NGGH2.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH3.Value = True Then
                strto = strto + Sheet2.Range("raaquino").Value + "; "
            ElseIf Sheet2.opt_CCGGH3.Value = True Then
                strcc = strcc + Sheet2.Range("raaquino").Value + "; "
            ElseIf Sheet2.opt_NGGH3.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH5.Value = True Then
                strto = strto + Sheet2.Range("rrcerdon").Value + "; "
            ElseIf Sheet2.opt_CCGGH5.Value = True Then
                strcc = strcc + Sheet2.Range("rrcerdon").Value + "; "
            ElseIf Sheet2.opt_NGGH5.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH6.Value = True Then
                strto = strto + Sheet2.Range("lmcoma").Value + "; "
            ElseIf Sheet2.opt_CCGGH6.Value = True Then
                strcc = strcc + Sheet2.Range("lmcoma").Value + "; "
            ElseIf Sheet2.opt_NGGH6.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH7.Value = True Then
                strto = strto + Sheet2.Range("radeguzman").Value + "; "
            ElseIf Sheet2.opt_CCGGH7.Value = True Then
                strcc = strcc + Sheet2.Range("radeguzman").Value + "; "
            ElseIf Sheet2.opt_NGGH7.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH9.Value = True Then
                strto = strto + Sheet2.Range("nsreas").Value + "; "
            ElseIf Sheet2.opt_CCGGH9.Value = True Then
                strcc = strcc + Sheet2.Range("nsreas").Value + "; "
            ElseIf Sheet2.opt_NGGH9.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH10.Value = True Then
                strto = strto + Sheet2.Range("jugabriel").Value + "; "
            ElseIf Sheet2.opt_CCGGH10.Value = True Then
                strcc = strcc + Sheet2.Range("jugabriel").Value + "; "
            ElseIf Sheet2.opt_NGGH10.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH11.Value = True Then
                strto = strto + Sheet2.Range("mbgaspay").Value + "; "
            ElseIf Sheet2.opt_CCGGH11.Value = True Then
                strcc = strcc + Sheet2.Range("mbgaspay").Value + "; "
            ElseIf Sheet2.opt_NGGH11.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH12.Value = True Then
                strto = strto + Sheet2.Range("cgkatalbas").Value + "; "
            ElseIf Sheet2.opt_CCGGH12.Value = True Then
                strcc = strcc + Sheet2.Range("cgkatalbas").Value + "; "
            ElseIf Sheet2.opt_NGGH12.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH13.Value = True Then
                strto = strto + Sheet2.Range("nrlapus").Value + "; "
            ElseIf Sheet2.opt_CCGGH13.Value = True Then
                strcc = strcc + Sheet2.Range("nrlapus").Value + "; "
            ElseIf Sheet2.opt_NGGH13.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH14.Value = True Then
                strto = strto + Sheet2.Range("fslizardo").Value + "; "
            ElseIf Sheet2.opt_CCGGH14.Value = True Then
                strcc = strcc + Sheet2.Range("fslizardo").Value + "; "
            ElseIf Sheet2.opt_NGGH14.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH15.Value = True Then
                strto = strto + Sheet2.Range("drparrocha").Value + "; "
            ElseIf Sheet2.opt_CCGGH15.Value = True Then
                strcc = strcc + Sheet2.Range("drparrocha").Value + "; "
            ElseIf Sheet2.opt_NGGH15.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH16.Value = True Then
                strto = strto + Sheet2.Range("jdsarinas").Value + "; "
            ElseIf Sheet2.opt_CCGGH16.Value = True Then
                strcc = strcc + Sheet2.Range("jdsarinas").Value + "; "
            ElseIf Sheet2.opt_NGGH16.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH17.Value = True Then
                strto = strto + Sheet2.Range("rteves").Value + "; "
            ElseIf Sheet2.opt_CCGGH17.Value = True Then
                strcc = strcc + Sheet2.Range("rteves").Value + "; "
            ElseIf Sheet2.opt_NGGH17.Value = True Then
            End If

            If Sheet2.opt_ToFCNO1.Value = True Then
                strto = strto + Sheet2.Range("fcbaul").Value + "; "
            ElseIf Sheet2.opt_CCFCNO1.Value = True Then
                strcc = strcc + Sheet2.Range("fcbaul").Value + "; "
            ElseIf Sheet2.opt_NFCNO1.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO2.Value = True Then
                strto = strto + Sheet2.Range("picables").Value + "; "
            ElseIf Sheet2.opt_CCFCNO2.Value = True Then
                strcc = strcc + Sheet2.Range("picables").Value + "; "
            ElseIf Sheet2.opt_NFCNO2.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO3.Value = True Then
                strto = strto + Sheet2.Range("aljimenez").Value + "; "
            ElseIf Sheet2.opt_CCFCNO3.Value = True Then
                strcc = strcc + Sheet2.Range("aljimenez").Value + "; "
            ElseIf Sheet2.opt_NFCNO3.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO4.Value = True Then
                strto = strto + Sheet2.Range("rtlampa").Value + "; "
            ElseIf Sheet2.opt_CCFCNO4.Value = True Then
                strcc = strcc + Sheet2.Range("rtlampa").Value + "; "
            ElseIf Sheet2.opt_NFCNO4.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO5.Value = True Then
                strto = strto + Sheet2.Range("oalinco").Value + "; "
            ElseIf Sheet2.opt_CCFCNO5.Value = True Then
                strcc = strcc + Sheet2.Range("oalinco").Value + "; "
            ElseIf Sheet2.opt_NFCNO5.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO6.Value = True Then
                strto = strto + Sheet2.Range("mfsantos").Value + "; "
            ElseIf Sheet2.opt_CCFCNO6.Value = True Then
                strcc = strcc + Sheet2.Range("mfsantos").Value + "; "
            ElseIf Sheet2.opt_NFCNO6.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO7.Value = True Then
                strto = strto + Sheet2.Range("aeenano").Value + "; "
            ElseIf Sheet2.opt_CCFCNO7.Value = True Then
                strcc = strcc + Sheet2.Range("aeenano").Value + "; "
            ElseIf Sheet2.opt_NFCNO7.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO8.Value = True Then
                strto = strto + Sheet2.Range("apinciong").Value + "; "
            ElseIf Sheet2.opt_CCFCNO8.Value = True Then
                strcc = strcc + Sheet2.Range("apinciong").Value + "; "
            ElseIf Sheet2.opt_NFCNO8.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO9.Value = True Then
                strto = strto + Sheet2.Range("eanieva").Value + "; "
            ElseIf Sheet2.opt_CCFCNO9.Value = True Then
                strcc = strcc + Sheet2.Range("eanieva").Value + "; "
            ElseIf Sheet2.opt_NFCNO9.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO10.Value = True Then
                strto = strto + Sheet2.Range("jutamondong").Value + "; "
            ElseIf Sheet2.opt_CCFCNO10.Value = True Then
                strcc = strcc + Sheet2.Range("jutamondong").Value + "; "
            ElseIf Sheet2.opt_NFCNO10.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO11.Value = True Then
                strto = strto + Sheet2.Range("rcroque").Value + "; "
            ElseIf Sheet2.opt_CCFCNO11.Value = True Then
                strcc = strcc + Sheet2.Range("rcroque").Value + "; "
            ElseIf Sheet2.opt_NFCNO11.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO12.Value = True Then
                strto = strto + Sheet2.Range("mosena").Value + "; "
            ElseIf Sheet2.opt_CCFCNO12.Value = True Then
                strcc = strcc + Sheet2.Range("mosena").Value + "; "
            ElseIf Sheet2.opt_NFCNO12.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO.Value = True Then
                strto = strto + Sheet2.Range("afcapiral").Value + "; "
            ElseIf Sheet2.opt_CCFCNO.Value = True Then
                strcc = strcc + Sheet2.Range("afcapiral").Value + "; "
            ElseIf Sheet2.opt_NFCNO.Value = True Then
            End If
                      
            strbody = ""
            strbody = strbody + "GARNET" + vbCrLf
            For i = 1 To Status.lst_G.ListCount - 1
                If Status.lst_G.List(i) Like "*Off-site Documentation Audit*" Or Status.lst_G.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_G.List(i) + vbCrLf
                End If
            Next i
            strbody = strbody + "GREENHILLS" + vbCrLf
            For i = 1 To Status.lst_GH.ListCount - 1
                If Status.lst_GH.List(i) Like "*Off-site Documentation Audit*" Or Status.lst_GH.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_GH.List(i) + vbCrLf
                End If
            Next i
        End If
    End If
    
    If Status.MultiPage1.Value = 5 Then
        If Status.MultiPage1.Pages(5).Caption = "*SAMPALOC" Then
            If Sheet2.opt_ToS1.Value = True Then
                strto = strto + Sheet2.Range("emalsol").Value + "; "
            ElseIf Sheet2.opt_CCS1.Value = True Then
                strcc = strcc + Sheet2.Range("emalsol").Value + "; "
            ElseIf Sheet2.opt_NS1.Value = True Then
            End If
            
            If Sheet2.opt_ToS0.Value = True Then
                strto = strto + Sheet2.Range("jnadia").Value + "; "
            ElseIf Sheet2.opt_CCS0.Value = True Then
                strcc = strcc + Sheet2.Range("jnadia").Value + "; "
            ElseIf Sheet2.opt_NS0.Value = True Then
            End If
            
            If Sheet2.opt_ToS2.Value = True Then
                strto = strto + Sheet2.Range("aaagbayani").Value + "; "
            ElseIf Sheet2.opt_CCS2.Value = True Then
                strcc = strcc + Sheet2.Range("aaagbayani").Value + "; "
            ElseIf Sheet2.opt_NS2.Value = True Then
            End If
            
            If Sheet2.opt_ToS3.Value = True Then
                strto = strto + Sheet2.Range("gpaquino").Value + "; "
            ElseIf Sheet2.opt_CCS3.Value = True Then
                strcc = strcc + Sheet2.Range("gpaquino").Value + "; "
            ElseIf Sheet2.opt_NS3.Value = True Then
            End If
            
            If Sheet2.opt_ToS4.Value = True Then
                strto = strto + Sheet2.Range("rhatendido").Value + "; "
            ElseIf Sheet2.opt_CCS4.Value = True Then
                strcc = strcc + Sheet2.Range("rhatendido").Value + "; "
            ElseIf Sheet2.opt_NS4.Value = True Then
            End If
            
            If Sheet2.opt_ToS5.Value = True Then
                strto = strto + Sheet2.Range("papagtalunan").Value + "; "
            ElseIf Sheet2.opt_CCS5.Value = True Then
                strcc = strcc + Sheet2.Range("papagtalunan").Value + "; "
            ElseIf Sheet2.opt_NS5.Value = True Then
            End If
            
            If Sheet2.opt_ToS6.Value = True Then
                strto = strto + Sheet2.Range("lpparadero").Value + "; "
            ElseIf Sheet2.opt_CCS6.Value = True Then
                strcc = strcc + Sheet2.Range("lpparadero").Value + "; "
            ElseIf Sheet2.opt_NS6.Value = True Then
            End If

            If Sheet2.opt_ToS7.Value = True Then
                strto = strto + Sheet2.Range("nnelegado").Value + "; "
            ElseIf Sheet2.opt_CCS7.Value = True Then
                strcc = strcc + Sheet2.Range("nnelegado").Value + "; "
            ElseIf Sheet2.opt_NS7.Value = True Then
            End If
            
            If Sheet2.opt_ToS9.Value = True Then
                strto = strto + Sheet2.Range("almgonzales").Value + "; "
            ElseIf Sheet2.opt_CCS9.Value = True Then
                strcc = strcc + Sheet2.Range("almgonzales").Value + "; "
            ElseIf Sheet2.opt_NS9.Value = True Then
            End If
            
            If Sheet2.opt_ToS11.Value = True Then
                strto = strto + Sheet2.Range("rvhundana").Value + "; "
            ElseIf Sheet2.opt_CCS11.Value = True Then
                strcc = strcc + Sheet2.Range("rvhundana").Value + "; "
            ElseIf Sheet2.opt_NS11.Value = True Then
            End If
            
            If Sheet2.opt_ToS12.Value = True Then
                strto = strto + Sheet2.Range("famacadaeg").Value + "; "
            ElseIf Sheet2.opt_CCS12.Value = True Then
                strcc = strcc + Sheet2.Range("famacadaeg").Value + "; "
            ElseIf Sheet2.opt_NS12.Value = True Then
            End If
            
            If Sheet2.opt_ToS13.Value = True Then
                strto = strto + Sheet2.Range("rsmariano").Value + "; "
            ElseIf Sheet2.opt_CCS13.Value = True Then
                strcc = strcc + Sheet2.Range("rsmariano").Value + "; "
            ElseIf Sheet2.opt_NS13.Value = True Then
            End If
            
            If Sheet2.opt_ToS14.Value = True Then
                strto = strto + Sheet2.Range("aenito").Value + "; "
            ElseIf Sheet2.opt_CCS14.Value = True Then
                strcc = strcc + Sheet2.Range("aenito").Value + "; "
            ElseIf Sheet2.opt_NS14.Value = True Then
            End If
            
            If Sheet2.opt_ToS15.Value = True Then
                strto = strto + Sheet2.Range("basoriano").Value + "; "
            ElseIf Sheet2.opt_CCS15.Value = True Then
                strcc = strcc + Sheet2.Range("basoriano").Value + "; "
            ElseIf Sheet2.opt_NS15.Value = True Then
            End If

            If Sheet2.opt_ToDLS1.Value = True Then
                strto = strto + Sheet2.Range("csalejo").Value + "; "
            ElseIf Sheet2.opt_CCDLS1.Value = True Then
                strcc = strcc + Sheet2.Range("csalejo").Value + "; "
            ElseIf Sheet2.opt_NDLS1.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS3.Value = True Then
                strto = strto + Sheet2.Range("ecbuera").Value + "; "
            ElseIf Sheet2.opt_CCDLS3.Value = True Then
                strcc = strcc + Sheet2.Range("ecbuera").Value + "; "
            ElseIf Sheet2.opt_NDLS3.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS4.Value = True Then
                strto = strto + Sheet2.Range("kdcruz").Value + "; "
            ElseIf Sheet2.opt_CCDLS4.Value = True Then
                strcc = strcc + Sheet2.Range("kdcruz").Value + "; "
            ElseIf Sheet2.opt_NDLS4.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS5.Value = True Then
                strto = strto + Sheet2.Range("wdgdebelen").Value + "; "
            ElseIf Sheet2.opt_CCDLS5.Value = True Then
                strcc = strcc + Sheet2.Range("wdgdebelen").Value + "; "
            ElseIf Sheet2.opt_NDLS5.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS6.Value = True Then
                strto = strto + Sheet2.Range("nmestevez").Value + "; "
            ElseIf Sheet2.opt_CCDLS6.Value = True Then
                strcc = strcc + Sheet2.Range("nmestevez").Value + "; "
            ElseIf Sheet2.opt_NDLS6.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS7.Value = True Then
                strto = strto + Sheet2.Range("vdgrino").Value + "; "
            ElseIf Sheet2.opt_CCDLS7.Value = True Then
                strcc = strcc + Sheet2.Range("vdgrino").Value + "; "
            ElseIf Sheet2.opt_NDLS7.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS8.Value = True Then
                strto = strto + Sheet2.Range("jmhernandez").Value + "; "
            ElseIf Sheet2.opt_CCDLS8.Value = True Then
                strcc = strcc + Sheet2.Range("jmhernandez").Value + "; "
            ElseIf Sheet2.opt_NDLS8.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS9.Value = True Then
                strto = strto + Sheet2.Range("anmantes").Value + "; "
            ElseIf Sheet2.opt_CCDLS9.Value = True Then
                strcc = strcc + Sheet2.Range("anmantes").Value + "; "
            ElseIf Sheet2.opt_NDLS9.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS10.Value = True Then
                strto = strto + Sheet2.Range("lgpagay").Value + "; "
            ElseIf Sheet2.opt_CCDLS10.Value = True Then
                strcc = strcc + Sheet2.Range("lgpagay").Value + "; "
            ElseIf Sheet2.opt_NDLS10.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS11.Value = True Then
                strto = strto + Sheet2.Range("rjtorio").Value + "; "
            ElseIf Sheet2.opt_CCDLS11.Value = True Then
                strcc = strcc + Sheet2.Range("rjtorio").Value + "; "
            ElseIf Sheet2.opt_NDLS11.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS12.Value = True Then
                strto = strto + Sheet2.Range("vauypala").Value + "; "
            ElseIf Sheet2.opt_CCDLS12.Value = True Then
                strcc = strcc + Sheet2.Range("vauypala").Value + "; "
            ElseIf Sheet2.opt_NDLS12.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS13.Value = True Then
                strto = strto + Sheet2.Range("rdvizmanos").Value + "; "
            ElseIf Sheet2.opt_CCDLS13.Value = True Then
                strcc = strcc + Sheet2.Range("rdvizmanos").Value + "; "
            ElseIf Sheet2.opt_NDLS13.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS.Value = True Then
                strto = strto + Sheet2.Range("asgaba").Value + "; "
            ElseIf Sheet2.opt_CCDLS.Value = True Then
                strcc = strcc + Sheet2.Range("asgaba").Value + "; "
            ElseIf Sheet2.opt_NDLS.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO1.Value = True Then
                strto = strto + Sheet2.Range("fcbaul").Value + "; "
            ElseIf Sheet2.opt_CCFCNO1.Value = True Then
                strcc = strcc + Sheet2.Range("fcbaul").Value + "; "
            ElseIf Sheet2.opt_NFCNO1.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO2.Value = True Then
                strto = strto + Sheet2.Range("picables").Value + "; "
            ElseIf Sheet2.opt_CCFCNO2.Value = True Then
                strcc = strcc + Sheet2.Range("picables").Value + "; "
            ElseIf Sheet2.opt_NFCNO2.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO3.Value = True Then
                strto = strto + Sheet2.Range("aljimenez").Value + "; "
            ElseIf Sheet2.opt_CCFCNO3.Value = True Then
                strcc = strcc + Sheet2.Range("aljimenez").Value + "; "
            ElseIf Sheet2.opt_NFCNO3.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO4.Value = True Then
                strto = strto + Sheet2.Range("rtlampa").Value + "; "
            ElseIf Sheet2.opt_CCFCNO4.Value = True Then
                strcc = strcc + Sheet2.Range("rtlampa").Value + "; "
            ElseIf Sheet2.opt_NFCNO4.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO5.Value = True Then
                strto = strto + Sheet2.Range("oalinco").Value + "; "
            ElseIf Sheet2.opt_CCFCNO5.Value = True Then
                strcc = strcc + Sheet2.Range("oalinco").Value + "; "
            ElseIf Sheet2.opt_NFCNO5.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO6.Value = True Then
                strto = strto + Sheet2.Range("mfsantos").Value + "; "
            ElseIf Sheet2.opt_CCFCNO6.Value = True Then
                strcc = strcc + Sheet2.Range("mfsantos").Value + "; "
            ElseIf Sheet2.opt_NFCNO6.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO.Value = True Then
                strto = strto + Sheet2.Range("afcapiral").Value + "; "
            ElseIf Sheet2.opt_CCFCNO.Value = True Then
                strcc = strcc + Sheet2.Range("afcapiral").Value + "; "
            ElseIf Sheet2.opt_NFCNO.Value = True Then
            End If
                      
            strbody = ""
            strbody = strbody + "SAMPALOC" + vbCrLf
            For i = 1 To Status.lst_S.ListCount - 1
                If Status.lst_S.List(i) Like "*Off-site Documentation Audit*" Or Status.lst_S.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_S.List(i) + vbCrLf
                End If
            Next i
        End If
    End If
    
    If Status.MultiPage1.Value = 7 Then
        If Status.MultiPage1.Pages(7).Caption = "*CEBU" And chk_APFMC.Value = False Then
            If Sheet2.opt_ToC1.Value = True Then
                strto = strto + Sheet2.Range("emallosada").Value + "; "
            ElseIf Sheet2.opt_CCC1.Value = True Then
                strcc = strcc + Sheet2.Range("emallosada").Value + "; "
            ElseIf Sheet2.opt_NC1.Value = True Then
            End If
            
            If Sheet2.opt_ToC2.Value = True Then
                strto = strto + Sheet2.Range("rbarias").Value + "; "
            ElseIf Sheet2.opt_CCC2.Value = True Then
                strcc = strcc + Sheet2.Range("rbarias").Value + "; "
            ElseIf Sheet2.opt_NC2.Value = True Then
            End If
            
            If Sheet2.opt_ToC3.Value = True Then
                strto = strto + Sheet2.Range("aabaguio").Value + "; "
            ElseIf Sheet2.opt_CCC3.Value = True Then
                strcc = strcc + Sheet2.Range("aabaguio").Value + "; "
            ElseIf Sheet2.opt_NC3.Value = True Then
            End If
            
            If Sheet2.opt_ToC4.Value = True Then
                strto = strto + Sheet2.Range("lacabigas").Value + "; "
            ElseIf Sheet2.opt_CCC4.Value = True Then
                strcc = strcc + Sheet2.Range("lacabigas").Value + "; "
            ElseIf Sheet2.opt_NC4.Value = True Then
            End If
            
            If Sheet2.opt_ToC5.Value = True Then
                strto = strto + Sheet2.Range("jjcarumba").Value + "; "
            ElseIf Sheet2.opt_CCC5.Value = True Then
                strcc = strcc + Sheet2.Range("jjcarumba").Value + "; "
            ElseIf Sheet2.opt_NC5.Value = True Then
            End If
            
            If Sheet2.opt_ToC6.Value = True Then
                strto = strto + Sheet2.Range("rgconchas").Value + "; "
            ElseIf Sheet2.opt_CCC6.Value = True Then
                strcc = strcc + Sheet2.Range("rgconchas").Value + "; "
            ElseIf Sheet2.opt_NC6.Value = True Then
            End If
            
            If Sheet2.opt_ToC7.Value = True Then
                strto = strto + Sheet2.Range("vbcuevas").Value + "; "
            ElseIf Sheet2.opt_CCC7.Value = True Then
                strcc = strcc + Sheet2.Range("vbcuevas").Value + "; "
            ElseIf Sheet2.opt_NC7.Value = True Then
            End If
            
            If Sheet2.opt_ToC8.Value = True Then
                strto = strto + Sheet2.Range("jbdesemprado").Value + "; "
            ElseIf Sheet2.opt_CCC8.Value = True Then
                strcc = strcc + Sheet2.Range("jbdesemprado").Value + "; "
            ElseIf Sheet2.opt_NC8.Value = True Then
            End If
            
            If Sheet2.opt_ToC9.Value = True Then
                strto = strto + Sheet2.Range("rddesquitado").Value + "; "
            ElseIf Sheet2.opt_CCC9.Value = True Then
                strcc = strcc + Sheet2.Range("rddesquitado").Value + "; "
            ElseIf Sheet2.opt_NC9.Value = True Then
            End If
            
            If Sheet2.opt_ToC10.Value = True Then
                strto = strto + Sheet2.Range("mmdevero").Value + "; "
            ElseIf Sheet2.opt_CCC10.Value = True Then
                strcc = strcc + Sheet2.Range("mmdevero").Value + "; "
            ElseIf Sheet2.opt_NC10.Value = True Then
            End If
            
            If Sheet2.opt_ToC11.Value = True Then
                strto = strto + Sheet2.Range("jcfelisarta").Value + "; "
            ElseIf Sheet2.opt_CCC11.Value = True Then
                strcc = strcc + Sheet2.Range("jcfelisarta").Value + "; "
            ElseIf Sheet2.opt_NC11.Value = True Then
            End If
            
            If Sheet2.opt_ToC12.Value = True Then
                strto = strto + Sheet2.Range("rdflores").Value + "; "
            ElseIf Sheet2.opt_CCC12.Value = True Then
                strcc = strcc + Sheet2.Range("rdflores").Value + "; "
            ElseIf Sheet2.opt_NC12.Value = True Then
            End If
            
            If Sheet2.opt_ToC13.Value = True Then
                strto = strto + Sheet2.Range("gpintes").Value + "; "
            ElseIf Sheet2.opt_CCC13.Value = True Then
                strcc = strcc + Sheet2.Range("gpintes").Value + "; "
            ElseIf Sheet2.opt_NC13.Value = True Then
            End If
            
            If Sheet2.opt_ToC14.Value = True Then
                strto = strto + Sheet2.Range("dsisleta").Value + "; "
            ElseIf Sheet2.opt_CCC14.Value = True Then
                strcc = strcc + Sheet2.Range("dsisleta").Value + "; "
            ElseIf Sheet2.opt_NC14.Value = True Then
            End If
            
            If Sheet2.opt_ToC15.Value = True Then
                strto = strto + Sheet2.Range("rmlocaylocay").Value + "; "
            ElseIf Sheet2.opt_CCC15.Value = True Then
                strcc = strcc + Sheet2.Range("rmlocaylocay").Value + "; "
            ElseIf Sheet2.opt_NC15.Value = True Then
            End If
            
            If Sheet2.opt_ToC16.Value = True Then
                strto = strto + Sheet2.Range("mmnadal").Value + "; "
            ElseIf Sheet2.opt_CCC16.Value = True Then
                strcc = strcc + Sheet2.Range("mmnadal").Value + "; "
            ElseIf Sheet2.opt_NC16.Value = True Then
            End If
            
            If Sheet2.opt_ToC17.Value = True Then
                strto = strto + Sheet2.Range("npompad").Value + "; "
            ElseIf Sheet2.opt_CCC17.Value = True Then
                strcc = strcc + Sheet2.Range("npompad").Value + "; "
            ElseIf Sheet2.opt_NC17.Value = True Then
            End If
            
            If Sheet2.opt_ToC18.Value = True Then
                strto = strto + Sheet2.Range("mpoplas").Value + "; "
            ElseIf Sheet2.opt_CCC18.Value = True Then
                strcc = strcc + Sheet2.Range("mpoplas").Value + "; "
            ElseIf Sheet2.opt_NC18.Value = True Then
            End If
            
            If Sheet2.opt_ToC19.Value = True Then
                strto = strto + Sheet2.Range("dbpepito").Value + "; "
            ElseIf Sheet2.opt_CCC19.Value = True Then
                strcc = strcc + Sheet2.Range("dbpepito").Value + "; "
            ElseIf Sheet2.opt_NC19.Value = True Then
            End If
            
            If Sheet2.opt_ToC20.Value = True Then
                strto = strto + Sheet2.Range("izpono").Value + "; "
            ElseIf Sheet2.opt_CCC20.Value = True Then
                strcc = strcc + Sheet2.Range("izpono").Value + "; "
            ElseIf Sheet2.opt_NC20.Value = True Then
            End If
            
            If Sheet2.opt_ToC21.Value = True Then
                strto = strto + Sheet2.Range("clrosales").Value + "; "
            ElseIf Sheet2.opt_CCC21.Value = True Then
                strcc = strcc + Sheet2.Range("clrosales").Value + "; "
            ElseIf Sheet2.opt_NC21.Value = True Then
            End If
            
            If Sheet2.opt_ToC22.Value = True Then
                strto = strto + Sheet2.Range("rasalas").Value + "; "
            ElseIf Sheet2.opt_CCC22.Value = True Then
                strcc = strcc + Sheet2.Range("rasalas").Value + "; "
            ElseIf Sheet2.opt_NC22.Value = True Then
            End If

            If Sheet2.opt_ToC24.Value = True Then
                strto = strto + Sheet2.Range("mdsarcuaga").Value + "; "
            ElseIf Sheet2.opt_CCC24.Value = True Then
                strcc = strcc + Sheet2.Range("mdsarcuaga").Value + "; "
            ElseIf Sheet2.opt_NC24.Value = True Then
            End If
            
            If Sheet2.opt_ToC25.Value = True Then
                strto = strto + Sheet2.Range("mlsarmiento").Value + "; "
            ElseIf Sheet2.opt_CCC25.Value = True Then
                strcc = strcc + Sheet2.Range("mlsarmiento").Value + "; "
            ElseIf Sheet2.opt_NC25.Value = True Then
            End If
            
            If Sheet2.opt_ToC26.Value = True Then
                strto = strto + Sheet2.Range("vjson").Value + "; "
            ElseIf Sheet2.opt_CCC26.Value = True Then
                strcc = strcc + Sheet2.Range("vjson").Value + "; "
            ElseIf Sheet2.opt_NC26.Value = True Then
            End If
            
            If Sheet2.opt_ToC27.Value = True Then
                strto = strto + Sheet2.Range("jltacan").Value + "; "
            ElseIf Sheet2.opt_CCC27.Value = True Then
                strcc = strcc + Sheet2.Range("jltacan").Value + "; "
            ElseIf Sheet2.opt_NC27.Value = True Then
            End If
            
            If Sheet2.opt_ToC28.Value = True Then
                strto = strto + Sheet2.Range("fotamarra").Value + "; "
            ElseIf Sheet2.opt_CCC28.Value = True Then
                strcc = strcc + Sheet2.Range("fotamarra").Value + "; "
            ElseIf Sheet2.opt_NC28.Value = True Then
            End If
            
            If Sheet2.opt_ToC29.Value = True Then
                strto = strto + Sheet2.Range("metejano").Value + "; "
            ElseIf Sheet2.opt_CCC29.Value = True Then
                strcc = strcc + Sheet2.Range("metejano").Value + "; "
            ElseIf Sheet2.opt_NC29.Value = True Then
            End If
            
            If Sheet2.opt_ToC291.Value = True Then
                strto = strto + Sheet2.Range("blynot").Value + "; "
            ElseIf Sheet2.opt_CCC291.Value = True Then
                strcc = strcc + Sheet2.Range("blynot").Value + "; "
            ElseIf Sheet2.opt_NC291.Value = True Then
            End If
            
            If Sheet2.opt_ToC30.Value = True Then
                strto = strto + Sheet2.Range("eddivinagracia").Value + "; "
            ElseIf Sheet2.opt_CCC30.Value = True Then
                strcc = strcc + Sheet2.Range("eddivinagracia").Value + "; "
            ElseIf Sheet2.opt_NC30.Value = True Then
            End If
            
            If Sheet2.opt_ToC31.Value = True Then
                strto = strto + Sheet2.Range("rngloria").Value + "; "
            ElseIf Sheet2.opt_CCC31.Value = True Then
                strcc = strcc + Sheet2.Range("rngloria").Value + "; "
            ElseIf Sheet2.opt_NC31.Value = True Then
            End If
            
            If Sheet2.opt_ToC32.Value = True Then
                strto = strto + Sheet2.Range("rrson").Value + "; "
            ElseIf Sheet2.opt_CCC32.Value = True Then
                strcc = strcc + Sheet2.Range("rrson").Value + "; "
            ElseIf Sheet2.opt_NC32.Value = True Then
            End If
            
            If Sheet2.opt_ToC33.Value = True Then
                strto = strto + Sheet2.Range("cminoferio").Value + "; "
            ElseIf Sheet2.opt_CCC33.Value = True Then
                strcc = strcc + Sheet2.Range("cminoferio").Value + "; "
            ElseIf Sheet2.opt_NC33.Value = True Then
            End If
            
            If Sheet2.opt_ToC34.Value = True Then
                strto = strto + Sheet2.Range("jrmaninang").Value + "; "
            ElseIf Sheet2.opt_CCC34.Value = True Then
                strcc = strcc + Sheet2.Range("jrmaninang").Value + "; "
            ElseIf Sheet2.opt_NC34.Value = True Then
            End If
            
            If Sheet2.opt_ToC35.Value = True Then
                strto = strto + Sheet2.Range("mjvendiola").Value + "; "
            ElseIf Sheet2.opt_CCC35.Value = True Then
                strcc = strcc + Sheet2.Range("mjvendiola").Value + "; "
            ElseIf Sheet2.opt_NC35.Value = True Then
            End If
       
            strbody = ""
            strbody = strbody + "CEBU" + vbCrLf
            For i = 1 To Status.lst_C.ListCount - 1
                If Status.lst_C.List(i) Like "*Off-site Documentation Audit*" Or Status.lst_C.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_C.List(i) + vbCrLf
                End If
            Next i
            
        ElseIf Status.MultiPage1.Pages(7).Caption = "*CEBU" And Status.chk_APFMC.Value = True Then
            If Sheet2.opt_ToCAF1.Value = True Then
                strto = strto + Sheet2.Range("jjdagay").Value + "; "
            ElseIf Sheet2.opt_CCCAF1.Value = True Then
                strcc = strcc + Sheet2.Range("jjdagay").Value + "; "
            ElseIf Sheet2.opt_NCAF1.Value = True Then
            End If
            
            If Sheet2.opt_ToCAF2.Value = True Then
                strto = strto + Sheet2.Range("lgelicanal").Value + "; "
            ElseIf Sheet2.opt_CCCAF2.Value = True Then
                strcc = strcc + Sheet2.Range("lgelicanal").Value + "; "
            ElseIf Sheet2.opt_NCAF2.Value = True Then
            End If
            
            If Sheet2.opt_ToCAF3.Value = True Then
                strto = strto + Sheet2.Range("mycondrillon").Value + "; "
            ElseIf Sheet2.opt_CCCAF3.Value = True Then
                strcc = strcc + Sheet2.Range("mycondrillon").Value + "; "
            ElseIf Sheet2.opt_NCAF3.Value = True Then
            End If
            
            strbody = ""
            strbody = strbody + "CEBU" + vbCrLf
            For i = 1 To Status.lst_C.ListCount - 1
                If Status.lst_C.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_C.List(i) + vbCrLf
                End If
            Next i
        End If
    End If
    
    If Status.MultiPage1.Pages(3).Caption = "*DILIMAN" Then
        If Status.chk_APFMDL.Value = True And Status.MultiPage1.Value = 3 Then
            If Sheet2.opt_ToMMAF1.Value = True Then
                strto = strto + Sheet2.Range("acty").Value + "; "
            ElseIf Sheet2.opt_CCMMAF1.Value = True Then
                strcc = strcc + Sheet2.Range("acty").Value + "; "
            ElseIf Sheet2.opt_NMMAF1.Value = True Then
            End If
                
            If Sheet2.opt_ToMMAF3.Value = True Then
                strto = strto + Sheet2.Range("achaling").Value + "; "
            ElseIf Sheet2.opt_CCMMAF3.Value = True Then
                strcc = strcc + Sheet2.Range("achaling").Value + "; "
            ElseIf Sheet2.opt_NMMAF3.Value = True Then
            End If
                
            If Sheet2.opt_ToMMAF4.Value = True Then
                strto = strto + Sheet2.Range("rdmontemayor").Value + "; "
            ElseIf Sheet2.opt_CCMMAF4.Value = True Then
                strcc = strcc + Sheet2.Range("rdmontemayor").Value + "; "
            ElseIf Sheet2.opt_NMMAF4.Value = True Then
            End If
                
            If Sheet2.opt_ToMMAF8.Value = True Then
                strto = strto + Sheet2.Range("rgreyes").Value + "; "
            ElseIf Sheet2.opt_CCMMAF8.Value = True Then
                strcc = strcc + Sheet2.Range("rgreyes").Value + "; "
            ElseIf Sheet2.opt_NMMAF8.Value = True Then
            End If
                
            If Sheet2.opt_ToMMAF9.Value = True Then
                strto = strto + Sheet2.Range("ggancaya").Value + "; "
            ElseIf Sheet2.opt_CCMMAF9.Value = True Then
                strcc = strcc + Sheet2.Range("ggancaya").Value + "; "
            ElseIf Sheet2.opt_NMMAF9.Value = True Then
            End If
                
            strbody = ""
            strbody = strbody + "DILIMAN" + vbCrLf
            For i = 1 To Status.lst_DL.ListCount - 1
                If Status.lst_DL.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_DL.List(i) + vbCrLf
                End If
            Next i
        End If
    End If
    
    If Status.MultiPage1.Pages(4).Caption = "*GARNET" Then
        If Status.chk_APFMG.Value = True And Status.MultiPage1.Value = 4 Then
            
            If Sheet2.opt_ToMMAF1.Value = True Then
                strto = strto + Sheet2.Range("acty").Value + "; "
            ElseIf Sheet2.opt_CCMMAF1.Value = True Then
                strcc = strcc + Sheet2.Range("acty").Value + "; "
            ElseIf Sheet2.opt_NMMAF1.Value = True Then
            End If
                
            If Sheet2.opt_ToMMAF6.Value = True Then
                strto = strto + Sheet2.Range("vfserrano").Value + "; "
            ElseIf Sheet2.opt_CCMMAF6.Value = True Then
                strcc = strcc + Sheet2.Range("vfserrano").Value + "; "
            ElseIf Sheet2.opt_NMMAF6.Value = True Then
            End If
            
            If Sheet2.opt_ToMMAF7.Value = True Then
                strto = strto + Sheet2.Range("apoh").Value + "; "
            ElseIf Sheet2.opt_CCMMAF7.Value = True Then
                strcc = strcc + Sheet2.Range("apoh").Value + "; "
            ElseIf Sheet2.opt_NMMAF7.Value = True Then
            End If
                
            If Sheet2.opt_ToMMAF8.Value = True Then
                strto = strto + Sheet2.Range("rgreyes").Value + "; "
            ElseIf Sheet2.opt_CCMMAF8.Value = True Then
                strcc = strcc + Sheet2.Range("rgreyes").Value + "; "
            ElseIf Sheet2.opt_NMMAF8.Value = True Then
            End If
                
            If Sheet2.opt_ToMMAF11.Value = True Then
                strto = strto + Sheet2.Range("jymendoza").Value + "; "
            ElseIf Sheet2.opt_CCMMAF11.Value = True Then
                strcc = strcc + Sheet2.Range("jymendoza").Value + "; "
            ElseIf Sheet2.opt_NMMAF11.Value = True Then
            End If
                
            strbody = ""
            strbody = strbody + "GARNET" + vbCrLf
            For i = 1 To Status.lst_G.ListCount - 1
                If Status.lst_G.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_G.List(i) + vbCrLf
                End If
            Next i
        End If
    End If
        
    If Status.MultiPage1.Pages(5).Caption = "*SAMPALOC" Then
        If Status.chk_APFMS.Value = True And Status.MultiPage1.Value = 5 Then
            If Sheet2.opt_ToMMAF1.Value = True Then
                strto = strto + Sheet2.Range("acty").Value + "; "
            ElseIf Sheet2.opt_CCMMAF1.Value = True Then
                strcc = strcc + Sheet2.Range("acty").Value + "; "
            ElseIf Sheet2.opt_NMMAF1.Value = True Then
            End If
                
            If Sheet2.opt_ToMMAF2.Value = True Then
                strto = strto + Sheet2.Range("rdrubin").Value + "; "
            ElseIf Sheet2.opt_CCMMAF2.Value = True Then
                strcc = strcc + Sheet2.Range("rdrubin").Value + "; "
            ElseIf Sheet2.opt_NMMAF2.Value = True Then
            End If
                
            strbody = ""
            strbody = strbody + "SAMPALOC" + vbCrLf
            For i = 1 To Status.lst_S.ListCount - 1
                If Status.lst_S.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_S.List(i) + vbCrLf
                End If
            Next i
        End If
    End If
            
    If Status.MultiPage1.Pages(7).Caption = "*GREENHILLS" Then
        If Status.chk_APFMGH.Value = True And Status.MultiPage1.Value = 7 Then
            If Sheet2.opt_ToMMAF1.Value = True Then
                strto = strto + Sheet2.Range("acty").Value + "; "
            ElseIf Sheet2.opt_CCMMAF1.Value = True Then
                strcc = strcc + Sheet2.Range("acty").Value + "; "
            ElseIf Sheet2.opt_NMMAF1.Value = True Then
            End If
                
            If Sheet2.opt_ToMMAF5.Value = True Then
                strto = strto + Sheet2.Range("evalayon").Value + "; "
            ElseIf Sheet2.opt_CCMMAF5.Value = True Then
                strcc = strcc + Sheet2.Range("evalayon").Value + "; "
            ElseIf Sheet2.opt_NMMAF5.Value = True Then
            End If
                
            If Sheet2.opt_ToMMAF8.Value = True Then
                strto = strto + Sheet2.Range("rgreyes").Value + "; "
            ElseIf Sheet2.opt_CCMMAF8.Value = True Then
                strcc = strcc + Sheet2.Range("rgreyes").Value + "; "
            ElseIf Sheet2.opt_NMMAF8.Value = True Then
            End If
                
            If Sheet2.opt_ToMMAF10.Value = True Then
                strto = strto + Sheet2.Range("dpstotomas").Value + "; "
            ElseIf Sheet2.opt_CCMMAF10.Value = True Then
                strcc = strcc + Sheet2.Range("dpstotomas").Value + "; "
            ElseIf Sheet2.opt_NMMAF10.Value = True Then
            End If
                
            strbody = ""
            strbody = strbody + "GREENHILLS" + vbCrLf
            For i = 1 To Status.lst_GH.ListCount - 1
                If Status.lst_GH.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_GH.List(i) + vbCrLf
                End If
            Next i
        End If
    End If
    
    Dim signature As String
    
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
        .Subject = "BC Gov IA Deadline"
        .Body = strbody
        .Attachments.Add (Sheet3.Range("A11").Value)
        .Attachments.Add Sheet3.Range("A5").Value, olByValue, 0
        .HTMLBody = "<style>body{color:red;font-weight:bold;}</style><h2>DUE DATE NEAR EXPIRATION:</h2>" & .HTMLBody & Sheet3.Range("A2").Value & "<br><br>" & "<img src='cid:" & Sheet3.Range("A8").Value & "'>" & signature
        .Send
    End With
    On Error GoTo 0
    
    Set OutMail = Nothing
    Set OutApp = Nothing
    
    MsgBox "Email Sent!", vbOKOnly + vbInformation, "Alert"

End Sub

Private Sub btn_View_Click()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Dim strto As String
    Dim strcc As String
    Dim i As Integer

    If Sheet2.opt_To2.Value = True Then
        strto = strto + Sheet2.Range("rdvillasenor").Value + "; "
    ElseIf Sheet2.opt_CC2.Value = True Then
        strcc = strcc + Sheet2.Range("rdvillasenor").Value + "; "
    ElseIf Sheet2.opt_N2.Value = True Then
    End If
    
    If Sheet2.opt_To3.Value = True Then
        strto = strto + Sheet2.Range("aiseeco").Value + "; "
    ElseIf Sheet2.opt_CC3.Value = True Then
        strcc = strcc + Sheet2.Range("aiseeco").Value + "; "
    ElseIf Sheet2.opt_N3.Value = True Then
    End If
    
    If Sheet2.opt_To4.Value = True Then
        strto = strto + Sheet2.Range("megutierrez").Value + "; "
    ElseIf Sheet2.opt_CC4.Value = True Then
        strcc = strcc + Sheet2.Range("megutierrez").Value + "; "
    ElseIf Sheet2.opt_N4.Value = True Then
    End If
    
    If Sheet2.opt_To5.Value = True Then
        strto = strto + Sheet2.Range("janacpil").Value + "; "
    ElseIf Sheet2.opt_CC5.Value = True Then
        strcc = strcc + Sheet2.Range("janacpil").Value + "; "
    ElseIf Sheet2.opt_N5.Value = True Then
    End If
    
    If Sheet2.opt_To6.Value = True Then
        strto = strto + Sheet2.Range("hclim").Value + "; "
    ElseIf Sheet2.opt_CC6.Value = True Then
        strcc = strcc + Sheet2.Range("hclim").Value + "; "
    ElseIf Sheet2.opt_N6.Value = True Then
    End If
    
    If Sheet2.opt_To7.Value = True Then
        strto = strto + Sheet2.Range("bccordova").Value + "; "
    ElseIf Sheet2.opt_CC7.Value = True Then
        strcc = strcc + Sheet2.Range("bccordova").Value + "; "
    ElseIf Sheet2.opt_N7.Value = True Then
    End If
    
    If Sheet2.opt_To8.Value = True Then
        strto = strto + Sheet2.Range("jpjandayan").Value + "; "
    ElseIf Sheet2.opt_CC8.Value = True Then
        strcc = strcc + Sheet2.Range("jpjandayan").Value + "; "
    ElseIf Sheet2.opt_N8.Value = True Then
    End If
    
    If Sheet2.opt_To9.Value = True Then
        strto = strto + Sheet2.Range("amarsua").Value + "; "
    ElseIf Sheet2.opt_CC9.Value = True Then
        strcc = strcc + Sheet2.Range("amarsua").Value + "; "
    ElseIf Sheet2.opt_N9.Value = True Then
    End If
    
    If Sheet2.opt_To10.Value = True Then
        strto = strto + Sheet2.Range("istolentino").Value + "; "
    ElseIf Sheet2.opt_CC10.Value = True Then
        strcc = strcc + Sheet2.Range("istolentino").Value + "; "
    ElseIf Sheet2.opt_N10.Value = True Then
    End If

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    strbody = ""
    If Status.lst_L.ListCount >= 2 Then
        strbody = strbody + "LUCLS" + vbCrLf
        For i = 1 To Status.lst_L.ListCount - 1
           strbody = strbody + "   " + Status.lst_L.List(i) + vbCrLf
        Next i
    End If
    If Status.lst_B.ListCount >= 2 Then
        strbody = strbody + "BCLS" + vbCrLf
        For i = 1 To Status.lst_B.ListCount - 1
            strbody = strbody + "   " + Status.lst_B.List(i) + vbCrLf
        Next i
    End If
    If Status.lst_DC.ListCount >= 2 Then
        strbody = strbody + "DCLS" + vbCrLf
        For i = 1 To Status.lst_DC.ListCount - 1
            strbody = strbody + "   " + Status.lst_DC.List(i) + vbCrLf
        Next i
    End If
    If Status.lst_DL.ListCount >= 2 Then
        strbody = strbody + "DILIMAN" + vbCrLf
        For i = 1 To Status.lst_DL.ListCount - 1
            strbody = strbody + "   " + Status.lst_DL.List(i) + vbCrLf
        Next i
    End If
    If Status.lst_G.ListCount >= 2 Then
        strbody = strbody + "GARNET" + vbCrLf
        For i = 1 To Status.lst_G.ListCount - 1
            strbody = strbody + "   " + Status.lst_G.List(i) + vbCrLf
        Next i
    End If
    If Status.lst_S.ListCount >= 2 Then
        strbody = strbody + "SAMPALOC" + vbCrLf
        For i = 1 To Status.lst_S.ListCount - 1
            strbody = strbody + "   " + Status.lst_S.List(i) + vbCrLf
        Next i
    End If
    If Status.lst_GH.ListCount >= 2 Then
        strbody = strbody + "GREENHILLS" + vbCrLf
        For i = 1 To Status.lst_GH.ListCount - 1
            strbody = strbody + "   " + Status.lst_GH.List(i) + vbCrLf
        Next i
    End If
    If Status.lst_C.ListCount >= 2 Then
        strbody = strbody + "CEBU" + vbCrLf
        For i = 1 To Status.lst_C.ListCount - 1
            strbody = strbody + "   " + Status.lst_C.List(i) + vbCrLf
        Next i
    End If

    Dim signature As String
    
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
        .Subject = "BC Gov IA Deadline"
        .Body = strbody
        .Attachments.Add (Sheet3.Range("A11").Value)
        .Attachments.Add Sheet3.Range("A5").Value, olByValue, 0
        .HTMLBody = "<style>body{color:red;font-weight:bold}</style><h2>DUE DATE NEAR EXPIRATION:</h2>" & .HTMLBody & Sheet3.Range("A2").Value & "<br><br>" & "<img src='cid:" & Sheet3.Range("A8").Value & "'>" & signature
        .Display
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub

Private Sub btn_ViewD_Click()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Dim strto As String
    Dim strcc As String
    Dim i As Integer
    
    If Sheet2.opt_To2.Value = True Then
        strto = strto + Sheet2.Range("rdvillasenor").Value + "; "
    ElseIf Sheet2.opt_CC2.Value = True Then
        strcc = strcc + Sheet2.Range("rdvillasenor").Value + "; "
    ElseIf Sheet2.opt_N2.Value = True Then
    End If
    
    If Sheet2.opt_To5.Value = True Then
        strto = strto + Sheet2.Range("janacpil").Value + "; "
    ElseIf Sheet2.opt_CC5.Value = True Then
        strcc = strcc + Sheet2.Range("janacpil").Value + "; "
    ElseIf Sheet2.opt_N5.Value = True Then
    End If
    
    If Sheet2.opt_To6.Value = True Then
        strto = strto + Sheet2.Range("hclim").Value + "; "
    ElseIf Sheet2.opt_CC6.Value = True Then
        strcc = strcc + Sheet2.Range("hclim").Value + "; "
    ElseIf Sheet2.opt_N6.Value = True Then
    End If
    
    If Status.MultiPage1.Value = 0 Then
        If Status.MultiPage1.Pages(0).Caption = "*LUCLS" Then
            If Sheet2.opt_ToL1.Value = True Then
                strto = strto + Sheet2.Range("mmginete").Value + "; "
            ElseIf Sheet2.opt_CCL1.Value = True Then
                strcc = strcc + Sheet2.Range("mmginete").Value + "; "
            ElseIf Sheet2.opt_NL1.Value = True Then
            End If
            
            If Sheet2.opt_ToL2.Value = True Then
                strto = strto + Sheet2.Range("jmjacinto").Value + "; "
            ElseIf Sheet2.opt_CCL2.Value = True Then
                strcc = strcc + Sheet2.Range("jmjacinto").Value + "; "
            ElseIf Sheet2.opt_NL2.Value = True Then
            End If
            
            If Sheet2.opt_ToL3.Value = True Then
                strto = strto + Sheet2.Range("vprodriguez").Value + "; "
            ElseIf Sheet2.opt_CCL3.Value = True Then
                strcc = strcc + Sheet2.Range("vprodriguez").Value + "; "
            ElseIf Sheet2.opt_NL3.Value = True Then
            End If
            
            If Sheet2.opt_ToL4.Value = True Then
                strto = strto + Sheet2.Range("ecganuelas").Value + "; "
            ElseIf Sheet2.opt_CCL4.Value = True Then
                strcc = strcc + Sheet2.Range("ecganuelas").Value + "; "
            ElseIf Sheet2.opt_NL4.Value = True Then
            End If
            
            If Sheet2.opt_ToL5.Value = True Then
                strto = strto + Sheet2.Range("mmmones").Value + "; "
            ElseIf Sheet2.opt_CCL5.Value = True Then
                strcc = strcc + Sheet2.Range("mmmones").Value + "; "
            ElseIf Sheet2.opt_NL5.Value = True Then
            End If
            
            If Sheet2.opt_ToL6.Value = True Then
                strto = strto + Sheet2.Range("ljechave").Value + "; "
            ElseIf Sheet2.opt_CCL6.Value = True Then
                strcc = strcc + Sheet2.Range("ljechave").Value + "; "
            ElseIf Sheet2.opt_NL6.Value = True Then
            End If
            
            If Sheet2.opt_ToL7.Value = True Then
                strto = strto + Sheet2.Range("ccmendoza").Value + "; "
            ElseIf Sheet2.opt_CCL7.Value = True Then
                strcc = strcc + Sheet2.Range("ccmendoza").Value + "; "
            ElseIf Sheet2.opt_NL7.Value = True Then
            End If
            
            If Sheet2.opt_ToL8.Value = True Then
                strto = strto + Sheet2.Range("jlmoldez").Value + "; "
            ElseIf Sheet2.opt_CCL8.Value = True Then
                strcc = strcc + Sheet2.Range("jlmoldez").Value + "; "
            ElseIf Sheet2.opt_NL8.Value = True Then
            End If
            
            If Sheet2.opt_ToL9.Value = True Then
                strto = strto + Sheet2.Range("jogutierrez").Value + "; "
            ElseIf Sheet2.opt_CCL9.Value = True Then
                strcc = strcc + Sheet2.Range("jogutierrez").Value + "; "
            ElseIf Sheet2.opt_NL9.Value = True Then
            End If

            If Sheet2.opt_ToCLS2.Value = True Then
                strto = strto + Sheet2.Range("emgacayan").Value + "; "
            ElseIf Sheet2.opt_CCCLS2.Value = True Then
                strcc = strcc + Sheet2.Range("emgacayan").Value + "; "
            ElseIf Sheet2.opt_NCLS2.Value = True Then
            End If
            
            strbody = ""
            strbody = strbody + "LUCLS" + vbCrLf
            For i = 1 To Status.lst_L.ListCount - 1
                If Status.lst_L.List(i) Like "*Off-site Documentation Audit*" Or Status.lst_L.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_L.List(i) + vbCrLf
                End If
            Next i
        End If
    End If
    
    If Status.MultiPage1.Value = 1 Then
        If Status.MultiPage1.Pages(1).Caption = "*BCLS" Then
            If Sheet2.opt_ToB1.Value = True Then
                strto = strto + Sheet2.Range("jdbustamante").Value + "; "
            ElseIf Sheet2.opt_CCB1.Value = True Then
                strcc = strcc + Sheet2.Range("jdbustamante").Value + "; "
            ElseIf Sheet2.opt_NB1.Value = True Then
            End If
            
            If Sheet2.opt_ToB2.Value = True Then
                strto = strto + Sheet2.Range("apcaringal").Value + "; "
            ElseIf Sheet2.opt_CCB2.Value = True Then
                strcc = strcc + Sheet2.Range("apcaringal").Value + "; "
            ElseIf Sheet2.opt_NB2.Value = True Then
            End If
            
            If Sheet2.opt_ToB3.Value = True Then
                strto = strto + Sheet2.Range("aacatibog").Value + "; "
            ElseIf Sheet2.opt_CCB3.Value = True Then
                strcc = strcc + Sheet2.Range("aacatibog").Value + "; "
            ElseIf Sheet2.opt_NB3.Value = True Then
            End If
            
            If Sheet2.opt_ToB4.Value = True Then
                strto = strto + Sheet2.Range("ebdeleon").Value + "; "
            ElseIf Sheet2.opt_CCB4.Value = True Then
                strcc = strcc + Sheet2.Range("ebdeleon").Value + "; "
            ElseIf Sheet2.opt_NB4.Value = True Then
            End If
            
            If Sheet2.opt_ToB5.Value = True Then
                strto = strto + Sheet2.Range("rjenriquez").Value + "; "
            ElseIf Sheet2.opt_CCB5.Value = True Then
                strcc = strcc + Sheet2.Range("rjenriquez").Value + "; "
            ElseIf Sheet2.opt_NB5.Value = True Then
            End If
            
            If Sheet2.opt_ToB6.Value = True Then
                strto = strto + Sheet2.Range("rjetcobanez").Value + "; "
            ElseIf Sheet2.opt_CCB6.Value = True Then
                strcc = strcc + Sheet2.Range("rjetcobanez").Value + "; "
            ElseIf Sheet2.opt_NB6.Value = True Then
            End If
            
            If Sheet2.opt_ToB7.Value = True Then
                strto = strto + Sheet2.Range("rpmanago").Value + "; "
            ElseIf Sheet2.opt_CCB7.Value = True Then
                strcc = strcc + Sheet2.Range("rpmanago").Value + "; "
            ElseIf Sheet2.opt_NB7.Value = True Then
            End If
            
            If Sheet2.opt_ToB8.Value = True Then
                strto = strto + Sheet2.Range("hdpilar").Value + "; "
            ElseIf Sheet2.opt_CCB8.Value = True Then
                strcc = strcc + Sheet2.Range("hdpilar").Value + "; "
            ElseIf Sheet2.opt_NB8.Value = True Then
            End If
            
            If Sheet2.opt_ToB9.Value = True Then
                strto = strto + Sheet2.Range("acreyes").Value + "; "
            ElseIf Sheet2.opt_CCB9.Value = True Then
                strcc = strcc + Sheet2.Range("acreyes").Value + "; "
            ElseIf Sheet2.opt_NB9.Value = True Then
            End If
            
            If Sheet2.opt_ToB10.Value = True Then
                strto = strto + Sheet2.Range("aevizconde").Value + "; "
            ElseIf Sheet2.opt_CCB10.Value = True Then
                strcc = strcc + Sheet2.Range("aevizconde").Value + "; "
            ElseIf Sheet2.opt_NB10.Value = True Then
            End If

            If Sheet2.opt_ToCLS2.Value = True Then
                strto = strto + Sheet2.Range("emgacayan").Value + "; "
            ElseIf Sheet2.opt_CCCLS2.Value = True Then
                strcc = strcc + Sheet2.Range("emgacayan").Value + "; "
            ElseIf Sheet2.opt_NCLS2.Value = True Then
            End If
            
            
            strbody = ""
            strbody = strbody + "BCLS" + vbCrLf
            For i = 1 To Status.lst_B.ListCount - 1
                If Status.lst_B.List(i) Like "*Off-site Documentation Audit*" Or Status.lst_B.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_B.List(i) + vbCrLf
                End If
            Next i
        End If
    End If
    
    If Status.MultiPage1.Value = 2 Then
        If Status.MultiPage1.Pages(2).Caption = "*DCLS" Then
            If Sheet2.opt_ToDC1.Value = True Then
                strto = strto + Sheet2.Range("vvpunzalan").Value + "; "
            ElseIf Sheet2.opt_CCDC1.Value = True Then
                strcc = strcc + Sheet2.Range("vvpunzalan").Value + "; "
            ElseIf Sheet2.opt_NDC1.Value = True Then
            End If
            
            If Sheet2.opt_ToDC2.Value = True Then
                strto = strto + Sheet2.Range("neromero").Value + "; "
            ElseIf Sheet2.opt_CCDC2.Value = True Then
                strcc = strcc + Sheet2.Range("neromero").Value + "; "
            ElseIf Sheet2.opt_NDC2.Value = True Then
            End If
            
            If Sheet2.opt_ToDC3.Value = True Then
                strto = strto + Sheet2.Range("dltatad").Value + "; "
            ElseIf Sheet2.opt_CCDC3.Value = True Then
                strcc = strcc + Sheet2.Range("dltatad").Value + "; "
            ElseIf Sheet2.opt_NDC3.Value = True Then
            End If
            
            If Sheet2.opt_ToDC4.Value = True Then
                strto = strto + Sheet2.Range("josalvador").Value + "; "
            ElseIf Sheet2.opt_CCDC4.Value = True Then
                strcc = strcc + Sheet2.Range("josalvador").Value + "; "
            ElseIf Sheet2.opt_NDC4.Value = True Then
            End If
            
            If Sheet2.opt_ToDC5.Value = True Then
                strto = strto + Sheet2.Range("mbdevilla").Value + "; "
            ElseIf Sheet2.opt_CCDC5.Value = True Then
                strcc = strcc + Sheet2.Range("mbdevilla").Value + "; "
            ElseIf Sheet2.opt_NDC5.Value = True Then
            End If
            
            If Sheet2.opt_ToDC6.Value = True Then
                strto = strto + Sheet2.Range("rspanta").Value + "; "
            ElseIf Sheet2.opt_CCDC6.Value = True Then
                strcc = strcc + Sheet2.Range("rspanta").Value + "; "
            ElseIf Sheet2.opt_NDC6.Value = True Then
            End If
            
            If Sheet2.opt_ToDC7.Value = True Then
                strto = strto + Sheet2.Range("jlandicoy").Value + "; "
            ElseIf Sheet2.opt_CCDC7.Value = True Then
                strcc = strcc + Sheet2.Range("jlandicoy").Value + "; "
            ElseIf Sheet2.opt_NDC7.Value = True Then
            End If

            If Sheet2.opt_ToCLS2.Value = True Then
                strto = strto + Sheet2.Range("emgacayan").Value + "; "
            ElseIf Sheet2.opt_CCCLS2.Value = True Then
                strcc = strcc + Sheet2.Range("emgacayan").Value + "; "
            ElseIf Sheet2.opt_NCLS2.Value = True Then
            End If
                      
            strbody = ""
            strbody = strbody + "DCLS" + vbCrLf
            For i = 1 To Status.lst_DC.ListCount - 1
                If Status.lst_DC.List(i) Like "*Off-site Documentation Audit*" Or Status.lst_DC.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_DC.List(i) + vbCrLf
                End If
            Next i
        End If
    End If
    
    If Status.MultiPage1.Value = 3 Then
        If Status.MultiPage1.Pages(3).Caption = "*DILIMAN" And chk_APFMDL.Value = False Then
            If Sheet2.opt_ToDL1.Value = True Then
                strto = strto + Sheet2.Range("aoabary").Value + "; "
            ElseIf Sheet2.opt_CCDL1.Value = True Then
                strcc = strcc + Sheet2.Range("aoabary").Value + "; "
            ElseIf Sheet2.opt_NDL1.Value = True Then
            End If
            
            If Sheet2.opt_ToDL2.Value = True Then
                strto = strto + Sheet2.Range("rogbanez").Value + "; "
            ElseIf Sheet2.opt_CCDL2.Value = True Then
                strcc = strcc + Sheet2.Range("rogbanez").Value + "; "
            ElseIf Sheet2.opt_NDL2.Value = True Then
            End If
            
            If Sheet2.opt_ToDL3.Value = True Then
                strto = strto + Sheet2.Range("blbautista").Value + "; "
            ElseIf Sheet2.opt_CCDL3.Value = True Then
                strcc = strcc + Sheet2.Range("blbautista").Value + "; "
            ElseIf Sheet2.opt_NDL3.Value = True Then
            End If
            
            If Sheet2.opt_ToDL4.Value = True Then
                strto = strto + Sheet2.Range("ebbayle").Value + "; "
            ElseIf Sheet2.opt_CCDL4.Value = True Then
                strcc = strcc + Sheet2.Range("ebbayle").Value + "; "
            ElseIf Sheet2.opt_NDL4.Value = True Then
            End If
            
            If Sheet2.opt_ToDL6.Value = True Then
                strto = strto + Sheet2.Range("accruz").Value + "; "
            ElseIf Sheet2.opt_CCDL6.Value = True Then
                strcc = strcc + Sheet2.Range("accruz").Value + "; "
            ElseIf Sheet2.opt_NDL6.Value = True Then
            End If
            
            If Sheet2.opt_ToDL7.Value = True Then
                strto = strto + Sheet2.Range("mgdioso").Value + "; "
            ElseIf Sheet2.opt_CCDL7.Value = True Then
                strcc = strcc + Sheet2.Range("mgdioso").Value + "; "
            ElseIf Sheet2.opt_NDL7.Value = True Then
            End If
            
            If Sheet2.opt_ToDL8.Value = True Then
                strto = strto + Sheet2.Range("amestrella").Value + "; "
            ElseIf Sheet2.opt_CCDL8.Value = True Then
                strcc = strcc + Sheet2.Range("amestrella").Value + "; "
            ElseIf Sheet2.opt_NDL8.Value = True Then
            End If
            
            If Sheet2.opt_ToDL9.Value = True Then
                strto = strto + Sheet2.Range("drfonacier").Value + "; "
            ElseIf Sheet2.opt_CCDL9.Value = True Then
                strcc = strcc + Sheet2.Range("drfonacier").Value + "; "
            ElseIf Sheet2.opt_NDL9.Value = True Then
            End If
            
            If Sheet2.opt_ToDL10.Value = True Then
                strto = strto + Sheet2.Range("rjlim").Value + "; "
            ElseIf Sheet2.opt_CCDL10.Value = True Then
                strcc = strcc + Sheet2.Range("rjlim").Value + "; "
            ElseIf Sheet2.opt_NDL10.Value = True Then
            End If
            
            If Sheet2.opt_ToDL11.Value = True Then
                strto = strto + Sheet2.Range("ecmadrilejo").Value + "; "
            ElseIf Sheet2.opt_CCDL11.Value = True Then
                strcc = strcc + Sheet2.Range("ecmadrilejo").Value + "; "
            ElseIf Sheet2.opt_NDL11.Value = True Then
            End If
            
            If Sheet2.opt_ToDL12.Value = True Then
                strto = strto + Sheet2.Range("jvnaval").Value + "; "
            ElseIf Sheet2.opt_CCDL12.Value = True Then
                strcc = strcc + Sheet2.Range("jvnaval").Value + "; "
            ElseIf Sheet2.opt_NDL12.Value = True Then
            End If
            
            If Sheet2.opt_ToDL13.Value = True Then
                strto = strto + Sheet2.Range("wmsabile").Value + "; "
            ElseIf Sheet2.opt_CCDL13.Value = True Then
                strcc = strcc + Sheet2.Range("wmsabile").Value + "; "
            ElseIf Sheet2.opt_NDL13.Value = True Then
            End If
            
            If Sheet2.opt_ToDL14.Value = True Then
                strto = strto + Sheet2.Range("agsoro").Value + "; "
            ElseIf Sheet2.opt_CCDL14.Value = True Then
                strcc = strcc + Sheet2.Range("agsoro").Value + "; "
            ElseIf Sheet2.opt_NDL14.Value = True Then
            End If

            If Sheet2.opt_ToDLS1.Value = True Then
                strto = strto + Sheet2.Range("csalejo").Value + "; "
            ElseIf Sheet2.opt_CCDLS1.Value = True Then
                strcc = strcc + Sheet2.Range("csalejo").Value + "; "
            ElseIf Sheet2.opt_NDLS1.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS3.Value = True Then
                strto = strto + Sheet2.Range("ecbuera").Value + "; "
            ElseIf Sheet2.opt_CCDLS3.Value = True Then
                strcc = strcc + Sheet2.Range("ecbuera").Value + "; "
            ElseIf Sheet2.opt_NDLS3.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS4.Value = True Then
                strto = strto + Sheet2.Range("kdcruz").Value + "; "
            ElseIf Sheet2.opt_CCDLS4.Value = True Then
                strcc = strcc + Sheet2.Range("kdcruz").Value + "; "
            ElseIf Sheet2.opt_NDLS4.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS5.Value = True Then
                strto = strto + Sheet2.Range("wdgdebelen").Value + "; "
            ElseIf Sheet2.opt_CCDLS5.Value = True Then
                strcc = strcc + Sheet2.Range("wdgdebelen").Value + "; "
            ElseIf Sheet2.opt_NDLS5.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS6.Value = True Then
                strto = strto + Sheet2.Range("nmestevez").Value + "; "
            ElseIf Sheet2.opt_CCDLS6.Value = True Then
                strcc = strcc + Sheet2.Range("nmestevez").Value + "; "
            ElseIf Sheet2.opt_NDLS6.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS7.Value = True Then
                strto = strto + Sheet2.Range("vdgrino").Value + "; "
            ElseIf Sheet2.opt_CCDLS7.Value = True Then
                strcc = strcc + Sheet2.Range("vdgrino").Value + "; "
            ElseIf Sheet2.opt_NDLS7.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS8.Value = True Then
                strto = strto + Sheet2.Range("jmhernandez").Value + "; "
            ElseIf Sheet2.opt_CCDLS8.Value = True Then
                strcc = strcc + Sheet2.Range("jmhernandez").Value + "; "
            ElseIf Sheet2.opt_NDLS8.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS9.Value = True Then
                strto = strto + Sheet2.Range("anmantes").Value + "; "
            ElseIf Sheet2.opt_CCDLS9.Value = True Then
                strcc = strcc + Sheet2.Range("anmantes").Value + "; "
            ElseIf Sheet2.opt_NDLS9.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS10.Value = True Then
                strto = strto + Sheet2.Range("lgpagay").Value + "; "
            ElseIf Sheet2.opt_CCDLS10.Value = True Then
                strcc = strcc + Sheet2.Range("lgpagay").Value + "; "
            ElseIf Sheet2.opt_NDLS10.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS11.Value = True Then
                strto = strto + Sheet2.Range("rjtorio").Value + "; "
            ElseIf Sheet2.opt_CCDLS11.Value = True Then
                strcc = strcc + Sheet2.Range("rjtorio").Value + "; "
            ElseIf Sheet2.opt_NDLS11.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS12.Value = True Then
                strto = strto + Sheet2.Range("vauypala").Value + "; "
            ElseIf Sheet2.opt_CCDLS12.Value = True Then
                strcc = strcc + Sheet2.Range("vauypala").Value + "; "
            ElseIf Sheet2.opt_NDLS12.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS13.Value = True Then
                strto = strto + Sheet2.Range("rdvizmanos").Value + "; "
            ElseIf Sheet2.opt_CCDLS13.Value = True Then
                strcc = strcc + Sheet2.Range("rdvizmanos").Value + "; "
            ElseIf Sheet2.opt_NDLS13.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS.Value = True Then
                strto = strto + Sheet2.Range("asgaba").Value + "; "
            ElseIf Sheet2.opt_CCDLS.Value = True Then
                strcc = strcc + Sheet2.Range("asgaba").Value + "; "
            ElseIf Sheet2.opt_NDLS.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO1.Value = True Then
                strto = strto + Sheet2.Range("fcbaul").Value + "; "
            ElseIf Sheet2.opt_CCFCNO1.Value = True Then
                strcc = strcc + Sheet2.Range("fcbaul").Value + "; "
            ElseIf Sheet2.opt_NFCNO1.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO2.Value = True Then
                strto = strto + Sheet2.Range("picables").Value + "; "
            ElseIf Sheet2.opt_CCFCNO2.Value = True Then
                strcc = strcc + Sheet2.Range("picables").Value + "; "
            ElseIf Sheet2.opt_NFCNO2.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO3.Value = True Then
                strto = strto + Sheet2.Range("aljimenez").Value + "; "
            ElseIf Sheet2.opt_CCFCNO3.Value = True Then
                strcc = strcc + Sheet2.Range("aljimenez").Value + "; "
            ElseIf Sheet2.opt_NFCNO3.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO4.Value = True Then
                strto = strto + Sheet2.Range("rtlampa").Value + "; "
            ElseIf Sheet2.opt_CCFCNO4.Value = True Then
                strcc = strcc + Sheet2.Range("rtlampa").Value + "; "
            ElseIf Sheet2.opt_NFCNO4.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO5.Value = True Then
                strto = strto + Sheet2.Range("oalinco").Value + "; "
            ElseIf Sheet2.opt_CCFCNO5.Value = True Then
                strcc = strcc + Sheet2.Range("oalinco").Value + "; "
            ElseIf Sheet2.opt_NFCNO5.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO6.Value = True Then
                strto = strto + Sheet2.Range("mfsantos").Value + "; "
            ElseIf Sheet2.opt_CCFCNO6.Value = True Then
                strcc = strcc + Sheet2.Range("mfsantos").Value + "; "
            ElseIf Sheet2.opt_NFCNO6.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO.Value = True Then
                strto = strto + Sheet2.Range("afcapiral").Value + "; "
            ElseIf Sheet2.opt_CCFCNO.Value = True Then
                strcc = strcc + Sheet2.Range("afcapiral").Value + "; "
            ElseIf Sheet2.opt_NFCNO.Value = True Then
            End If
                      
            strbody = ""
            strbody = strbody + "DILIMAN" + vbCrLf
            For i = 1 To Status.lst_DL.ListCount - 1
                If Status.lst_DL.List(i) Like "*Off-site Documentation Audit*" Or Status.lst_DL.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_DL.List(i) + vbCrLf
                End If
            Next i
        End If
    End If
    
    If Status.MultiPage1.Value = 4 Or Status.MultiPage1.Value = 6 Then
        If Status.MultiPage1.Pages(4).Caption = "*GARNET" Or Status.MultiPage1.Pages(6).Caption = "*GREENHILLS" Then
            If Sheet2.opt_ToGGH1.Value = True Then
                strto = strto + Sheet2.Range("ntalcantara").Value + "; "
            ElseIf Sheet2.opt_CCGGH1.Value = True Then
                strcc = strcc + Sheet2.Range("ntalcantara").Value + "; "
            ElseIf Sheet2.opt_NGGH1.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH2.Value = True Then
                strto = strto + Sheet2.Range("lpabapo").Value + "; "
            ElseIf Sheet2.opt_CCGGH2.Value = True Then
                strcc = strcc + Sheet2.Range("lpabapo").Value + "; "
            ElseIf Sheet2.opt_NGGH2.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH3.Value = True Then
                strto = strto + Sheet2.Range("raaquino").Value + "; "
            ElseIf Sheet2.opt_CCGGH3.Value = True Then
                strcc = strcc + Sheet2.Range("raaquino").Value + "; "
            ElseIf Sheet2.opt_NGGH3.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH5.Value = True Then
                strto = strto + Sheet2.Range("rrcerdon").Value + "; "
            ElseIf Sheet2.opt_CCGGH5.Value = True Then
                strcc = strcc + Sheet2.Range("rrcerdon").Value + "; "
            ElseIf Sheet2.opt_NGGH5.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH6.Value = True Then
                strto = strto + Sheet2.Range("lmcoma").Value + "; "
            ElseIf Sheet2.opt_CCGGH6.Value = True Then
                strcc = strcc + Sheet2.Range("lmcoma").Value + "; "
            ElseIf Sheet2.opt_NGGH6.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH7.Value = True Then
                strto = strto + Sheet2.Range("radeguzman").Value + "; "
            ElseIf Sheet2.opt_CCGGH7.Value = True Then
                strcc = strcc + Sheet2.Range("radeguzman").Value + "; "
            ElseIf Sheet2.opt_NGGH7.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH9.Value = True Then
                strto = strto + Sheet2.Range("nsreas").Value + "; "
            ElseIf Sheet2.opt_CCGGH9.Value = True Then
                strcc = strcc + Sheet2.Range("nsreas").Value + "; "
            ElseIf Sheet2.opt_NGGH9.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH10.Value = True Then
                strto = strto + Sheet2.Range("jugabriel").Value + "; "
            ElseIf Sheet2.opt_CCGGH10.Value = True Then
                strcc = strcc + Sheet2.Range("jugabriel").Value + "; "
            ElseIf Sheet2.opt_NGGH10.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH11.Value = True Then
                strto = strto + Sheet2.Range("mbgaspay").Value + "; "
            ElseIf Sheet2.opt_CCGGH11.Value = True Then
                strcc = strcc + Sheet2.Range("mbgaspay").Value + "; "
            ElseIf Sheet2.opt_NGGH11.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH12.Value = True Then
                strto = strto + Sheet2.Range("cgkatalbas").Value + "; "
            ElseIf Sheet2.opt_CCGGH12.Value = True Then
                strcc = strcc + Sheet2.Range("cgkatalbas").Value + "; "
            ElseIf Sheet2.opt_NGGH12.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH13.Value = True Then
                strto = strto + Sheet2.Range("nrlapus").Value + "; "
            ElseIf Sheet2.opt_CCGGH13.Value = True Then
                strcc = strcc + Sheet2.Range("nrlapus").Value + "; "
            ElseIf Sheet2.opt_NGGH13.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH14.Value = True Then
                strto = strto + Sheet2.Range("fslizardo").Value + "; "
            ElseIf Sheet2.opt_CCGGH14.Value = True Then
                strcc = strcc + Sheet2.Range("fslizardo").Value + "; "
            ElseIf Sheet2.opt_NGGH14.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH15.Value = True Then
                strto = strto + Sheet2.Range("drparrocha").Value + "; "
            ElseIf Sheet2.opt_CCGGH15.Value = True Then
                strcc = strcc + Sheet2.Range("drparrocha").Value + "; "
            ElseIf Sheet2.opt_NGGH15.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH16.Value = True Then
                strto = strto + Sheet2.Range("jdsarinas").Value + "; "
            ElseIf Sheet2.opt_CCGGH16.Value = True Then
                strcc = strcc + Sheet2.Range("jdsarinas").Value + "; "
            ElseIf Sheet2.opt_NGGH16.Value = True Then
            End If
            
            If Sheet2.opt_ToGGH17.Value = True Then
                strto = strto + Sheet2.Range("rteves").Value + "; "
            ElseIf Sheet2.opt_CCGGH17.Value = True Then
                strcc = strcc + Sheet2.Range("rteves").Value + "; "
            ElseIf Sheet2.opt_NGGH17.Value = True Then
            End If

            If Sheet2.opt_ToFCNO1.Value = True Then
                strto = strto + Sheet2.Range("fcbaul").Value + "; "
            ElseIf Sheet2.opt_CCFCNO1.Value = True Then
                strcc = strcc + Sheet2.Range("fcbaul").Value + "; "
            ElseIf Sheet2.opt_NFCNO1.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO2.Value = True Then
                strto = strto + Sheet2.Range("picables").Value + "; "
            ElseIf Sheet2.opt_CCFCNO2.Value = True Then
                strcc = strcc + Sheet2.Range("picables").Value + "; "
            ElseIf Sheet2.opt_NFCNO2.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO3.Value = True Then
                strto = strto + Sheet2.Range("aljimenez").Value + "; "
            ElseIf Sheet2.opt_CCFCNO3.Value = True Then
                strcc = strcc + Sheet2.Range("aljimenez").Value + "; "
            ElseIf Sheet2.opt_NFCNO3.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO4.Value = True Then
                strto = strto + Sheet2.Range("rtlampa").Value + "; "
            ElseIf Sheet2.opt_CCFCNO4.Value = True Then
                strcc = strcc + Sheet2.Range("rtlampa").Value + "; "
            ElseIf Sheet2.opt_NFCNO4.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO5.Value = True Then
                strto = strto + Sheet2.Range("oalinco").Value + "; "
            ElseIf Sheet2.opt_CCFCNO5.Value = True Then
                strcc = strcc + Sheet2.Range("oalinco").Value + "; "
            ElseIf Sheet2.opt_NFCNO5.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO6.Value = True Then
                strto = strto + Sheet2.Range("mfsantos").Value + "; "
            ElseIf Sheet2.opt_CCFCNO6.Value = True Then
                strcc = strcc + Sheet2.Range("mfsantos").Value + "; "
            ElseIf Sheet2.opt_NFCNO6.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO7.Value = True Then
                strto = strto + Sheet2.Range("aeenano").Value + "; "
            ElseIf Sheet2.opt_CCFCNO7.Value = True Then
                strcc = strcc + Sheet2.Range("aeenano").Value + "; "
            ElseIf Sheet2.opt_NFCNO7.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO8.Value = True Then
                strto = strto + Sheet2.Range("apinciong").Value + "; "
            ElseIf Sheet2.opt_CCFCNO8.Value = True Then
                strcc = strcc + Sheet2.Range("apinciong").Value + "; "
            ElseIf Sheet2.opt_NFCNO8.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO9.Value = True Then
                strto = strto + Sheet2.Range("eanieva").Value + "; "
            ElseIf Sheet2.opt_CCFCNO9.Value = True Then
                strcc = strcc + Sheet2.Range("eanieva").Value + "; "
            ElseIf Sheet2.opt_NFCNO9.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO10.Value = True Then
                strto = strto + Sheet2.Range("jutamondong").Value + "; "
            ElseIf Sheet2.opt_CCFCNO10.Value = True Then
                strcc = strcc + Sheet2.Range("jutamondong").Value + "; "
            ElseIf Sheet2.opt_NFCNO10.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO11.Value = True Then
                strto = strto + Sheet2.Range("rcroque").Value + "; "
            ElseIf Sheet2.opt_CCFCNO11.Value = True Then
                strcc = strcc + Sheet2.Range("rcroque").Value + "; "
            ElseIf Sheet2.opt_NFCNO11.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO12.Value = True Then
                strto = strto + Sheet2.Range("mosena").Value + "; "
            ElseIf Sheet2.opt_CCFCNO12.Value = True Then
                strcc = strcc + Sheet2.Range("mosena").Value + "; "
            ElseIf Sheet2.opt_NFCNO12.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO.Value = True Then
                strto = strto + Sheet2.Range("afcapiral").Value + "; "
            ElseIf Sheet2.opt_CCFCNO.Value = True Then
                strcc = strcc + Sheet2.Range("afcapiral").Value + "; "
            ElseIf Sheet2.opt_NFCNO.Value = True Then
            End If
                      
            strbody = ""
            strbody = strbody + "GARNET" + vbCrLf
            For i = 1 To Status.lst_G.ListCount - 1
                If Status.lst_G.List(i) Like "*Off-site Documentation Audit*" Or Status.lst_G.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_G.List(i) + vbCrLf
                End If
            Next i
            strbody = strbody + "GREENHILLS" + vbCrLf
            For i = 1 To Status.lst_GH.ListCount - 1
                If Status.lst_GH.List(i) Like "*Off-site Documentation Audit*" Or Status.lst_GH.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_GH.List(i) + vbCrLf
                End If
            Next i
        End If
    End If
    
    If Status.MultiPage1.Value = 5 Then
        If Status.MultiPage1.Pages(5).Caption = "*SAMPALOC" Then
            If Sheet2.opt_ToS1.Value = True Then
                strto = strto + Sheet2.Range("emalsol").Value + "; "
            ElseIf Sheet2.opt_CCS1.Value = True Then
                strcc = strcc + Sheet2.Range("emalsol").Value + "; "
            ElseIf Sheet2.opt_NS1.Value = True Then
            End If
            
            If Sheet2.opt_ToS2.Value = True Then
                strto = strto + Sheet2.Range("aaagbayani").Value + "; "
            ElseIf Sheet2.opt_CCS2.Value = True Then
                strcc = strcc + Sheet2.Range("aaagbayani").Value + "; "
            ElseIf Sheet2.opt_NS2.Value = True Then
            End If
            
            If Sheet2.opt_ToS3.Value = True Then
                strto = strto + Sheet2.Range("gpaquino").Value + "; "
            ElseIf Sheet2.opt_CCS3.Value = True Then
                strcc = strcc + Sheet2.Range("gpaquino").Value + "; "
            ElseIf Sheet2.opt_NS3.Value = True Then
            End If
            
            If Sheet2.opt_ToS4.Value = True Then
                strto = strto + Sheet2.Range("rhatendido").Value + "; "
            ElseIf Sheet2.opt_CCS4.Value = True Then
                strcc = strcc + Sheet2.Range("rhatendido").Value + "; "
            ElseIf Sheet2.opt_NS4.Value = True Then
            End If
            
            If Sheet2.opt_ToS7.Value = True Then
                strto = strto + Sheet2.Range("nnelegado").Value + "; "
            ElseIf Sheet2.opt_CCS7.Value = True Then
                strcc = strcc + Sheet2.Range("nnelegado").Value + "; "
            ElseIf Sheet2.opt_NS7.Value = True Then
            End If
            
            If Sheet2.opt_ToS9.Value = True Then
                strto = strto + Sheet2.Range("almgonzales").Value + "; "
            ElseIf Sheet2.opt_CCS9.Value = True Then
                strcc = strcc + Sheet2.Range("almgonzales").Value + "; "
            ElseIf Sheet2.opt_NS9.Value = True Then
            End If
            
            If Sheet2.opt_ToS11.Value = True Then
                strto = strto + Sheet2.Range("rvhundana").Value + "; "
            ElseIf Sheet2.opt_CCS11.Value = True Then
                strcc = strcc + Sheet2.Range("rvhundana").Value + "; "
            ElseIf Sheet2.opt_NS11.Value = True Then
            End If
            
            If Sheet2.opt_ToS12.Value = True Then
                strto = strto + Sheet2.Range("famacadaeg").Value + "; "
            ElseIf Sheet2.opt_CCS12.Value = True Then
                strcc = strcc + Sheet2.Range("famacadaeg").Value + "; "
            ElseIf Sheet2.opt_NS12.Value = True Then
            End If
            
            If Sheet2.opt_ToS13.Value = True Then
                strto = strto + Sheet2.Range("rsmariano").Value + "; "
            ElseIf Sheet2.opt_CCS13.Value = True Then
                strcc = strcc + Sheet2.Range("rsmariano").Value + "; "
            ElseIf Sheet2.opt_NS13.Value = True Then
            End If
            
            If Sheet2.opt_ToS14.Value = True Then
                strto = strto + Sheet2.Range("aenito").Value + "; "
            ElseIf Sheet2.opt_CCS14.Value = True Then
                strcc = strcc + Sheet2.Range("aenito").Value + "; "
            ElseIf Sheet2.opt_NS14.Value = True Then
            End If
            
            If Sheet2.opt_ToS15.Value = True Then
                strto = strto + Sheet2.Range("basoriano").Value + "; "
            ElseIf Sheet2.opt_CCS15.Value = True Then
                strcc = strcc + Sheet2.Range("basoriano").Value + "; "
            ElseIf Sheet2.opt_NS15.Value = True Then
            End If

            If Sheet2.opt_ToDLS1.Value = True Then
                strto = strto + Sheet2.Range("csalejo").Value + "; "
            ElseIf Sheet2.opt_CCDLS1.Value = True Then
                strcc = strcc + Sheet2.Range("csalejo").Value + "; "
            ElseIf Sheet2.opt_NDLS1.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS3.Value = True Then
                strto = strto + Sheet2.Range("ecbuera").Value + "; "
            ElseIf Sheet2.opt_CCDLS3.Value = True Then
                strcc = strcc + Sheet2.Range("ecbuera").Value + "; "
            ElseIf Sheet2.opt_NDLS3.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS4.Value = True Then
                strto = strto + Sheet2.Range("kdcruz").Value + "; "
            ElseIf Sheet2.opt_CCDLS4.Value = True Then
                strcc = strcc + Sheet2.Range("kdcruz").Value + "; "
            ElseIf Sheet2.opt_NDLS4.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS5.Value = True Then
                strto = strto + Sheet2.Range("wdgdebelen").Value + "; "
            ElseIf Sheet2.opt_CCDLS5.Value = True Then
                strcc = strcc + Sheet2.Range("wdgdebelen").Value + "; "
            ElseIf Sheet2.opt_NDLS5.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS6.Value = True Then
                strto = strto + Sheet2.Range("nmestevez").Value + "; "
            ElseIf Sheet2.opt_CCDLS6.Value = True Then
                strcc = strcc + Sheet2.Range("nmestevez").Value + "; "
            ElseIf Sheet2.opt_NDLS6.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS7.Value = True Then
                strto = strto + Sheet2.Range("vdgrino").Value + "; "
            ElseIf Sheet2.opt_CCDLS7.Value = True Then
                strcc = strcc + Sheet2.Range("vdgrino").Value + "; "
            ElseIf Sheet2.opt_NDLS7.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS8.Value = True Then
                strto = strto + Sheet2.Range("jmhernandez").Value + "; "
            ElseIf Sheet2.opt_CCDLS8.Value = True Then
                strcc = strcc + Sheet2.Range("jmhernandez").Value + "; "
            ElseIf Sheet2.opt_NDLS8.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS9.Value = True Then
                strto = strto + Sheet2.Range("anmantes").Value + "; "
            ElseIf Sheet2.opt_CCDLS9.Value = True Then
                strcc = strcc + Sheet2.Range("anmantes").Value + "; "
            ElseIf Sheet2.opt_NDLS9.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS10.Value = True Then
                strto = strto + Sheet2.Range("lgpagay").Value + "; "
            ElseIf Sheet2.opt_CCDLS10.Value = True Then
                strcc = strcc + Sheet2.Range("lgpagay").Value + "; "
            ElseIf Sheet2.opt_NDLS10.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS11.Value = True Then
                strto = strto + Sheet2.Range("rjtorio").Value + "; "
            ElseIf Sheet2.opt_CCDLS11.Value = True Then
                strcc = strcc + Sheet2.Range("rjtorio").Value + "; "
            ElseIf Sheet2.opt_NDLS11.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS12.Value = True Then
                strto = strto + Sheet2.Range("vauypala").Value + "; "
            ElseIf Sheet2.opt_CCDLS12.Value = True Then
                strcc = strcc + Sheet2.Range("vauypala").Value + "; "
            ElseIf Sheet2.opt_NDLS12.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS13.Value = True Then
                strto = strto + Sheet2.Range("rdvizmanos").Value + "; "
            ElseIf Sheet2.opt_CCDLS13.Value = True Then
                strcc = strcc + Sheet2.Range("rdvizmanos").Value + "; "
            ElseIf Sheet2.opt_NDLS13.Value = True Then
            End If
            
            If Sheet2.opt_ToDLS.Value = True Then
                strto = strto + Sheet2.Range("asgaba").Value + "; "
            ElseIf Sheet2.opt_CCDLS.Value = True Then
                strcc = strcc + Sheet2.Range("asgaba").Value + "; "
            ElseIf Sheet2.opt_NDLS.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO1.Value = True Then
                strto = strto + Sheet2.Range("fcbaul").Value + "; "
            ElseIf Sheet2.opt_CCFCNO1.Value = True Then
                strcc = strcc + Sheet2.Range("fcbaul").Value + "; "
            ElseIf Sheet2.opt_NFCNO1.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO2.Value = True Then
                strto = strto + Sheet2.Range("picables").Value + "; "
            ElseIf Sheet2.opt_CCFCNO2.Value = True Then
                strcc = strcc + Sheet2.Range("picables").Value + "; "
            ElseIf Sheet2.opt_NFCNO2.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO3.Value = True Then
                strto = strto + Sheet2.Range("aljimenez").Value + "; "
            ElseIf Sheet2.opt_CCFCNO3.Value = True Then
                strcc = strcc + Sheet2.Range("aljimenez").Value + "; "
            ElseIf Sheet2.opt_NFCNO3.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO4.Value = True Then
                strto = strto + Sheet2.Range("rtlampa").Value + "; "
            ElseIf Sheet2.opt_CCFCNO4.Value = True Then
                strcc = strcc + Sheet2.Range("rtlampa").Value + "; "
            ElseIf Sheet2.opt_NFCNO4.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO5.Value = True Then
                strto = strto + Sheet2.Range("oalinco").Value + "; "
            ElseIf Sheet2.opt_CCFCNO5.Value = True Then
                strcc = strcc + Sheet2.Range("oalinco").Value + "; "
            ElseIf Sheet2.opt_NFCNO5.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO6.Value = True Then
                strto = strto + Sheet2.Range("mfsantos").Value + "; "
            ElseIf Sheet2.opt_CCFCNO6.Value = True Then
                strcc = strcc + Sheet2.Range("mfsantos").Value + "; "
            ElseIf Sheet2.opt_NFCNO6.Value = True Then
            End If
            
            If Sheet2.opt_ToFCNO.Value = True Then
                strto = strto + Sheet2.Range("afcapiral").Value + "; "
            ElseIf Sheet2.opt_CCFCNO.Value = True Then
                strcc = strcc + Sheet2.Range("afcapiral").Value + "; "
            ElseIf Sheet2.opt_NFCNO.Value = True Then
            End If
                      
            strbody = ""
            strbody = strbody + "SAMPALOC" + vbCrLf
            For i = 1 To Status.lst_S.ListCount - 1
                If Status.lst_S.List(i) Like "*Off-site Documentation Audit*" Or Status.lst_S.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_S.List(i) + vbCrLf
                End If
            Next i
        End If
    End If
    
    If Status.MultiPage1.Value = 7 Then
        If Status.MultiPage1.Pages(7).Caption = "*CEBU" And chk_APFMC.Value = False Then
            If Sheet2.opt_ToC1.Value = True Then
                strto = strto + Sheet2.Range("emallosada").Value + "; "
            ElseIf Sheet2.opt_CCC1.Value = True Then
                strcc = strcc + Sheet2.Range("emallosada").Value + "; "
            ElseIf Sheet2.opt_NC1.Value = True Then
            End If
            
            If Sheet2.opt_ToC2.Value = True Then
                strto = strto + Sheet2.Range("rbarias").Value + "; "
            ElseIf Sheet2.opt_CCC2.Value = True Then
                strcc = strcc + Sheet2.Range("rbarias").Value + "; "
            ElseIf Sheet2.opt_NC2.Value = True Then
            End If
            
            If Sheet2.opt_ToC3.Value = True Then
                strto = strto + Sheet2.Range("aabaguio").Value + "; "
            ElseIf Sheet2.opt_CCC3.Value = True Then
                strcc = strcc + Sheet2.Range("aabaguio").Value + "; "
            ElseIf Sheet2.opt_NC3.Value = True Then
            End If
            
            If Sheet2.opt_ToC4.Value = True Then
                strto = strto + Sheet2.Range("lacabigas").Value + "; "
            ElseIf Sheet2.opt_CCC4.Value = True Then
                strcc = strcc + Sheet2.Range("lacabigas").Value + "; "
            ElseIf Sheet2.opt_NC4.Value = True Then
            End If
            
            If Sheet2.opt_ToC5.Value = True Then
                strto = strto + Sheet2.Range("jjcarumba").Value + "; "
            ElseIf Sheet2.opt_CCC5.Value = True Then
                strcc = strcc + Sheet2.Range("jjcarumba").Value + "; "
            ElseIf Sheet2.opt_NC5.Value = True Then
            End If
            
            If Sheet2.opt_ToC6.Value = True Then
                strto = strto + Sheet2.Range("rgconchas").Value + "; "
            ElseIf Sheet2.opt_CCC6.Value = True Then
                strcc = strcc + Sheet2.Range("rgconchas").Value + "; "
            ElseIf Sheet2.opt_NC6.Value = True Then
            End If
            
            If Sheet2.opt_ToC7.Value = True Then
                strto = strto + Sheet2.Range("vbcuevas").Value + "; "
            ElseIf Sheet2.opt_CCC7.Value = True Then
                strcc = strcc + Sheet2.Range("vbcuevas").Value + "; "
            ElseIf Sheet2.opt_NC7.Value = True Then
            End If
            
            If Sheet2.opt_ToC8.Value = True Then
                strto = strto + Sheet2.Range("jbdesemprado").Value + "; "
            ElseIf Sheet2.opt_CCC8.Value = True Then
                strcc = strcc + Sheet2.Range("jbdesemprado").Value + "; "
            ElseIf Sheet2.opt_NC8.Value = True Then
            End If
            
            If Sheet2.opt_ToC9.Value = True Then
                strto = strto + Sheet2.Range("rddesquitado").Value + "; "
            ElseIf Sheet2.opt_CCC9.Value = True Then
                strcc = strcc + Sheet2.Range("rddesquitado").Value + "; "
            ElseIf Sheet2.opt_NC9.Value = True Then
            End If
            
            If Sheet2.opt_ToC10.Value = True Then
                strto = strto + Sheet2.Range("mmdevero").Value + "; "
            ElseIf Sheet2.opt_CCC10.Value = True Then
                strcc = strcc + Sheet2.Range("mmdevero").Value + "; "
            ElseIf Sheet2.opt_NC10.Value = True Then
            End If
            
            If Sheet2.opt_ToC11.Value = True Then
                strto = strto + Sheet2.Range("jcfelisarta").Value + "; "
            ElseIf Sheet2.opt_CCC11.Value = True Then
                strcc = strcc + Sheet2.Range("jcfelisarta").Value + "; "
            ElseIf Sheet2.opt_NC11.Value = True Then
            End If
            
            If Sheet2.opt_ToC12.Value = True Then
                strto = strto + Sheet2.Range("rdflores").Value + "; "
            ElseIf Sheet2.opt_CCC12.Value = True Then
                strcc = strcc + Sheet2.Range("rdflores").Value + "; "
            ElseIf Sheet2.opt_NC12.Value = True Then
            End If
            
            If Sheet2.opt_ToC13.Value = True Then
                strto = strto + Sheet2.Range("gpintes").Value + "; "
            ElseIf Sheet2.opt_CCC13.Value = True Then
                strcc = strcc + Sheet2.Range("gpintes").Value + "; "
            ElseIf Sheet2.opt_NC13.Value = True Then
            End If
            
            If Sheet2.opt_ToC14.Value = True Then
                strto = strto + Sheet2.Range("dsisleta").Value + "; "
            ElseIf Sheet2.opt_CCC14.Value = True Then
                strcc = strcc + Sheet2.Range("dsisleta").Value + "; "
            ElseIf Sheet2.opt_NC14.Value = True Then
            End If
            
            If Sheet2.opt_ToC15.Value = True Then
                strto = strto + Sheet2.Range("rmlocaylocay").Value + "; "
            ElseIf Sheet2.opt_CCC15.Value = True Then
                strcc = strcc + Sheet2.Range("rmlocaylocay").Value + "; "
            ElseIf Sheet2.opt_NC15.Value = True Then
            End If
            
            If Sheet2.opt_ToC16.Value = True Then
                strto = strto + Sheet2.Range("mmnadal").Value + "; "
            ElseIf Sheet2.opt_CCC16.Value = True Then
                strcc = strcc + Sheet2.Range("mmnadal").Value + "; "
            ElseIf Sheet2.opt_NC16.Value = True Then
            End If
            
            If Sheet2.opt_ToC17.Value = True Then
                strto = strto + Sheet2.Range("npompad").Value + "; "
            ElseIf Sheet2.opt_CCC17.Value = True Then
                strcc = strcc + Sheet2.Range("npompad").Value + "; "
            ElseIf Sheet2.opt_NC17.Value = True Then
            End If
            
            If Sheet2.opt_ToC18.Value = True Then
                strto = strto + Sheet2.Range("mpoplas").Value + "; "
            ElseIf Sheet2.opt_CCC18.Value = True Then
                strcc = strcc + Sheet2.Range("mpoplas").Value + "; "
            ElseIf Sheet2.opt_NC18.Value = True Then
            End If
            
            If Sheet2.opt_ToC19.Value = True Then
                strto = strto + Sheet2.Range("dbpepito").Value + "; "
            ElseIf Sheet2.opt_CCC19.Value = True Then
                strcc = strcc + Sheet2.Range("dbpepito").Value + "; "
            ElseIf Sheet2.opt_NC19.Value = True Then
            End If
            
            If Sheet2.opt_ToC20.Value = True Then
                strto = strto + Sheet2.Range("izpono").Value + "; "
            ElseIf Sheet2.opt_CCC20.Value = True Then
                strcc = strcc + Sheet2.Range("izpono").Value + "; "
            ElseIf Sheet2.opt_NC20.Value = True Then
            End If
            
            If Sheet2.opt_ToC21.Value = True Then
                strto = strto + Sheet2.Range("clrosales").Value + "; "
            ElseIf Sheet2.opt_CCC21.Value = True Then
                strcc = strcc + Sheet2.Range("clrosales").Value + "; "
            ElseIf Sheet2.opt_NC21.Value = True Then
            End If
            
            If Sheet2.opt_ToC22.Value = True Then
                strto = strto + Sheet2.Range("rasalas").Value + "; "
            ElseIf Sheet2.opt_CCC22.Value = True Then
                strcc = strcc + Sheet2.Range("rasalas").Value + "; "
            ElseIf Sheet2.opt_NC22.Value = True Then
            End If
    
            If Sheet2.opt_ToC24.Value = True Then
                strto = strto + Sheet2.Range("mdsarcuaga").Value + "; "
            ElseIf Sheet2.opt_CCC24.Value = True Then
                strcc = strcc + Sheet2.Range("mdsarcuaga").Value + "; "
            ElseIf Sheet2.opt_NC24.Value = True Then
            End If
            
            If Sheet2.opt_ToC25.Value = True Then
                strto = strto + Sheet2.Range("mlsarmiento").Value + "; "
            ElseIf Sheet2.opt_CCC25.Value = True Then
                strcc = strcc + Sheet2.Range("mlsarmiento").Value + "; "
            ElseIf Sheet2.opt_NC25.Value = True Then
            End If
            
            If Sheet2.opt_ToC26.Value = True Then
                strto = strto + Sheet2.Range("vjson").Value + "; "
            ElseIf Sheet2.opt_CCC26.Value = True Then
                strcc = strcc + Sheet2.Range("vjson").Value + "; "
            ElseIf Sheet2.opt_NC26.Value = True Then
            End If
            
            If Sheet2.opt_ToC27.Value = True Then
                strto = strto + Sheet2.Range("jltacan").Value + "; "
            ElseIf Sheet2.opt_CCC27.Value = True Then
                strcc = strcc + Sheet2.Range("jltacan").Value + "; "
            ElseIf Sheet2.opt_NC27.Value = True Then
            End If
            
            If Sheet2.opt_ToC28.Value = True Then
                strto = strto + Sheet2.Range("fotamarra").Value + "; "
            ElseIf Sheet2.opt_CCC28.Value = True Then
                strcc = strcc + Sheet2.Range("fotamarra").Value + "; "
            ElseIf Sheet2.opt_NC28.Value = True Then
            End If
            
            If Sheet2.opt_ToC29.Value = True Then
                strto = strto + Sheet2.Range("metejano").Value + "; "
            ElseIf Sheet2.opt_CCC29.Value = True Then
                strcc = strcc + Sheet2.Range("metejano").Value + "; "
            ElseIf Sheet2.opt_NC29.Value = True Then
            End If
            
            If Sheet2.opt_ToC291.Value = True Then
                strto = strto + Sheet2.Range("blynot").Value + "; "
            ElseIf Sheet2.opt_CCC291.Value = True Then
                strcc = strcc + Sheet2.Range("blynot").Value + "; "
            ElseIf Sheet2.opt_NC291.Value = True Then
            End If
            
            If Sheet2.opt_ToC30.Value = True Then
                strto = strto + Sheet2.Range("eddivinagracia").Value + "; "
            ElseIf Sheet2.opt_CCC30.Value = True Then
                strcc = strcc + Sheet2.Range("eddivinagracia").Value + "; "
            ElseIf Sheet2.opt_NC30.Value = True Then
            End If
            
            If Sheet2.opt_ToC31.Value = True Then
                strto = strto + Sheet2.Range("rngloria").Value + "; "
            ElseIf Sheet2.opt_CCC31.Value = True Then
                strcc = strcc + Sheet2.Range("rngloria").Value + "; "
            ElseIf Sheet2.opt_NC31.Value = True Then
            End If
            
            If Sheet2.opt_ToC32.Value = True Then
                strto = strto + Sheet2.Range("rrson").Value + "; "
            ElseIf Sheet2.opt_CCC32.Value = True Then
                strcc = strcc + Sheet2.Range("rrson").Value + "; "
            ElseIf Sheet2.opt_NC32.Value = True Then
            End If
            
            If Sheet2.opt_ToC33.Value = True Then
                strto = strto + Sheet2.Range("cminoferio").Value + "; "
            ElseIf Sheet2.opt_CCC33.Value = True Then
                strcc = strcc + Sheet2.Range("cminoferio").Value + "; "
            ElseIf Sheet2.opt_NC33.Value = True Then
            End If
            
            If Sheet2.opt_ToC34.Value = True Then
                strto = strto + Sheet2.Range("jrmaninang").Value + "; "
            ElseIf Sheet2.opt_CCC34.Value = True Then
                strcc = strcc + Sheet2.Range("jrmaninang").Value + "; "
            ElseIf Sheet2.opt_NC34.Value = True Then
            End If
            
            If Sheet2.opt_ToC35.Value = True Then
                strto = strto + Sheet2.Range("mjvendiola").Value + "; "
            ElseIf Sheet2.opt_CCC35.Value = True Then
                strcc = strcc + Sheet2.Range("mjvendiola").Value + "; "
            ElseIf Sheet2.opt_NC35.Value = True Then
            End If

            strbody = ""
            strbody = strbody + "CEBU" + vbCrLf
            For i = 1 To Status.lst_C.ListCount - 1
                If Status.lst_C.List(i) Like "*Off-site Documentation Audit*" Or Status.lst_C.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_C.List(i) + vbCrLf
                End If
            Next i
            
        ElseIf Status.MultiPage1.Pages(7).Caption = "*CEBU" And Status.chk_APFMC.Value = True Then
            If Sheet2.opt_ToCAF1.Value = True Then
                strto = strto + Sheet2.Range("jjdagay").Value + "; "
            ElseIf Sheet2.opt_CCCAF1.Value = True Then
                strcc = strcc + Sheet2.Range("jjdagay").Value + "; "
            ElseIf Sheet2.opt_NCAF1.Value = True Then
            End If
            
            If Sheet2.opt_ToCAF2.Value = True Then
                strto = strto + Sheet2.Range("lgelicanal").Value + "; "
            ElseIf Sheet2.opt_CCCAF2.Value = True Then
                strcc = strcc + Sheet2.Range("lgelicanal").Value + "; "
            ElseIf Sheet2.opt_NCAF2.Value = True Then
            End If
            
            If Sheet2.opt_ToCAF3.Value = True Then
                strto = strto + Sheet2.Range("mycondrillon").Value + "; "
            ElseIf Sheet2.opt_CCCAF3.Value = True Then
                strcc = strcc + Sheet2.Range("mycondrillon").Value + "; "
            ElseIf Sheet2.opt_NCAF3.Value = True Then
            End If
            
            strbody = ""
            strbody = strbody + "CEBU" + vbCrLf
            For i = 1 To Status.lst_C.ListCount - 1
                If Status.lst_C.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_C.List(i) + vbCrLf
                End If
            Next i
        End If
    End If
    
    If Status.MultiPage1.Pages(3).Caption = "*DILIMAN" Then
        If Status.chk_APFMDL.Value = True And Status.MultiPage1.Value = 3 Then
            If Sheet2.opt_ToMMAF1.Value = True Then
                strto = strto + Sheet2.Range("acty").Value + "; "
            ElseIf Sheet2.opt_CCMMAF1.Value = True Then
                strcc = strcc + Sheet2.Range("acty").Value + "; "
            ElseIf Sheet2.opt_NMMAF1.Value = True Then
            End If
                
            If Sheet2.opt_ToMMAF3.Value = True Then
                strto = strto + Sheet2.Range("achaling").Value + "; "
            ElseIf Sheet2.opt_CCMMAF3.Value = True Then
                strcc = strcc + Sheet2.Range("achaling").Value + "; "
            ElseIf Sheet2.opt_NMMAF3.Value = True Then
            End If
                
            If Sheet2.opt_ToMMAF4.Value = True Then
                strto = strto + Sheet2.Range("rdmontemayor").Value + "; "
            ElseIf Sheet2.opt_CCMMAF4.Value = True Then
                strcc = strcc + Sheet2.Range("rdmontemayor").Value + "; "
            ElseIf Sheet2.opt_NMMAF4.Value = True Then
            End If
                
            If Sheet2.opt_ToMMAF8.Value = True Then
                strto = strto + Sheet2.Range("rgreyes").Value + "; "
            ElseIf Sheet2.opt_CCMMAF8.Value = True Then
                strcc = strcc + Sheet2.Range("rgreyes").Value + "; "
            ElseIf Sheet2.opt_NMMAF8.Value = True Then
            End If
                
            If Sheet2.opt_ToMMAF9.Value = True Then
                strto = strto + Sheet2.Range("ggancaya").Value + "; "
            ElseIf Sheet2.opt_CCMMAF9.Value = True Then
                strcc = strcc + Sheet2.Range("ggancaya").Value + "; "
            ElseIf Sheet2.opt_NMMAF9.Value = True Then
            End If
                
            strbody = ""
            strbody = strbody + "DILIMAN" + vbCrLf
            For i = 1 To Status.lst_DL.ListCount - 1
                If Status.lst_DL.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_DL.List(i) + vbCrLf
                End If
            Next i
        End If
    End If
    
    If Status.MultiPage1.Pages(4).Caption = "*GARNET" Then
        If Status.chk_APFMG.Value = True And Status.MultiPage1.Value = 4 Then
            
            If Sheet2.opt_ToMMAF1.Value = True Then
                strto = strto + Sheet2.Range("acty").Value + "; "
            ElseIf Sheet2.opt_CCMMAF1.Value = True Then
                strcc = strcc + Sheet2.Range("acty").Value + "; "
            ElseIf Sheet2.opt_NMMAF1.Value = True Then
            End If
                
            If Sheet2.opt_ToMMAF6.Value = True Then
                strto = strto + Sheet2.Range("vfserrano").Value + "; "
            ElseIf Sheet2.opt_CCMMAF6.Value = True Then
                strcc = strcc + Sheet2.Range("vfserrano").Value + "; "
            ElseIf Sheet2.opt_NMMAF6.Value = True Then
            End If
            
            If Sheet2.opt_ToMMAF7.Value = True Then
                strto = strto + Sheet2.Range("apoh").Value + "; "
            ElseIf Sheet2.opt_CCMMAF7.Value = True Then
                strcc = strcc + Sheet2.Range("apoh").Value + "; "
            ElseIf Sheet2.opt_NMMAF7.Value = True Then
            End If
                
            If Sheet2.opt_ToMMAF8.Value = True Then
                strto = strto + Sheet2.Range("rgreyes").Value + "; "
            ElseIf Sheet2.opt_CCMMAF8.Value = True Then
                strcc = strcc + Sheet2.Range("rgreyes").Value + "; "
            ElseIf Sheet2.opt_NMMAF8.Value = True Then
            End If
                
            If Sheet2.opt_ToMMAF11.Value = True Then
                strto = strto + Sheet2.Range("jymendoza").Value + "; "
            ElseIf Sheet2.opt_CCMMAF11.Value = True Then
                strcc = strcc + Sheet2.Range("jymendoza").Value + "; "
            ElseIf Sheet2.opt_NMMAF11.Value = True Then
            End If
                
            strbody = ""
            strbody = strbody + "GARNET" + vbCrLf
            For i = 1 To Status.lst_G.ListCount - 1
                If Status.lst_G.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_G.List(i) + vbCrLf
                End If
            Next i
        End If
    End If
        
    If Status.MultiPage1.Pages(5).Caption = "*SAMPALOC" Then
        If Status.chk_APFMS.Value = True And Status.MultiPage1.Value = 5 Then
            If Sheet2.opt_ToMMAF1.Value = True Then
                strto = strto + Sheet2.Range("acty").Value + "; "
            ElseIf Sheet2.opt_CCMMAF1.Value = True Then
                strcc = strcc + Sheet2.Range("acty").Value + "; "
            ElseIf Sheet2.opt_NMMAF1.Value = True Then
            End If
                
            If Sheet2.opt_ToMMAF2.Value = True Then
                strto = strto + Sheet2.Range("rdrubin").Value + "; "
            ElseIf Sheet2.opt_CCMMAF2.Value = True Then
                strcc = strcc + Sheet2.Range("rdrubin").Value + "; "
            ElseIf Sheet2.opt_NMMAF2.Value = True Then
            End If
                
            strbody = ""
            strbody = strbody + "SAMPALOC" + vbCrLf
            For i = 1 To Status.lst_S.ListCount - 1
                If Status.lst_S.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_S.List(i) + vbCrLf
                End If
            Next i
        End If
    End If
            
    If Status.MultiPage1.Pages(7).Caption = "*GREENHILLS" Then
        If Status.chk_APFMGH.Value = True And Status.MultiPage1.Value = 7 Then
            If Sheet2.opt_ToMMAF1.Value = True Then
                strto = strto + Sheet2.Range("acty").Value + "; "
            ElseIf Sheet2.opt_CCMMAF1.Value = True Then
                strcc = strcc + Sheet2.Range("acty").Value + "; "
            ElseIf Sheet2.opt_NMMAF1.Value = True Then
            End If
                
            If Sheet2.opt_ToMMAF5.Value = True Then
                strto = strto + Sheet2.Range("evalayon").Value + "; "
            ElseIf Sheet2.opt_CCMMAF5.Value = True Then
                strcc = strcc + Sheet2.Range("evalayon").Value + "; "
            ElseIf Sheet2.opt_NMMAF5.Value = True Then
            End If
                
            If Sheet2.opt_ToMMAF8.Value = True Then
                strto = strto + Sheet2.Range("rgreyes").Value + "; "
            ElseIf Sheet2.opt_CCMMAF8.Value = True Then
                strcc = strcc + Sheet2.Range("rgreyes").Value + "; "
            ElseIf Sheet2.opt_NMMAF8.Value = True Then
            End If
                
            If Sheet2.opt_ToMMAF10.Value = True Then
                strto = strto + Sheet2.Range("dpstotomas").Value + "; "
            ElseIf Sheet2.opt_CCMMAF10.Value = True Then
                strcc = strcc + Sheet2.Range("dpstotomas").Value + "; "
            ElseIf Sheet2.opt_NMMAF10.Value = True Then
            End If
                
            strbody = ""
            strbody = strbody + "GREENHILLS" + vbCrLf
            For i = 1 To Status.lst_GH.ListCount - 1
                If Status.lst_GH.List(i) Like "*Facility Audit*" Then
                    strbody = strbody + "   " + Status.lst_GH.List(i) + vbCrLf
                End If
            Next i
        End If
    End If
    
    Dim signature As String
    
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
        .Subject = "BC Gov IA Deadline"
        .Body = strbody
        .Attachments.Add (Sheet3.Range("A11").Value)
        .Attachments.Add Sheet3.Range("A5").Value, olByValue, 0
        .HTMLBody = "<style>body{color:red;font-weight:bold}</style><h2>DUE DATE NEAR EXPIRATION:</h2>" & .HTMLBody & Sheet3.Range("A2").Value & "<br><br>" & "<img src='cid:" & Sheet3.Range("A8").Value & "'>" & signature
        .Display
    End With
    On Error GoTo 0
    
    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub

