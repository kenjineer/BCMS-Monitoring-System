VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+'
'    Title: BC Gov IA Calendar Automatic Email and Deadline Notifier     '
'    Author: Engr. Kenneth Caro Karamihan                                '
'    Company: PLDT Inc.                                                  '
'    Division: BC Governance and Reporting                               '
'    Date: May 7, 2018                                                   '
'    Code version: 1.0                                                   '
'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+'



Private Sub Workbook_Open()
    Call Notif
End Sub

