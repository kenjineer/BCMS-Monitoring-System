VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-'
'    Title: BCMS Project Monitoring (All Sites)     '
'    Author: Engr. Kenneth Caro Karamihan           '
'    Company: PLDT Inc.                             '
'    Division: BC Governance and Reporting          '
'    Date: May 24, 2018                             '
'    Code version: 2.0                              '
'-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-'

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Selection.Count = 1 Then
        If Not Intersect(Target, Range("A2:AH" & Cells(Rows.Count, "B").End(xlUp).Row + 1)) Is Nothing Then
            Call Status
        End If
    End If
End Sub
