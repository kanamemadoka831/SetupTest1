VERSION 5.00
Begin VB.Form fRijndael 
   Caption         =   "Form1"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   10065
   Visible         =   0   'False
End
Attribute VB_Name = "fRijndael"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim objSWbemLocator As New SWbemLocator
Dim objSWbemServices As SWbemServices
Dim objSWbemObjectSet As SWbemObjectSet
Dim objSWbemObject As SWbemObject
Private Sub Form_Load()
'InitCommonControlsVBd
Dim check As Boolean
check = Load()
If check = True Then
Else
End
End If
Dim catia As Object
Dim w As Object
Set w = GetObject("winmgmts:")
Dim p As Object
Dim i As Object
Set p = w.ExecQuery("select * from win32_process where name='CNEXT.exe' ")
If p.Count = 0 Then
  '这里直接接createObject
  Set catia = CreateObject("CATIA.Application")
  catia.Visible = True
Else
'  For Each i In p
'    MsgBox "进程 " & i.Name & " 的 PID 是 " & i.ProcessId
'  Next
    If p.Count = 1 Then
        Set catia = GetObject(, "CATIA.Application")
    Else
        MsgBox "打开了多个CATIA软件，请关闭"
    End If
    
End If
Dim oPartDoc As Document
Set oPartDoc = catia.Documents.Add("Part")

End Sub
