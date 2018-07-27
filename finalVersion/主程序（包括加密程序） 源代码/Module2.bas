Attribute VB_Name = "Module2"
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim pid As Long
    GetWindowThreadProcessId hwnd, pid
    If pid = lParam Then
        ShowWindow hwnd, vbHide
    End If
    EnumWindowsProc = 1
End Function
 
Public Function RunApp(ByVal sApp As String) As Boolean
    Dim pid As Long
    pid = Shell(sApp, vbHide)
    EnumWindows AddressOf EnumWindowsProc, pid
End Function
