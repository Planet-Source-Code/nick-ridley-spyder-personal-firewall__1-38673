Attribute VB_Name = "modGeneral"
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_BOTTOM = 1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1

Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40


Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const SW_HIDE = 0
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10

Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Boolean
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long

Public Function WindowPos(frm As Object, setting As Integer)

Dim i As Integer
Select Case setting
Case 1
i = HWND_TOPMOST
Case 2
i = HWND_TOP
Case 3
i = HWND_NOTOPMOST
Case 4
i = HWND_BOTTOM
End Select

SetWindowPos frm.hwnd, i, frm.Left / 15, _
frm.Top / 15, frm.Width / 15, _
frm.Height / 15, SWP_SHOWWINDOW Or SWP_NOACTIVATE

End Function

Public Sub SetFGWindow(ByVal hwnd As Long, Show As Boolean)
If Show Then
If IsIconic(hwnd) Then
ShowWindow hwnd, SW_RESTORE
Else
BringWindowToTop hwnd
End If
Else
ShowWindow hwnd, SW_MINIMIZE
End If
End Sub

Public Function LoadSettings()

Dim ff As Long
Dim data(4) As String
ff = FreeFile

Open App.Path & "\SPFlog.log" For Input As #ff

Do Until EOF(ff)

    Input #ff, data(0), data(1), data(2), data(3), data(4)

    frmOptions.lstLog.ListItems.Add , , data(0)
    frmOptions.lstLog.ListItems(frmOptions.lstLog.ListItems.Count).SubItems(1) = data(1)
    frmOptions.lstLog.ListItems(frmOptions.lstLog.ListItems.Count).SubItems(2) = data(2)
    frmOptions.lstLog.ListItems(frmOptions.lstLog.ListItems.Count).SubItems(3) = data(3)
    frmOptions.lstLog.ListItems(frmOptions.lstLog.ListItems.Count).SubItems(4) = data(4)

Loop

Close #ff

If modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\options", "active") = 1 Then
    
    frmMain.Toolbar1.Buttons(4).Caption = "all"
    frmMain.Toolbar1.Buttons(4).ToolTipText = "show all connections"
    frmMain.Toolbar1.Buttons(4).Value = tbrPressed
    frmMain.RefreshTable True

End If

modReg.bIndexReg HKEY_LOCAL_MACHINE, SREG & "\block\ip", frmOptions.lstIP, "1"
modReg.bIndexReg HKEY_LOCAL_MACHINE, SREG & "\block\lp", frmOptions.lstLocP, "1"
modReg.bIndexReg HKEY_LOCAL_MACHINE, SREG & "\block\rp", frmOptions.lstRemP, "1"

End Function

Public Function Log(RemA As String, RemP As String, LocP As String, Txt As String)

Dim ff As Long
ff = FreeFile

Open App.Path & "\SPFlog.log" For Append As #ff

Write #ff, Time & "-" & Date, RemA, RemP, LocP, Txt

frmOptions.lstLog.ListItems.Add , , Time & "-" & Date
frmOptions.lstLog.ListItems(frmOptions.lstLog.ListItems.Count).SubItems(1) = RemA
frmOptions.lstLog.ListItems(frmOptions.lstLog.ListItems.Count).SubItems(2) = RemP
frmOptions.lstLog.ListItems(frmOptions.lstLog.ListItems.Count).SubItems(3) = LocP
frmOptions.lstLog.ListItems(frmOptions.lstLog.ListItems.Count).SubItems(4) = Txt

Close #ff

End Function
