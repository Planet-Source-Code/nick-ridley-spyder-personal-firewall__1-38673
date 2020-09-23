VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Form1"
   ClientHeight    =   345
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   345
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuMain 
      Caption         =   "mnuMain"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuConn 
      Caption         =   "mnuConn"
      Begin VB.Menu mnuDis 
         Caption         =   "Disconnect (dont block)"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBlockAll 
         Caption         =   "Block All"
      End
      Begin VB.Menu mnuBlockIP 
         Caption         =   "Block IP"
      End
      Begin VB.Menu mnuBlockRP 
         Caption         =   "Block Remote Port"
      End
      Begin VB.Menu mnuBlockLP 
         Caption         =   "Block Local Port"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuBlockAll_Click()

BlockList True, True, True

End Sub

Private Sub mnuBlockIP_Click()

BlockList True, False, False

End Sub

Private Sub mnuBlockLP_Click()

BlockList False, False, True

End Sub

Private Sub mnuBlockRP_Click()

BlockList False, True, False

End Sub

Private Sub mnuDis_Click()

Dim tcpt As MIB_TCPTABLE
Dim l As Long
Dim i As Long
Dim RemA As String, RemP As String, LocP As String

i = Right(frmMain.ListView1.SelectedItem.key, Len(frmMain.ListView1.SelectedItem.key) - 1) + 1

RemA = frmMain.ListView1.ListItems(i)
RemP = frmMain.ListView1.ListItems(i).SubItems(1)
LocP = frmMain.ListView1.ListItems(i).SubItems(2)

l = Len(MIB_TCPTABLE)
GetTcpTable tcpt, l, 0

tcpt.table(i - 1).dwState = 12
SetTcpEntry tcpt.table(i - 1)
DoEvents

Log RemA, RemP, LocP, "disconnected connection (no block)"

End Sub

Private Sub mnuExit_Click()
frmWarnExit.Show
End Sub

Private Sub mnuOpen_Click()

If frmMain.WindowState = 0 Then
    
    frmMain.Show
    frmMain.WindowState = 0

Else

    frmMain.Show
    SetFGWindow frmMain.hwnd, True
    
End If

End Sub

Public Function BlockList(ra As Boolean, rp As Boolean, lp As Boolean)

Dim i As Long
Dim RemA As String, RemP As String, LocP As String

i = Right(frmMain.ListView1.SelectedItem.key, Len(frmMain.ListView1.SelectedItem.key) - 1) + 1

RemA = frmMain.ListView1.ListItems(i)
RemP = frmMain.ListView1.ListItems(i).SubItems(1)
LocP = frmMain.ListView1.ListItems(i).SubItems(2)

If ra Then modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\block\ip", RemA, "1"
If rp Then modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\block\rp", RemP, "1"
If lp Then modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\block\lp", LocP, "1"

frmMain.RefreshTable True
Log RemA, RemP, LocP, "blocked all"

End Function
