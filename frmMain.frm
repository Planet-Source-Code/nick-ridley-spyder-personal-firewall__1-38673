VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spyder Personal Firewall"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1111
      ButtonWidth     =   1111
      ButtonHeight    =   1058
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgCold"
      HotImageList    =   "imgHot"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "stop"
            Object.ToolTipText     =   "stop monitoring"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "refresh"
            Object.ToolTipText     =   "refresh the connections list"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "active"
            Object.ToolTipText     =   "show active connections only"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "options"
            Object.ToolTipText     =   "show the options form"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "block"
            Object.ToolTipText     =   "view and edit the blocking list"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "log"
            Object.ToolTipText     =   "view the log"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgHot 
      Left            =   180
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1296
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D4E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgCold 
      Left            =   180
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2806
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D62
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3712
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C6E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   250
      Left            =   180
      Top             =   720
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3615
      Left            =   37
      TabIndex        =   0
      Top             =   720
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   6376
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Remote IP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Remote Port"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Local Port"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "State"
         Object.Width           =   3387
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lC As Integer
Public Blk As String

Private a_RemA(1000) As String
Private a_LocP(1000) As String
Private a_RemP(1000) As String

Private a_Count As Long

Public Function RefreshTable(Optional force As Boolean = False)

On Error Resume Next

Dim tcpt As MIB_TCPTABLE, l As Long
Dim x As Integer, i As Integer
Dim RemA As String, LocP As String, RemP As String
Dim state As Integer

l = Len(MIB_TCPTABLE)
GetTcpTable tcpt, l, 0
x = tcpt.dwNumEntries

If x < lC Or x > lC Or force Then

    lC = x
    
    ListView1.ListItems.Clear
    
    For i = 0 To x - 1
    
        RemA = GetAscIP(tcpt.table(i).dwRemoteAddr)
        RemP = ntohs(tcpt.table(i).dwRemotePort)
        LocP = ntohs(tcpt.table(i).dwLocalPort)
        
        state = tcpt.table(i).dwState
        
        If state = 2 Then
        
            If Toolbar1.Buttons(4).Caption = "active" Then
        
                ListView1.ListItems.Add , "x" & i, "n/a"
                ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "n/a"
                ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = LocP
                ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = modNetstat.IP_States(state)
        
            End If
            
        Else
        
            ListView1.ListItems.Add , "x" & i, RemA
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = RemP
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = LocP
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = modNetstat.IP_States(state)
        
        End If
    
        CheckBlock RemA, RemP, LocP, state, i
        
    Next i

End If

End Function

Private Sub Form_Load()

Me.Hide
DoEvents

Blk = 1

modNetstat.InitStates
LoadSettings
Hook Me.hwnd
AddIconToTray Me.hwnd, Me.Icon, Me.Icon.Handle, "Spyder Personal Firewall"

RefreshTable

Me.Hide

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
If modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\options", "warnsys") <> "1" Then frmWarnSys.Show
Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Form_QueryUnload(0, 0)
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 And ListView1.ListItems.Count > 0 Then
    frmMain.PopupMenu frmMenu.mnuConn
End If
End Sub

Private Sub Timer1_Timer()
Text1 = ListView1.ColumnHeaders(1).Width & " - " & ListView1.ColumnHeaders(2).Width & " - " & ListView1.ColumnHeaders(3).Width & " - " & ListView1.ColumnHeaders(4).Width
End Sub

Private Sub tmrRefresh_Timer()

RefreshTable

End Sub

Public Function CheckBlock(RemA As String, RemP As String, LocP As String, state As Integer, i As Integer)

On Error Resume Next
Dim tcpt As MIB_TCPTABLE
Dim l As Long
Dim x As Long
Dim b As Boolean, a As Boolean
Dim lvl As Long

If frmOptions.Check1.Value = 1 Then

    If Left(RemA, 7) = "192.168" Then Exit Function
    If Left(RemA, 3) = "10." Then Exit Function

End If

For x = 0 To a_Count - 1

    If a_RemA(x) = RemA Then Exit Function
    If a_RemP(x) = RemP Then Exit Function
    If a_LocP(x) = LocP Then Exit Function

Next x

'modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\block\ip", RemA) = "1" Or modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\block\rp", Str(RemP)) = "1" Or modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\block\lp", Str(LocP)) = "1"
'modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\block\ip", RemA) <> "0" And modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\block\rp", RemP) <> "0" And modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\block\lp", LocP) <> "0"

lvl = frmOptions.Slider1

If lvl = 1 Then

    If modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\block\ip", RemA) = "1" Or modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\block\rp", Str(RemP)) = "1" Or modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\block\lp", Str(LocP)) = "1" Then b = True
    If modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\block\ip", RemA) <> "0" Or modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\block\rp", RemP) <> "0" Or modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\block\lp", LocP) <> "0" Then a = True

ElseIf lvl = 2 Then

    If modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\block\ip", RemA) = "1" Or modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\block\rp", Str(RemP)) = "1" Or modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\block\lp", Str(LocP)) = "1" Then b = True
    If modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\block\ip", RemA) <> "0" And modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\block\rp", RemP) <> "0" And modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\block\lp", LocP) <> "0" Then a = True

ElseIf lvl = 3 Then

    If IsTrojan(LocP) And modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\block\ip", RemA) <> "0" And modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\block\rp", RemP) <> "0" And modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\block\lp", LocP) <> "0" Then a = True
    If a = False And IsTrojan(LocP) Then a = True

End If

If state <> 2 Then

    If b Then
    
        l = Len(MIB_TCPTABLE)
        GetTcpTable tcpt, l, 0
    
        tcpt.table(i).dwState = 12
        SetTcpEntry tcpt.table(i)
        DoEvents
        
        Log RemA, RemP, LocP, "disconnected blocked connection"
    
    ElseIf a Then
            
        Dim z As New frmAlert
        
        z.t = i
        z.RemA = RemA
        z.RemP = RemP
        z.LocP = LocP
        
        z.t = a_Count
        a_RemA(a_Count) = RemA
        a_RemP(a_Count) = RemA
        a_LocP(a_Count) = RemA
        
        a_Count = a_Count + 1
        
        z.b_Type = 1
        
        Load z
        DoEvents
        
    End If

End If

End Function

Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index

    Case 1
    
        If Button.Caption = "stop" Then
        
            Button.Caption = "stopped"
            Button.ToolTipText = "resume monitoring"
            tmrRefresh.Enabled = False
            
        Else
        
            Button.Caption = "stop"
            Button.ToolTipText = "stop monitoring"
            tmrRefresh.Enabled = True
            
        End If
    
    Case 2
    
        RefreshTable
    
    Case 4
    
        If Button.Caption = "active" Then
        
            Button.Caption = "all"
            Button.ToolTipText = "show all connections"
            RefreshTable True
            
        Else
        
            Button.Caption = "active"
            Button.ToolTipText = "show active connections only"
            RefreshTable True
            
        End If
        
        modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\options", "active", Toolbar1.Buttons(4).Value
    
    Case 5
        
        frmOptions.Show
        frmOptions.SSTab1.Tab = 2
    
    Case 7
        
        frmOptions.Show
        frmOptions.SSTab1.Tab = 0
        
    Case 8
        
        frmOptions.Show
        frmOptions.SSTab1.Tab = 1
    
    
End Select

End Sub

Public Sub SysTray()
       
If Me.Visible = True Then

    Me.Hide

Else

    Me.Show
    SetFGWindow Me.hwnd, True
    
End If

End Sub

Public Sub SysTray2()

PopupMenu frmMenu.mnuMain, vbPopupMenuRightButton

End Sub

Public Sub ShiftAlertList(t As Long)

Dim i As Long

For i = t + 1 To a_Count - 1

    a_RemA(i - 1) = a_RemA(i)
    a_RemP(i - 1) = a_RemP(i)
    a_LocP(i - 1) = a_LocP(i)

Next i

a_Count = a_Count - 1

End Sub

Public Function IsTrojan(port) As Boolean

Dim ff As Long
Dim p As Long, l As String
Dim po As String, de As String

ff = FreeFile

Open App.Path & "\trojan.dat" For Input As #ff

Do Until EOF(ff)

    Line Input #ff, l
    
    p = InStr(1, l, Chr(1))
    po = Left(l, p - 1)

    If po = port Then IsTrojan = True: Close #ff: Exit Function

Loop

Close #ff

End Function

Public Function GetTrojan(port) As String

Dim ff As Long
Dim p As Long, l As String
Dim po As String, de As String

ff = FreeFile

Open App.Path & "\trojan.dat" For Input As #ff

Do Until EOF(ff)

    Line Input #ff, l
    
    p = InStr(1, l, Chr(1))
    po = Left(l, p - 1)
    l = Right(l, Len(l) - p)
    p = InStr(1, l, Chr(1))
    de = Left(l, p - 1)

    If po = port Then GetTrojan = de: Close #ff: Exit Function

Loop

Close #ff

End Function
