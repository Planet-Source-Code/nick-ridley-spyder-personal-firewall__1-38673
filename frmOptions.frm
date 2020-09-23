VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Spyder Personal Firewall - Options"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5235
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   9234
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Block"
      TabPicture(0)   =   "frmOptions.frx":038A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "SSTab2"
      Tab(0).Control(1)=   "Command2"
      Tab(0).Control(2)=   "Command3"
      Tab(0).Control(3)=   "txtAdd"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Log"
      TabPicture(1)   =   "frmOptions.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstLog"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Options"
      TabPicture(2)   =   "frmOptions.frx":03C2
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Picture1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Check1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Check2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Check3"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.CheckBox Check3 
         Caption         =   "warn when Spyder Personal Firewall is exiting"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   2340
         Width           =   6255
      End
      Begin VB.Frame Frame1 
         Caption         =   "security level"
         Height          =   2355
         Left            =   180
         TabIndex        =   13
         Top             =   2700
         Width           =   6195
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            Height          =   780
            Left            =   5340
            Picture         =   "frmOptions.frx":03DE
            ScaleHeight     =   720
            ScaleWidth      =   720
            TabIndex        =   17
            Top             =   1500
            Width           =   780
         End
         Begin MSComctlLib.Slider Slider1 
            Height          =   1995
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   3519
            _Version        =   393216
            Orientation     =   1
            LargeChange     =   1
            Min             =   1
            Max             =   3
            SelStart        =   1
            TickStyle       =   2
            Value           =   1
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   780
            Picture         =   "frmOptions.frx":0BC5
            Top             =   240
            Width           =   240
         End
         Begin VB.Label lblDes 
            Caption         =   "[description]"
            Height          =   1755
            Left            =   1140
            TabIndex        =   16
            Top             =   480
            Width           =   4155
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "High Security"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1080
            TabIndex        =   15
            Top             =   240
            Width           =   1125
         End
      End
      Begin VB.CheckBox Check2 
         Caption         =   "warn when Spyder Personal Firewall is put in the system tray"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   1980
         Width           =   6255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "auto accept network traffic (subnets 192.168.x.x , 10.x.x.x)"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   1620
         Width           =   6255
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   960
         Left            =   180
         Picture         =   "frmOptions.frx":0F4F
         ScaleHeight     =   900
         ScaleWidth      =   6180
         TabIndex        =   10
         Top             =   540
         Width           =   6240
      End
      Begin VB.TextBox txtAdd 
         Height          =   315
         Left            =   -73050
         TabIndex        =   9
         Top             =   4800
         Width           =   2835
      End
      Begin VB.CommandButton Command3 
         Caption         =   "remove"
         Height          =   315
         Left            =   -74850
         TabIndex        =   8
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "add"
         Height          =   315
         Left            =   -70170
         TabIndex        =   7
         Top             =   4800
         Width           =   1575
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   4215
         Left            =   -74850
         TabIndex        =   3
         Top             =   480
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   7435
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "IP"
         TabPicture(0)   =   "frmOptions.frx":3260
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lstIP"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Remote Port"
         TabPicture(1)   =   "frmOptions.frx":327C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lstRemP"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Local Port"
         TabPicture(2)   =   "frmOptions.frx":3298
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lstLocP"
         Tab(2).ControlCount=   1
         Begin VB.ListBox lstLocP 
            Height          =   3570
            Left            =   -74880
            TabIndex        =   6
            Top             =   480
            Width           =   6075
         End
         Begin VB.ListBox lstRemP 
            Height          =   3570
            Left            =   -74880
            TabIndex        =   5
            Top             =   480
            Width           =   6075
         End
         Begin VB.ListBox lstIP 
            Height          =   3570
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   6075
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "clear log"
         Height          =   315
         Left            =   -70440
         TabIndex        =   2
         Top             =   4800
         Width           =   1995
      End
      Begin MSComctlLib.ListView lstLog 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   1
         Top             =   420
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   7646
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Timestamp"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "IP"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "RemotePort"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "LocalPort"
            Object.Width           =   1270
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Info"
            Object.Width           =   3105
         EndProperty
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()

modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\options", "subnet", Check1.Value

End Sub

Private Sub Check2_Click()

If Check2.Value = "1" Then

    modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\options", "warnsys", "0"

Else

    modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\options", "warnsys", "1"

End If

End Sub

Private Sub Check3_Click()

If Check3.Value = "1" Then

    modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\options", "warnexit", "0"

Else

    modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\options", "warnexit", "1"

End If

End Sub

Private Sub Command1_Click()

Dim r As String

r = MsgBox("firewall logs may be necessary for reporting and prosecution of potential hackers!" & vbCrLf & vbCrLf & "clear log?", vbQuestion & vbYesNo, "Warning!")

If r = vbYes Then

    Dim ff As Long
    ff = FreeFile

    Open App.Path & "\SPFlog.log" For Output As #ff

    Close #ff
    
    lstLog.ListItems.Clear
    
End If

End Sub

Private Sub Command2_Click()

Dim key As String
Dim lstBox As ListBox

Select Case SSTab2.Tab

    Case 0
    key = "\block\ip"
    Set lstBox = lstIP
    
    Case 1
    key = "\block\rp"
    Set lstBox = lstRemP
    If Not IsNumeric(txtAdd) Then Exit Sub
    
    Case 2
    key = "\block\lp"
    Set lstBox = lstLocP
    If Not IsNumeric(txtAdd) Then Exit Sub
    
End Select

modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & key, txtAdd, "1"
lstBox.AddItem txtAdd
txtAdd = ""

End Sub

Private Sub Command3_Click()

Dim key As String
Dim lstBox As ListBox

Select Case SSTab2.Tab

    Case 0
    key = "\block\ip"
    Set lstBox = lstIP
    
    Case 1
    key = "\block\rp"
    Set lstBox = lstRemP
    
    Case 2
    key = "\block\lp"
    Set lstBox = lstLocP
    
End Select

modReg.DelSetting HKEY_LOCAL_MACHINE, SREG & key, lstBox.List(lstBox.ListIndex)
lstBox.RemoveItem lstBox.ListIndex

End Sub

Private Sub Form_Load()

Dim x As String

x = modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\options", "subnet")

If x = "1" Then Check1.Value = 1

If modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\options", "warnsys") <> "1" Then

    Check2.Value = 1

End If

If modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\options", "warnexit") <> "1" Then

    Check3.Value = 1

End If

If modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\options", "seclvl") = "" Then modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\options", "seclvl", "2"
Slider1 = modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\options", "seclvl")

Call Slider1_Change

End Sub

Private Sub Slider1_Change()

Select Case Slider1

    Case 1
    lblTitle = "High Security"
    lblDes = "all attempted connections were any 3 criteria does not have a rule, comes up with an alert. only needed when a hacker is believed to be present in the system"
    
    Case 2
    lblTitle = "Medium Security"
    lblDes = "only alerts when all 3 criteria have no rule. recommended for most systems"
    
    Case 3
    lblTitle = "Low Security"
    lblDes = "all connections are let through. rules are ignored exept for on an IP based level. ports are only affected when a trajan is susspected (not added yet)"
    
End Select

modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\options", "seclvl", Slider1

End Sub

