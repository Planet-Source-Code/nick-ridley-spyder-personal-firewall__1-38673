VERSION 5.00
Begin VB.Form frmAlert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alert"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAlert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "block all"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   3180
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "block loc. port"
      Height          =   375
      Left            =   2580
      TabIndex        =   5
      Top             =   3180
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "block rem. port"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   3180
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "allow"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "block ip"
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   3180
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   127
      Picture         =   "frmAlert.frx":038A
      ScaleHeight     =   1380
      ScaleWidth      =   4800
      TabIndex        =   0
      Top             =   60
      Width           =   4860
   End
   Begin VB.Label Label1 
      Caption         =   "[description]"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4875
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public t As Long
Public RemA As String
Public RemP As String
Public LocP As String

Public b_Type As Integer

Private Sub Command1_Click()

modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\block\ip", RemA, "1"
frmMain.RefreshTable True
Log RemA, RemP, LocP, "blocked ip"
frmMain.ShiftAlertList t
Unload Me

End Sub

Private Sub Command2_Click()

modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\block\ip", RemA, "0"
modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\block\rp", RemP, "0"
modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\block\lp", LocP, "0"
Log RemA, RemP, LocP, "allowed connection"
frmMain.ShiftAlertList t
Unload Me

End Sub

Private Sub Command3_Click()

modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\block\rp", RemP, "1"
frmMain.RefreshTable True
Log RemA, RemP, LocP, "blocked remote port"
frmMain.ShiftAlertList t
Unload Me

End Sub

Private Sub Command4_Click()

modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\block\lp", LocP, "1"
frmMain.RefreshTable True
Log RemA, RemP, LocP, "blocked local port"
frmMain.ShiftAlertList t
Unload Me

End Sub

Private Sub Command5_Click()

modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\block\lp", LocP, "1"
modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\block\rp", RemP, "1"
modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\block\ip", RemA, "1"
frmMain.RefreshTable True
Log RemA, RemP, LocP, "blocked all"
frmMain.ShiftAlertList t
Unload Me

End Sub

Private Sub Form_Load()

Me.Show
WindowPos Me, 1
SetFGWindow Me.hwnd, True
DoEvents

Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Top = Screen.Height / 2 - Me.Height / 2

Select Case b_Type

    Case 1
    tro = frmMain.GetTrojan(LocP)
    If tro = "" Then tro = frmMain.GetTrojan(RemP)
    If tro = "" Then tro = "[none]"
    tro = tro & vbCrLf & vbCrLf
    
    Label1 = RemA & " has tried to access port " & LocP & " form remote port " & RemP & _
    "." & vbCrLf & vbCrLf & "Suspected trojans on this port:" & vbCrLf & tro & "would you like to block this computer?"

End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.ShiftAlertList t
End Sub
