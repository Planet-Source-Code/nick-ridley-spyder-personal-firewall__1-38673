VERSION 5.00
Begin VB.Form frmWarnExit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Warning"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWarnExit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "do not show this again"
      Height          =   195
      Left            =   75
      TabIndex        =   3
      Top             =   1560
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "no"
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
      Left            =   2265
      TabIndex        =   2
      Top             =   1020
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "yes"
      Height          =   375
      Left            =   885
      TabIndex        =   1
      Top             =   1020
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Do you want to continue?"
      Height          =   195
      Left            =   105
      TabIndex        =   4
      Top             =   660
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "If you close Spyder Personal Firewall you will no longer be protected while on the internet."
      Height          =   675
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   4335
   End
End
Attribute VB_Name = "frmWarnExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Check1.Value = "1" Then modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\options", "warnexit", "1"
Unhook
DoEvents
RemoveIconFromTray
DoEvents
frmSplash.tmrEnd.Enabled = True
End Sub

Private Sub Command2_Click()
If Check1.Value = "1" Then modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\options", "warnexit", "1"
Me.Hide
End Sub

Private Sub Form_Load()
If modReg.bGetRegValue(HKEY_LOCAL_MACHINE, SREG & "\options", "warnexit") = "1" Then Check1.Value = 1: Call Command1_Click
End Sub
