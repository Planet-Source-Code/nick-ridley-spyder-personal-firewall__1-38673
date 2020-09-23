VERSION 5.00
Begin VB.Form frmWarnSys 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Warning"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWarnSys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "ok"
      Height          =   375
      Left            =   1380
      TabIndex        =   3
      Top             =   2580
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "do not show this again"
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   2280
      Width           =   3675
   End
   Begin VB.Label Label2 
      Caption         =   $"frmWarnSys.frx":038A
      Height          =   855
      Left            =   60
      TabIndex        =   1
      Top             =   1380
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Spyder Person Firewall has not been closed. Instead it has been moved to the system tray."
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   997
      Picture         =   "frmWarnSys.frx":043B
      Top             =   660
      Width           =   1860
   End
End
Attribute VB_Name = "frmWarnSys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Check1.Value = "1" Then modReg.bSetRegValue HKEY_LOCAL_MACHINE, SREG & "\options", "warnsys", "1"
Me.Hide
End Sub
