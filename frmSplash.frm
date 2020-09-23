VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrEnd 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   900
      Top             =   1140
   End
   Begin VB.Image imgSplash 
      Height          =   675
      Left            =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

imgSplash = LoadPicture(App.Path & "\gfx\splash.gif")

Me.Width = imgSplash.Width
Me.Height = imgSplash.Height

Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Top = Screen.Height / 2 - Me.Height / 2

Me.Show

frmMain.Show
Me.Hide

End Sub

Private Sub tmrEnd_Timer()
End
End Sub
