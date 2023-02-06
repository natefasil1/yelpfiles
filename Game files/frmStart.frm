VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   ClientHeight    =   6975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   12780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   4635
      Picture         =   "frmStart.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   4830
      TabIndex        =   0
      Top             =   1710
      Width           =   4890
   End
   Begin VB.Image Image1 
      Height          =   5040
      Left            =   1710
      Picture         =   "frmStart.frx":8457
      Stretch         =   -1  'True
      Top             =   150
      Width           =   9075
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   frmStart.Width = Screen.Width
   frmStart.Height = Screen.Height
   Image1.Left = 0
   Image1.Top = 0
   Image1.Width = Screen.Width
   Image1.Height = Screen.Height
   Picture1.Left = (Screen.Width / 2) - (Picture1.Width / 2)
   Picture1.Top = (Screen.Height / 2) - (Picture1.Height / 2)
End Sub

Private Sub Picture1_Click()
  frmMain.Show
  Unload Me
End Sub
