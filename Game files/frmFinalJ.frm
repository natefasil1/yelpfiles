VERSION 5.00
Begin VB.Form frmFinalJ 
   BackColor       =   &H00800000&
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   9915
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1470
      Top             =   1260
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   2325
      Left            =   1260
      TabIndex        =   1
      Top             =   5295
      Visible         =   0   'False
      Width           =   7470
      Begin VB.CommandButton Command10 
         Caption         =   "Wrong"
         Height          =   420
         Left            =   5730
         TabIndex        =   10
         Top             =   1620
         Width           =   840
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Wrong"
         Height          =   420
         Left            =   3630
         TabIndex        =   9
         Top             =   1620
         Width           =   840
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Wrong"
         Height          =   420
         Left            =   1410
         TabIndex        =   8
         Top             =   1620
         Width           =   840
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Correct"
         Height          =   420
         Left            =   4845
         TabIndex        =   7
         Top             =   1620
         Width           =   840
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Correct"
         Height          =   420
         Left            =   2745
         TabIndex        =   6
         Top             =   1620
         Width           =   840
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Correct"
         Height          =   420
         Left            =   510
         TabIndex        =   5
         Top             =   1620
         Width           =   840
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   555
         Left            =   4770
         TabIndex        =   4
         Top             =   810
         Width           =   1950
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   555
         Left            =   2625
         TabIndex        =   3
         Top             =   810
         Width           =   1950
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   555
         Left            =   435
         TabIndex        =   2
         Top             =   810
         Width           =   1950
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4770
         TabIndex        =   13
         Top             =   225
         Width           =   1965
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2580
         TabIndex        =   12
         Top             =   225
         Width           =   1965
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   450
         TabIndex        =   11
         Top             =   210
         Width           =   1965
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   390
      Left            =   4515
      TabIndex        =   0
      Top             =   7875
      Width           =   1110
   End
   Begin VB.Image Image17 
      Height          =   2760
      Left            =   2505
      Picture         =   "frmFinalJ.frx":0000
      Stretch         =   -1  'True
      Top             =   2340
      Width           =   5040
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   2520
      TabIndex        =   14
      Top             =   2385
      Width           =   4980
   End
   Begin VB.Image Image2 
      Height          =   1215
      Left            =   2805
      Picture         =   "frmFinalJ.frx":18E23
      Stretch         =   -1  'True
      Top             =   1050
      Width           =   4500
   End
   Begin VB.Image Image12 
      Height          =   1020
      Left            =   2880
      Picture         =   "frmFinalJ.frx":2127A
      Stretch         =   -1  'True
      Top             =   255
      Width           =   4500
   End
End
Attribute VB_Name = "frmFinalJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XYZ As Integer
Private Sub Command1_Click()
  frmMain.Command13.Visible = True
  Unload Me
End Sub

Private Sub Form_Load()
frmFinalJ.Left = 200
frmFinalJ.Top = 550

         Image17.Visible = True
         Label15.Caption = "Final Category" + vbCrLf + finalCategory
         Label16.Caption = contestant1
         Label17.Caption = contestant2
         Label18.Caption = contestant3
         'reset final bids to 0
         Text6.Text = ""
         Text7.Text = ""
         Text8.Text = ""
End Sub

Private Sub Image17_Click()
      Image17.Visible = False
      Call ding
      Frame10.Visible = True
      gameOver = True
End Sub

Private Sub Label15_Click()
    Label15.Caption = finalQuestion
    Call mySounds("Jeopardy_Music")
    XYZ = 0
    Timer6.Enabled = True
End Sub

Private Sub Timer6_Timer()
If XYZ = 30 Then
         Timer6.Enabled = False
     End If
      XYZ = XYZ + 1
End Sub
Private Sub mySounds(sndName)
     soundname = App.Path + "/sounds/" + sndName + ".wav"
     gbResults = PlaySound(soundname, 0, SND_ASYNC)
End Sub

Private Sub ding()
     soundname = App.Path + "/sounds/ding.wav"
     gbResults = PlaySound(soundname, 0, SND_ASYNC)
End Sub
Private Sub Command5_Click()
   frmMain.Label10.Caption = Str(Val(frmMain.Label10.Caption) + Val(Text6.Text))
End Sub

Private Sub Command6_Click()
   frmMain.Label11.Caption = Str(Val(frmMain.Label11.Caption) + Val(Text7.Text))
End Sub

Private Sub Command7_Click()
    frmMain.Label12.Caption = Str(Val(frmMain.Label12.Caption) + Val(Text8.Text))
End Sub

Private Sub Command8_Click()
   frmMain.Label10.Caption = Str(Val(frmMain.Label10.Caption) - Val(Text6.Text))
End Sub

Private Sub Command9_Click()
   frmMain.Label11.Caption = Str(Val(frmMain.Label11.Caption) - Val(Text7.Text))
End Sub
Private Sub Command10_Click()
   frmMain.Label12.Caption = Str(Val(frmMain.Label12.Caption) - Val(Text8.Text))
End Sub
