VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MacJeopardy Help Screen"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   11145
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6060
      Left            =   450
      TabIndex        =   1
      Top             =   525
      Width           =   10395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close Help"
      Height          =   450
      Left            =   4965
      TabIndex        =   0
      Top             =   6765
      Width           =   1260
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub Form_Load()
selFile = App.Path + "\MacJeopardyHelp.txt"
Open selFile For Input As #1   ' Open file for input.
Do While Not EOF(1)   ' Check for end of file.
   Line Input #1, Inputdata   ' Read line of data.
   mypos = InStr(4, Inputdata, ",", 1)
     List1.AddItem (Inputdata)
Loop
Close #1   ' Close file.
End Sub
