VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   $"frmMain.frx":0000
   ClientHeight    =   9885
   ClientLeft      =   -15
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   9885
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4140
      Top             =   165
   End
   Begin VB.Frame frame12 
      BorderStyle     =   0  'None
      Caption         =   "a"
      Height          =   9330
      Left            =   3975
      TabIndex        =   82
      Top             =   9660
      Visible         =   0   'False
      Width           =   13680
      Begin VB.CommandButton Command15 
         Caption         =   "OK"
         Height          =   375
         Left            =   6480
         TabIndex        =   105
         Top             =   7320
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Enabled         =   0   'False
         Height          =   420
         Index           =   3
         Left            =   1950
         TabIndex        =   97
         Top             =   6465
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.OptionButton Option1 
         Enabled         =   0   'False
         Height          =   420
         Index           =   2
         Left            =   1950
         TabIndex        =   96
         Top             =   6045
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.OptionButton Option1 
         Enabled         =   0   'False
         Height          =   420
         Index           =   1
         Left            =   1950
         TabIndex        =   95
         Top             =   5610
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3720
         Left            =   1245
         TabIndex        =   84
         Top             =   1545
         Width           =   9270
         Begin VB.Frame frame1shape 
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   0
            Left            =   330
            TabIndex        =   101
            Top             =   3180
            Width           =   585
         End
         Begin VB.Label Label41 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8100
            TabIndex        =   108
            Top             =   2270
            Width           =   375
         End
         Begin VB.Label Label40 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7750
            TabIndex        =   107
            Top             =   2270
            Width           =   375
         End
         Begin VB.Label Label39 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7400
            TabIndex        =   106
            Top             =   2270
            Width           =   375
         End
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl-L"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   285
            TabIndex        =   102
            Top             =   2655
            Width           =   630
         End
         Begin VB.Shape shiftrshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   4725
            Shape           =   4  'Rounded Rectangle
            Top             =   2235
            Width           =   555
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Space"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1860
            TabIndex        =   91
            Top             =   2640
            Width           =   1890
         End
         Begin VB.Label Label30 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Ctrl-R"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5190
            TabIndex        =   90
            Top             =   2655
            Width           =   630
         End
         Begin VB.Label Label29 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Shift-L"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   315
            TabIndex        =   89
            Top             =   2295
            Width           =   840
         End
         Begin VB.Shape esc 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   285
            Shape           =   4  'Rounded Rectangle
            Top             =   675
            Width           =   375
         End
         Begin VB.Shape f2 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   1365
            Shape           =   4  'Rounded Rectangle
            Top             =   675
            Width           =   375
         End
         Begin VB.Shape f3 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   1725
            Shape           =   4  'Rounded Rectangle
            Top             =   675
            Width           =   375
         End
         Begin VB.Shape f4 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   2085
            Shape           =   4  'Rounded Rectangle
            Top             =   675
            Width           =   375
         End
         Begin VB.Shape f5 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   2685
            Shape           =   4  'Rounded Rectangle
            Top             =   675
            Width           =   375
         End
         Begin VB.Shape f6 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   3045
            Shape           =   4  'Rounded Rectangle
            Top             =   675
            Width           =   375
         End
         Begin VB.Shape f7 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   3405
            Shape           =   4  'Rounded Rectangle
            Top             =   675
            Width           =   375
         End
         Begin VB.Shape f8 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   3765
            Shape           =   4  'Rounded Rectangle
            Top             =   675
            Width           =   375
         End
         Begin VB.Shape f9 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   4365
            Shape           =   4  'Rounded Rectangle
            Top             =   675
            Width           =   375
         End
         Begin VB.Shape f10 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   4725
            Shape           =   4  'Rounded Rectangle
            Top             =   675
            Width           =   375
         End
         Begin VB.Shape f11 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   5085
            Shape           =   4  'Rounded Rectangle
            Top             =   675
            Width           =   375
         End
         Begin VB.Shape f12 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   5445
            Shape           =   4  'Rounded Rectangle
            Top             =   675
            Width           =   375
         End
         Begin VB.Shape prntscr 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   6045
            Shape           =   4  'Rounded Rectangle
            Top             =   675
            Width           =   375
         End
         Begin VB.Shape scrolllock 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   6405
            Shape           =   4  'Rounded Rectangle
            Top             =   675
            Width           =   375
         End
         Begin VB.Shape pause 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   6765
            Shape           =   4  'Rounded Rectangle
            Top             =   675
            Width           =   375
         End
         Begin VB.Shape tildashape 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   285
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   375
         End
         Begin VB.Shape br1shape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   645
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   375
         End
         Begin VB.Shape br2shape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   1005
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   375
         End
         Begin VB.Shape br3shape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   1380
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   375
         End
         Begin VB.Shape br4shape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   1725
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   375
         End
         Begin VB.Shape br5shape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   2100
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   375
         End
         Begin VB.Shape br6shape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   2445
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   375
         End
         Begin VB.Shape br7shape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   2805
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   375
         End
         Begin VB.Shape br8shape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   3165
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   375
         End
         Begin VB.Shape br9shape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   3525
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   375
         End
         Begin VB.Shape br0shape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   3885
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   375
         End
         Begin VB.Shape dashshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   4245
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   375
         End
         Begin VB.Shape plusshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   4605
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   375
         End
         Begin VB.Shape back 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   4965
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   855
         End
         Begin VB.Shape tab1 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   285
            Shape           =   4  'Rounded Rectangle
            Top             =   1515
            Width           =   615
         End
         Begin VB.Shape qshape 
            BackColor       =   &H00000000&
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   885
            Shape           =   4  'Rounded Rectangle
            Top             =   1515
            Width           =   375
         End
         Begin VB.Shape wshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   1245
            Shape           =   4  'Rounded Rectangle
            Top             =   1515
            Width           =   375
         End
         Begin VB.Shape eshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   1605
            Shape           =   4  'Rounded Rectangle
            Top             =   1515
            Width           =   375
         End
         Begin VB.Shape rshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   1965
            Shape           =   4  'Rounded Rectangle
            Top             =   1515
            Width           =   375
         End
         Begin VB.Shape tshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   2325
            Shape           =   4  'Rounded Rectangle
            Top             =   1515
            Width           =   375
         End
         Begin VB.Shape yshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   2685
            Shape           =   4  'Rounded Rectangle
            Top             =   1515
            Width           =   375
         End
         Begin VB.Shape ushape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   3045
            Shape           =   4  'Rounded Rectangle
            Top             =   1515
            Width           =   375
         End
         Begin VB.Shape ishape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   3405
            Shape           =   4  'Rounded Rectangle
            Top             =   1515
            Width           =   375
         End
         Begin VB.Shape oshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   3765
            Shape           =   4  'Rounded Rectangle
            Top             =   1515
            Width           =   375
         End
         Begin VB.Shape pshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   4125
            Shape           =   4  'Rounded Rectangle
            Top             =   1515
            Width           =   375
         End
         Begin VB.Shape š 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   4485
            Shape           =   4  'Rounded Rectangle
            Top             =   1515
            Width           =   375
         End
         Begin VB.Shape ð 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   4845
            Shape           =   4  'Rounded Rectangle
            Top             =   1515
            Width           =   375
         End
         Begin VB.Shape capslock 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   285
            Shape           =   4  'Rounded Rectangle
            Top             =   1875
            Width           =   735
         End
         Begin VB.Shape ashape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   1005
            Shape           =   4  'Rounded Rectangle
            Top             =   1875
            Width           =   375
         End
         Begin VB.Shape sshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   1365
            Shape           =   4  'Rounded Rectangle
            Top             =   1875
            Width           =   375
         End
         Begin VB.Shape dshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   1725
            Shape           =   4  'Rounded Rectangle
            Top             =   1875
            Width           =   375
         End
         Begin VB.Shape fshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   2085
            Shape           =   4  'Rounded Rectangle
            Top             =   1875
            Width           =   375
         End
         Begin VB.Shape gshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   2445
            Shape           =   4  'Rounded Rectangle
            Top             =   1875
            Width           =   375
         End
         Begin VB.Shape hshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   2805
            Shape           =   4  'Rounded Rectangle
            Top             =   1875
            Width           =   375
         End
         Begin VB.Shape jshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   3165
            Shape           =   4  'Rounded Rectangle
            Top             =   1875
            Width           =   375
         End
         Begin VB.Shape kshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   3525
            Shape           =   4  'Rounded Rectangle
            Top             =   1875
            Width           =   375
         End
         Begin VB.Shape lshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   3885
            Shape           =   4  'Rounded Rectangle
            Top             =   1875
            Width           =   375
         End
         Begin VB.Shape Colonshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   4245
            Shape           =   4  'Rounded Rectangle
            Top             =   1875
            Width           =   375
         End
         Begin VB.Shape quoteshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   4605
            Shape           =   4  'Rounded Rectangle
            Top             =   1875
            Width           =   375
         End
         Begin VB.Shape enter 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   735
            Left            =   5325
            Shape           =   4  'Rounded Rectangle
            Top             =   1515
            Width           =   495
         End
         Begin VB.Shape shiftl 
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   285
            Shape           =   4  'Rounded Rectangle
            Top             =   2235
            Width           =   840
         End
         Begin VB.Shape zshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   1125
            Shape           =   4  'Rounded Rectangle
            Top             =   2235
            Width           =   375
         End
         Begin VB.Shape xshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   1485
            Shape           =   4  'Rounded Rectangle
            Top             =   2235
            Width           =   375
         End
         Begin VB.Shape cshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   1845
            Shape           =   4  'Rounded Rectangle
            Top             =   2235
            Width           =   375
         End
         Begin VB.Shape vshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   2205
            Shape           =   4  'Rounded Rectangle
            Top             =   2235
            Width           =   375
         End
         Begin VB.Shape bshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   2565
            Shape           =   4  'Rounded Rectangle
            Top             =   2235
            Width           =   375
         End
         Begin VB.Shape nshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   2925
            Shape           =   4  'Rounded Rectangle
            Top             =   2235
            Width           =   375
         End
         Begin VB.Shape mshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   3300
            Shape           =   4  'Rounded Rectangle
            Top             =   2235
            Width           =   375
         End
         Begin VB.Shape lessthanshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   3645
            Shape           =   4  'Rounded Rectangle
            Top             =   2235
            Width           =   375
         End
         Begin VB.Shape greaterthanshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   4005
            Shape           =   4  'Rounded Rectangle
            Top             =   2235
            Width           =   375
         End
         Begin VB.Shape qmarkshape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   4365
            Shape           =   4  'Rounded Rectangle
            Top             =   2235
            Width           =   375
         End
         Begin VB.Shape shiftr 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   5325
            Shape           =   4  'Rounded Rectangle
            Top             =   2235
            Width           =   495
         End
         Begin VB.Shape insert 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   6045
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   375
         End
         Begin VB.Shape home 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   6405
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   375
         End
         Begin VB.Shape pgup 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   6765
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   375
         End
         Begin VB.Shape delete 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   6045
            Shape           =   4  'Rounded Rectangle
            Top             =   1515
            Width           =   375
         End
         Begin VB.Shape end1 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   6405
            Shape           =   4  'Rounded Rectangle
            Top             =   1515
            Width           =   375
         End
         Begin VB.Shape pgdwn 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   6765
            Shape           =   4  'Rounded Rectangle
            Top             =   1515
            Width           =   375
         End
         Begin VB.Shape levo 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   6045
            Shape           =   4  'Rounded Rectangle
            Top             =   2595
            Width           =   375
         End
         Begin VB.Shape dole 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   6405
            Shape           =   4  'Rounded Rectangle
            Top             =   2595
            Width           =   375
         End
         Begin VB.Shape desno 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   6765
            Shape           =   4  'Rounded Rectangle
            Top             =   2595
            Width           =   375
         End
         Begin VB.Shape gore 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   6405
            Shape           =   4  'Rounded Rectangle
            Top             =   2235
            Width           =   375
         End
         Begin VB.Shape ctrl1shape 
            FillColor       =   &H8000000F&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   285
            Shape           =   4  'Rounded Rectangle
            Top             =   2595
            Width           =   615
         End
         Begin VB.Shape alt 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   1365
            Shape           =   4  'Rounded Rectangle
            Top             =   2595
            Width           =   495
         End
         Begin VB.Shape Shape83 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   885
            Shape           =   4  'Rounded Rectangle
            Top             =   2595
            Width           =   495
         End
         Begin VB.Shape spaceshape 
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   1845
            Shape           =   4  'Rounded Rectangle
            Top             =   2595
            Width           =   1935
         End
         Begin VB.Shape Shape85 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   4245
            Shape           =   4  'Rounded Rectangle
            Top             =   2595
            Width           =   495
         End
         Begin VB.Shape altgr 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   3765
            Shape           =   4  'Rounded Rectangle
            Top             =   2595
            Width           =   495
         End
         Begin VB.Shape Shape87 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   4725
            Shape           =   4  'Rounded Rectangle
            Top             =   2595
            Width           =   495
         End
         Begin VB.Shape ctrl2shape 
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   5205
            Shape           =   4  'Rounded Rectangle
            Top             =   2595
            Width           =   615
         End
         Begin VB.Shape numlock 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   7365
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   375
         End
         Begin VB.Shape slash 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   7725
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   375
         End
         Begin VB.Shape zvezda 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   8085
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   375
         End
         Begin VB.Shape num7 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   7365
            Shape           =   4  'Rounded Rectangle
            Top             =   1515
            Width           =   375
         End
         Begin VB.Shape num8 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   7725
            Shape           =   4  'Rounded Rectangle
            Top             =   1515
            Width           =   375
         End
         Begin VB.Shape num9 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   8085
            Shape           =   4  'Rounded Rectangle
            Top             =   1515
            Width           =   375
         End
         Begin VB.Shape minus 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   8445
            Shape           =   4  'Rounded Rectangle
            Top             =   1155
            Width           =   375
         End
         Begin VB.Shape plus2 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   735
            Left            =   8445
            Shape           =   4  'Rounded Rectangle
            Top             =   1515
            Width           =   375
         End
         Begin VB.Shape num4 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   7365
            Shape           =   4  'Rounded Rectangle
            Top             =   1875
            Width           =   375
         End
         Begin VB.Shape num5 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   7725
            Shape           =   4  'Rounded Rectangle
            Top             =   1875
            Width           =   375
         End
         Begin VB.Shape num6 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   8085
            Shape           =   4  'Rounded Rectangle
            Top             =   1875
            Width           =   375
         End
         Begin VB.Shape num1 
            FillColor       =   &H0000C000&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   7365
            Shape           =   4  'Rounded Rectangle
            Top             =   2235
            Width           =   375
         End
         Begin VB.Shape num2 
            FillColor       =   &H0000C000&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   7725
            Shape           =   4  'Rounded Rectangle
            Top             =   2235
            Width           =   375
         End
         Begin VB.Shape num3 
            FillColor       =   &H0000C000&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   8085
            Shape           =   4  'Rounded Rectangle
            Top             =   2235
            Width           =   375
         End
         Begin VB.Shape enter1 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   735
            Left            =   8445
            Shape           =   4  'Rounded Rectangle
            Top             =   2235
            Width           =   375
         End
         Begin VB.Shape del 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   8085
            Shape           =   4  'Rounded Rectangle
            Top             =   2595
            Width           =   375
         End
         Begin VB.Shape num0 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   7365
            Shape           =   4  'Rounded Rectangle
            Top             =   2595
            Width           =   735
         End
         Begin VB.Shape f1 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   1005
            Shape           =   4  'Rounded Rectangle
            Top             =   675
            Width           =   375
         End
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5640
         TabIndex        =   100
         Top             =   5670
         Width           =   1215
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5640
         TabIndex        =   99
         Top             =   6135
         Width           =   1215
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5640
         TabIndex        =   98
         Top             =   6600
         Width           =   1215
      End
      Begin VB.Label Label34 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ctrl-R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8610
         TabIndex        =   94
         Top             =   6600
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label Label33 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Space"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8610
         TabIndex        =   93
         Top             =   6135
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label Label32 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Shift-L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8610
         TabIndex        =   92
         Top             =   5640
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Contestant # 3 (Right)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2430
         TabIndex        =   88
         Top             =   6600
         Width           =   2715
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Contestant # 2 (Center)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2430
         TabIndex        =   87
         Top             =   6120
         Width           =   3435
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Contestant # 1 (Left)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2430
         TabIndex        =   86
         Top             =   5655
         Width           =   2715
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":009D
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   2565
         TabIndex        =   85
         Top             =   405
         Width           =   7425
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label24 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   83
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3465
      Top             =   150
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   6705
      Left            =   4500
      TabIndex        =   103
      Top             =   60
      Visible         =   0   'False
      Width           =   7425
      Begin VB.Frame Frame15 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   210
         TabIndex        =   104
         Top             =   180
         Width           =   7080
         Begin VB.Image Image21 
            Height          =   6240
            Left            =   60
            Stretch         =   -1  'True
            Top             =   45
            Width           =   6900
         End
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Reset"
      Height          =   420
      Left            =   12150
      TabIndex        =   66
      Top             =   8925
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Winner"
      Height          =   330
      Left            =   12150
      TabIndex        =   78
      Top             =   9330
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command14 
      Height          =   315
      Left            =   15000
      TabIndex        =   81
      Top             =   15
      Width           =   270
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   1620
      Left            =   4890
      TabIndex        =   45
      Top             =   3780
      Visible         =   0   'False
      Width           =   3855
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   90
         TabIndex        =   67
         Top             =   225
         Width           =   1995
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Next"
         Height          =   390
         Left            =   3045
         TabIndex        =   60
         Top             =   240
         Width           =   570
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Prev"
         Height          =   375
         Left            =   3045
         TabIndex        =   59
         Top             =   615
         Width           =   570
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Go"
         Height          =   480
         Left            =   1845
         TabIndex        =   58
         Top             =   1095
         Width           =   480
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   2130
         TabIndex        =   57
         Text            =   "1"
         Top             =   225
         Width           =   900
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   2880
      Left            =   10290
      TabIndex        =   68
      Top             =   150
      Visible         =   0   'False
      Width           =   4830
      Begin VB.CommandButton Command12 
         Caption         =   "Update"
         Height          =   375
         Left            =   2010
         TabIndex        =   75
         Top             =   2430
         Width           =   1035
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         Left            =   3090
         TabIndex        =   74
         Top             =   1710
         Width           =   1050
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         Left            =   1800
         TabIndex        =   73
         Top             =   1710
         Width           =   1050
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   495
         TabIndex        =   72
         Top             =   1710
         Width           =   1050
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Change Scores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   510
         Left            =   960
         TabIndex        =   79
         Top             =   240
         Width           =   2955
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Script"
            Size            =   14.25
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   2985
         TabIndex        =   71
         Top             =   960
         Width           =   1320
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Script"
            Size            =   14.25
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1635
         TabIndex        =   70
         Top             =   960
         Width           =   1320
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Script"
            Size            =   14.25
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   285
         TabIndex        =   69
         Top             =   960
         Width           =   1320
      End
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2835
      Top             =   135
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   2205
      Top             =   120
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   3500
      Left            =   1590
      Top             =   150
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   150
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   150
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   3420
      Left            =   3240
      TabIndex        =   61
      Top             =   2430
      Visible         =   0   'False
      Width           =   4515
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1950
         TabIndex        =   62
         Top             =   2820
         Width           =   1080
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   64
         Top             =   2940
         Width           =   315
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         Height          =   435
         Left            =   3075
         Shape           =   3  'Circle
         Top             =   2835
         Width           =   390
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Wager:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   645
         TabIndex        =   63
         Top             =   2850
         Width           =   1275
      End
      Begin VB.Image Image7 
         Height          =   2505
         Left            =   360
         Picture         =   "frmMain.frx":013D
         Stretch         =   -1  'True
         Top             =   195
         Width           =   3630
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   2685
         Left            =   285
         Top             =   120
         Width           =   3840
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   2715
      Left            =   10185
      TabIndex        =   46
      Top             =   255
      Visible         =   0   'False
      Width           =   4890
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
         Height          =   420
         Left            =   2055
         TabIndex        =   53
         Top             =   2220
         Width           =   765
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2565
         MaxLength       =   8
         TabIndex        =   49
         Top             =   1500
         Width           =   1440
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2565
         MaxLength       =   8
         TabIndex        =   48
         Top             =   840
         Width           =   1440
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2565
         MaxLength       =   8
         TabIndex        =   47
         Top             =   225
         Width           =   1440
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Contestant # 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   330
         TabIndex        =   52
         Top             =   1530
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Contestant # 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   330
         TabIndex        =   51
         Top             =   900
         Width           =   1980
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Contestant # 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   330
         TabIndex        =   50
         Top             =   270
         Width           =   1980
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   7890
      Left            =   555
      TabIndex        =   0
      Top             =   8700
      Width           =   9990
      Begin VB.Image Image10 
         Height          =   7575
         Left            =   225
         Picture         =   "frmMain.frx":104B4
         Stretch         =   -1  'True
         Top             =   195
         Width           =   9690
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   6645
      Left            =   585
      TabIndex        =   8
      Top             =   1890
      Width           =   9510
      Begin VB.Frame Frame5 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   4530
         Left            =   1320
         TabIndex        =   39
         Top             =   1620
         Visible         =   0   'False
         Width           =   5565
         Begin VB.Image Image8 
            BorderStyle     =   1  'Fixed Single
            Height          =   465
            Left            =   3660
            Picture         =   "frmMain.frx":45271
            Stretch         =   -1  'True
            Top             =   3945
            Width           =   1815
         End
         Begin VB.Image Image6 
            BorderStyle     =   1  'Fixed Single
            Height          =   465
            Left            =   1860
            Picture         =   "frmMain.frx":5E094
            Stretch         =   -1  'True
            Top             =   3960
            Width           =   1815
         End
         Begin VB.Image Image3 
            BorderStyle     =   1  'Fixed Single
            Height          =   465
            Left            =   75
            Picture         =   "frmMain.frx":60153
            Stretch         =   -1  'True
            Top             =   3945
            Width           =   1740
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "dddd"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3660
            Left            =   255
            TabIndex        =   40
            Top             =   120
            Width           =   5145
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         Height          =   825
         Left            =   1515
         Top             =   945
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   26
         Left            =   8040
         TabIndex        =   38
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   21
         Left            =   6480
         TabIndex        =   37
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   16
         Left            =   4935
         TabIndex        =   36
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   11
         Left            =   3345
         TabIndex        =   35
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   6
         Left            =   1725
         TabIndex        =   34
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   1
         Left            =   240
         TabIndex        =   33
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "500"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   30
         Left            =   8040
         TabIndex        =   32
         Top             =   5595
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "400"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   29
         Left            =   8040
         TabIndex        =   31
         Top             =   4440
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "300"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   28
         Left            =   8040
         TabIndex        =   30
         Top             =   3165
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   27
         Left            =   8010
         TabIndex        =   29
         Top             =   1830
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "500"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   25
         Left            =   6480
         TabIndex        =   28
         Top             =   5595
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "400"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   24
         Left            =   6480
         TabIndex        =   27
         Top             =   4440
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "300"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   23
         Left            =   6480
         TabIndex        =   26
         Top             =   3165
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   22
         Left            =   6450
         TabIndex        =   25
         Top             =   1830
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "500"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   20
         Left            =   4935
         TabIndex        =   24
         Top             =   5595
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "400"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   19
         Left            =   4935
         TabIndex        =   23
         Top             =   4440
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "300"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   18
         Left            =   4935
         TabIndex        =   22
         Top             =   3165
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   17
         Left            =   4905
         TabIndex        =   21
         Top             =   1830
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "500"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   15
         Left            =   3345
         TabIndex        =   20
         Top             =   5595
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "400"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   14
         Left            =   3345
         TabIndex        =   19
         Top             =   4440
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "300"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   13
         Left            =   3345
         TabIndex        =   18
         Top             =   3165
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   12
         Left            =   3315
         TabIndex        =   17
         Top             =   1830
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "500"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   10
         Left            =   1725
         TabIndex        =   16
         Top             =   5595
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "400"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   9
         Left            =   1725
         TabIndex        =   15
         Top             =   4440
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "300"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   8
         Left            =   1725
         TabIndex        =   14
         Top             =   3165
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   7
         Left            =   1695
         TabIndex        =   13
         Top             =   1830
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "500"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   5
         Left            =   225
         TabIndex        =   12
         Top             =   5595
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "400"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   4
         Left            =   225
         TabIndex        =   11
         Top             =   4440
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "300"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   3
         Left            =   225
         TabIndex        =   10
         Top             =   3165
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   555
         Index           =   2
         Left            =   195
         TabIndex        =   9
         Top             =   1830
         Width           =   1155
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   1
         Left            =   15
         Picture         =   "frmMain.frx":6209B
         Stretch         =   -1  'True
         Top             =   210
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   2
         Left            =   15
         Picture         =   "frmMain.frx":629D4
         Stretch         =   -1  'True
         Top             =   1485
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   3
         Left            =   15
         Picture         =   "frmMain.frx":6330D
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   4
         Left            =   15
         Picture         =   "frmMain.frx":63C46
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   5
         Left            =   15
         Picture         =   "frmMain.frx":6457F
         Stretch         =   -1  'True
         Top             =   5295
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   6
         Left            =   1575
         Picture         =   "frmMain.frx":64EB8
         Stretch         =   -1  'True
         Top             =   210
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   7
         Left            =   1575
         Picture         =   "frmMain.frx":657F1
         Stretch         =   -1  'True
         Top             =   1485
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   8
         Left            =   1575
         Picture         =   "frmMain.frx":6612A
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   9
         Left            =   1590
         Picture         =   "frmMain.frx":66A63
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   10
         Left            =   1575
         Picture         =   "frmMain.frx":6739C
         Stretch         =   -1  'True
         Top             =   5295
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   11
         Left            =   3150
         Picture         =   "frmMain.frx":67CD5
         Stretch         =   -1  'True
         Top             =   210
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   12
         Left            =   3150
         Picture         =   "frmMain.frx":6860E
         Stretch         =   -1  'True
         Top             =   1485
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   13
         Left            =   3150
         Picture         =   "frmMain.frx":68F47
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   14
         Left            =   3150
         Picture         =   "frmMain.frx":69880
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   15
         Left            =   3150
         Picture         =   "frmMain.frx":6A1B9
         Stretch         =   -1  'True
         Top             =   5295
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   16
         Left            =   4725
         Picture         =   "frmMain.frx":6AAF2
         Stretch         =   -1  'True
         Top             =   210
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   17
         Left            =   4725
         Picture         =   "frmMain.frx":6B42B
         Stretch         =   -1  'True
         Top             =   1485
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   18
         Left            =   4725
         Picture         =   "frmMain.frx":6BD64
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   19
         Left            =   4725
         Picture         =   "frmMain.frx":6C69D
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   20
         Left            =   4725
         Picture         =   "frmMain.frx":6CFD6
         Stretch         =   -1  'True
         Top             =   5295
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   21
         Left            =   6300
         Picture         =   "frmMain.frx":6D90F
         Stretch         =   -1  'True
         Top             =   210
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   22
         Left            =   6300
         Picture         =   "frmMain.frx":6E248
         Stretch         =   -1  'True
         Top             =   1485
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   23
         Left            =   6300
         Picture         =   "frmMain.frx":6EB81
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   24
         Left            =   6300
         Picture         =   "frmMain.frx":6F4BA
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   25
         Left            =   6300
         Picture         =   "frmMain.frx":6FDF3
         Stretch         =   -1  'True
         Top             =   5295
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   26
         Left            =   7875
         Picture         =   "frmMain.frx":7072C
         Stretch         =   -1  'True
         Top             =   210
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   27
         Left            =   7875
         Picture         =   "frmMain.frx":71065
         Stretch         =   -1  'True
         Top             =   1485
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   28
         Left            =   7875
         Picture         =   "frmMain.frx":7199E
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   29
         Left            =   7875
         Picture         =   "frmMain.frx":722D7
         Stretch         =   -1  'True
         Top             =   4020
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1275
         Index           =   30
         Left            =   7875
         Picture         =   "frmMain.frx":72C10
         Stretch         =   -1  'True
         Top             =   5295
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00800000&
      Height          =   1290
      Left            =   525
      TabIndex        =   1
      Top             =   585
      Width           =   9585
      Begin VB.Image Image1 
         Height          =   1035
         Index           =   36
         Left            =   7965
         Picture         =   "frmMain.frx":73549
         Stretch         =   -1  'True
         Top             =   195
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   1035
         Index           =   35
         Left            =   6390
         Picture         =   "frmMain.frx":8C36C
         Stretch         =   -1  'True
         Top             =   195
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   1035
         Index           =   34
         Left            =   4830
         Picture         =   "frmMain.frx":A518F
         Stretch         =   -1  'True
         Top             =   195
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   1035
         Index           =   33
         Left            =   3255
         Picture         =   "frmMain.frx":BDFB2
         Stretch         =   -1  'True
         Top             =   195
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   1035
         Index           =   32
         Left            =   1695
         Picture         =   "frmMain.frx":D6DD5
         Stretch         =   -1  'True
         Top             =   195
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   1035
         Index           =   31
         Left            =   120
         Picture         =   "frmMain.frx":EFBF8
         Stretch         =   -1  'True
         Top             =   195
         Width           =   1515
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1065
         Index           =   5
         Left            =   7950
         TabIndex        =   7
         Top             =   180
         Width           =   1545
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1065
         Index           =   4
         Left            =   6375
         TabIndex        =   6
         Top             =   180
         Width           =   1545
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1065
         Index           =   3
         Left            =   4812
         TabIndex        =   5
         Top             =   180
         Width           =   1545
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1065
         Index           =   2
         Left            =   3243
         TabIndex        =   4
         Top             =   180
         Width           =   1545
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1065
         Index           =   1
         Left            =   1674
         TabIndex        =   3
         Top             =   180
         Width           =   1545
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1065
         Index           =   0
         Left            =   105
         TabIndex        =   2
         Top             =   180
         Width           =   1545
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Caption         =   "If Len(Trim(Text2.Text)) = 0 Then"
      Height          =   2505
      Left            =   10035
      TabIndex        =   41
      Top             =   6105
      Width           =   4980
      Begin VB.TextBox Text9 
         BackColor       =   &H00C00000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   2.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   15
         Left            =   840
         TabIndex        =   65
         Top             =   1350
         Width           =   30
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   120
         Index           =   3
         Left            =   3555
         Top             =   435
         Width           =   1245
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   120
         Index           =   2
         Left            =   1935
         Top             =   435
         Width           =   1245
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00008000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00008000&
         Height          =   120
         Index           =   1
         Left            =   285
         Top             =   435
         Width           =   1245
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         BorderStyle     =   2  'Dash
         Height          =   405
         Left            =   690
         Shape           =   3  'Circle
         Top             =   1140
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         Height          =   375
         Index           =   3
         Left            =   4020
         Shape           =   3  'Circle
         Top             =   1890
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         Height          =   375
         Index           =   2
         Left            =   2355
         Shape           =   3  'Circle
         Top             =   1890
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         Height          =   375
         Index           =   1
         Left            =   705
         Shape           =   3  'Circle
         Top             =   1890
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   4140
         TabIndex        =   56
         Top             =   630
         Width           =   195
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2445
         TabIndex        =   55
         Top             =   645
         Width           =   195
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   825
         TabIndex        =   54
         Top             =   630
         Width           =   195
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Script"
            Size            =   20.25
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3450
         TabIndex        =   44
         Top             =   45
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Script"
            Size            =   20.25
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         TabIndex        =   43
         Top             =   45
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Script"
            Size            =   20.25
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   135
         TabIndex        =   42
         Top             =   30
         Width           =   1575
      End
      Begin VB.Image Image5 
         Height          =   2220
         Left            =   30
         Picture         =   "frmMain.frx":108A1B
         Stretch         =   -1  'True
         Top             =   330
         Width           =   5010
      End
   End
   Begin VB.Label Label16 
      Caption         =   "Music"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   14085
      TabIndex        =   112
      Top             =   9285
      Width           =   675
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Starter"
      Height          =   210
      Index           =   3
      Left            =   13770
      TabIndex        =   111
      Top             =   8640
      Width           =   990
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Starter"
      Height          =   210
      Index           =   2
      Left            =   12180
      TabIndex        =   110
      Top             =   8625
      Width           =   900
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Starter"
      Height          =   210
      Index           =   1
      Left            =   10455
      TabIndex        =   109
      Top             =   8640
      Width           =   990
   End
   Begin VB.Label Label21 
      Caption         =   $"frmMain.frx":113AA0
      Height          =   495
      Left            =   2715
      TabIndex        =   80
      Top             =   8880
      Visible         =   0   'False
      Width           =   5100
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "DOUBLE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   465
      Left            =   2865
      TabIndex        =   77
      Top             =   180
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "DOUBLE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   465
      Left            =   2910
      TabIndex        =   76
      Top             =   150
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Image Image20 
      Height          =   675
      Left            =   13755
      Picture         =   "frmMain.frx":113B27
      Stretch         =   -1  'True
      Top             =   3855
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image19 
      Height          =   675
      Left            =   12090
      Picture         =   "frmMain.frx":11628F
      Stretch         =   -1  'True
      Top             =   3885
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image18 
      Height          =   675
      Left            =   10530
      Picture         =   "frmMain.frx":1189F7
      Stretch         =   -1  'True
      Top             =   3900
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image16 
      Height          =   1740
      Left            =   13545
      Picture         =   "frmMain.frx":11B15F
      Stretch         =   -1  'True
      Top             =   4305
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image Image15 
      Height          =   1740
      Left            =   11910
      Picture         =   "frmMain.frx":11C90B
      Stretch         =   -1  'True
      Top             =   4320
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image Image14 
      Height          =   1740
      Left            =   10215
      Picture         =   "frmMain.frx":11DC5C
      Stretch         =   -1  'True
      Top             =   4305
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image Image13 
      Height          =   3585
      Left            =   13425
      Picture         =   "frmMain.frx":11F1E7
      Stretch         =   -1  'True
      Top             =   2445
      Width           =   1650
   End
   Begin VB.Image Image11 
      Height          =   3585
      Left            =   11790
      Picture         =   "frmMain.frx":12187A
      Stretch         =   -1  'True
      Top             =   2445
      Width           =   1650
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   4605
      Picture         =   "frmMain.frx":123F0D
      Stretch         =   -1  'True
      Top             =   180
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Image Image9 
      Height          =   3585
      Left            =   10155
      Picture         =   "frmMain.frx":1276BF
      Stretch         =   -1  'True
      Top             =   2445
      Width           =   1665
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSetup 
         Caption         =   "Setup"
         Begin VB.Menu mnuAssignNames 
            Caption         =   "Players' Names"
         End
         Begin VB.Menu mnuInputKeys 
            Caption         =   "Input Keys"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuChangeScores 
            Caption         =   "Score Corrections"
         End
         Begin VB.Menu mnuReset 
            Caption         =   "Reset (Name BG colors)"
         End
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuGames 
      Caption         =   "Games"
      Begin VB.Menu mnuSelectGame 
         Caption         =   "Jeopardy!"
      End
      Begin VB.Menu mnuDJ 
         Caption         =   "Double Jeopardy!"
      End
      Begin VB.Menu mnuFinalJeopardy 
         Caption         =   "Final Jeopardy!"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpGen 
         Caption         =   "General"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu menuMac 
      Caption         =   "Mac"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myArrayOfAnswers(31) As String
Dim aGameNames(100) As String
Dim gamenum As Integer
Dim myGameNumber As Integer
Dim maxgames As Integer
Dim myRandomNum As Integer
Dim myRandomNum2 As Integer
Dim myRandomNumDJ As Integer 'for dbljeopardy
Dim myRandomNumDJ2 As Integer
Dim cPlayer As String
Dim labelclicked As Integer
Dim theWager As Long
Dim valToAdd As Long
Dim Contestant1NewValue As Long
Dim Contestant2NewValue As Long
Dim Contestant3NewValue As Long
Dim dailyDouble As Boolean
Dim gameOver As Boolean
Dim xx As Integer
Dim yy As Integer
Dim numQleft As Integer
Dim myArrayOfPix(10) As String
Dim picNum1 As Integer
Dim picNum2 As Integer
Dim picNum3 As Integer
Dim z As Integer
Dim zz As Integer

Dim myz As Integer
Dim XYZ As Integer
Dim player1guessed As Boolean
Dim player2guessed As Boolean
Dim player3guessed As Boolean
Dim anyQs_Left As Boolean
Dim valOfLabel As Integer
Dim contestant As Integer
Dim someoneClicked As Boolean
Dim bDoubleJeopardy As Boolean
Dim DJ As Integer
Dim theLabelClicked As Integer 'used so will know which label4 was clicked
Dim ok2click_1 As Boolean
Dim ok2click_2 As Boolean
Dim ok2click_3 As Boolean
Dim CONTESTANT1KEYCODE As Integer
Dim CONTESTANT2KEYCODE As Integer
Dim CONTESTANT3KEYCODE As Integer
Dim c1 As Boolean
Dim c2 As Boolean
Dim c3 As Boolean
Dim NumberLock As Boolean


Private Declare Function GetKeyState Lib _
"user32" (ByVal nVirtKey As Long) As Integer


















Private Sub Command1_Click()
    If Len(Trim(Text2.Text)) = 0 Then
        Label1.Caption = "Mac"
    Else
        Label1.Caption = Trim(Text2.Text)
    End If
    If Len(Trim(Text3.Text)) = 0 Then
        Label3.Caption = "Sam"
    Else
        Label3.Caption = Trim(Text3.Text)
        
    End If
    If Len(Trim(Text4.Text)) = 0 Then
        Label6.Caption = "Socorro"
    Else
        Label6.Caption = Trim(Text4.Text)
    End If
    contestant1 = Label1.Caption
    contestant2 = Label3.Caption
    contestant3 = Label6.Caption
    Frame7.Visible = False
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
End Sub
Private Sub Command11_Click()
    Label1.BackColor = vbWhite
    Label1.ForeColor = vbBlack
    Label3.BackColor = vbWhite
    Label3.ForeColor = vbBlack
    Label6.BackColor = vbWhite
    Label6.ForeColor = vbBlack
    Shape5.Visible = False

'    Text9.SetFocus
'     ok2click_1 = False
'     Shape6(1).BackColor = vbRed
'     ok2click_2 = False
'     Shape6(2).BackColor = vbRed
'     ok2click_3 = False
'     Shape6(3).BackColor = vbRed
     
   Command14.SetFocus

End Sub

Private Sub Command12_Click()
    Label10.Caption = Trim(Text11(1).Text)
    Label11.Caption = Trim(Text11(2).Text)
    Label12.Caption = Trim(Text11(3).Text)
    Frame11.Visible = False
    For X = 1 To 3  'reset scores in this box to nothing
      Text11(X).Text = ""
    Next X
End Sub

Private Sub Command13_Click()
    val1 = Val(Label10.Caption)
    val2 = Val(Label11.Caption)
    val3 = Val(Label12.Caption)
    If val1 > val2 Then
       If val1 > val3 Then
           MsgBox Label1.Caption + " is the Winner!"
       End If
    End If
    If val2 > val3 Then
       If val2 > val1 Then
           MsgBox Label3.Caption + " is the Winner!"
       End If
    End If
    If val3 > val2 Then
       If val3 > val1 Then
           MsgBox Label6.Caption + " is the Winner!"
       End If
    End If
'    mnuDJ.Enabled = False
    mnuFinalJeopardy.Enabled = False
End Sub

Private Sub Command15_Click()
   frame12.Visible = False
 End Sub

Private Sub Command2_Click()
    If Val(Text1.Text) < maxgames Then
        myGameNumber = Val(Text1.Text) + 1
        Text1.Text = Str(Val(Text1.Text) + 1)
        
        Call getGameName(myGameNumber, 0)
    Else
        MsgBox "There are only " + Str(maxgames) + " in this version."
    End If
End Sub

Private Sub Command3_Click()
    If Val(Text1.Text) > 1 Then
    myGameNumber = Val(Text1.Text) - 1
        Text1.Text = Str(Val(Text1.Text) - 1)
        Call getGameName(myGameNumber, 0)
    Else
        MsgBox "This is the first game."
    End If
End Sub

Private Sub Command4_Click()
    gamenum = Val(Trim(Text1.Text))
    If gamenum = 4 Then     'for St Pats Day
       For X = 1 To 30
          Image1(X).Picture = LoadPicture(App.Path + "/images/greenbg.jpg")
       Next X
       frmMain.BackColor = &H8000&    'GREEN
       Frame7.BackColor = &H8000&
       Image18.Visible = True   'Leprechaun Hats
       Image19.Visible = True
       Image20.Visible = True
    Else
        For X = 1 To 30
          Image1(X).Picture = LoadPicture(App.Path + "/images/jeopardybluebg.jpg")
       Next X
       frmMain.BackColor = &H800000    'Dk Blue
       Frame7.BackColor = &H800000
       Image18.Visible = False   'Leprechaun Hats NOT Visible
       Image19.Visible = False
       Image20.Visible = False
    End If
    Call setCategories(gamenum, DJ)
  
    Call fillboard
    numQleft = 30
    For X = 1 To 30   'set the numbers to clickable once game starts
        Label4(X).Enabled = True
        Label4(X).Visible = True  'reset to new game
    Next X
    
    Randomize
    myRandomNum = Int((30 * Rnd) + 1)    ' Generate random value between 1 and 30.
                                       'will use this to put in daily doubles
    myRandomNum2 = Int((30 * Rnd) + 1)
    While myRandomNum = myRandomNum2
        myRandomNum2 = Int((30 * Rnd) + 1)
    Wend
    
    Randomize
    myRandomNumDJ = Int((7 * Rnd + 1))  'used for the second random number in DBL Jeopardy
    myRandomNum3 = Int((7 * Rnd) + 1)
    While myRandomNumDJ = myRandomNum3
        myRandomNum3 = Int((7 * Rnd) + 1)
    Wend
    Frame2.Visible = False
    Frame6.Visible = False
    theWager = 0
    labelclicked = 0
    If DJ = 0 Then  'only do this if NOT in Double Jeopardy game...this changes to
                      'highest contestant at beginning of dbljeopardy
        Shape4(1).Visible = True  'this is so contestant 1 goes first
        Shape4(2).Visible = False
        Shape4(3).Visible = False
    End If
    Call popArray   'set up all answers
    Timer6.Enabled = False
'    Text6.Text = ""
'    Text7.Text = ""
'    Text8.Text = ""
    
End Sub
Private Sub fillboard()
        soundname = App.Path + "/sounds/jboardfill.wav"
        gbResults = PlaySound(soundname, 0, SND_ASYNC)
        xx = 1
        For X = 1 To 30 Step 2
            Timer1.Enabled = True
        Next X

End Sub
Private Sub fillboard2()
        yy = 30
        For X = 1 To 30 Step 2
            Timer5.Enabled = True
        Next X
End Sub
Private Sub setCategories(gamenum, DJ)
   Dim rs As ADODB.Recordset
'   Set cnn = DataEnvironment1.Connection1
  Call dbConnection  'located in Module1
   Set cmd = New ADODB.Command
'   cnn.Open
   Set cmd.ActiveConnection = cnn
   cmd.CommandText = "select * from jeopardy "
   cmd.CommandText = cmd.CommandText + "where gameNumber = " & Str(gamenum) & " and DBLJ = " & Str(DJ)

   Set rs = cmd.Execute()
    'set the categories at the top---------
       Label2(0).Caption = Trim(rs!cat1)
       Label2(1).Caption = Trim(rs!cat2)
       Label2(2).Caption = Trim(rs!cat3)
       Label2(3).Caption = Trim(rs!cat4)
       Label2(4).Caption = Trim(rs!cat5)
       Label2(5).Caption = Trim(rs!cat6)
    
   rs.Close
   cnn.Close
End Sub


Private Sub Form_Load()
Call setupKeycodes

For X = 1 To 30   'set the numbers to non-clickable until game starts
   Label4(X).Enabled = False
Next X
'stop all from clicking in until question is asked
ok2click_1 = False
Shape6(1).BackColor = vbRed
ok2click_2 = False
Shape6(2).BackColor = vbRed
ok2click_3 = False
Shape6(3).BackColor = vbRed

bDoubleJeopardy = False
mnuFinalJeopardy.Enabled = False
'mnuFinalJeopardy.Enabled = True   'temporary....
'mnuDJ.Enabled = False
Frame2.Left = 150
Frame2.Top = 600
myGameNumber = 1
gameOver = False
Command13.Visible = False
Dim rs As ADODB.Recordset
'Dim cnn As ADODB.Connection
''cnn = "C:\Program Files\Jeopardy\Jeopardy.mdb;DefaultDir=C:\Program Files\Jeopardy;Driver={Driver do Microsoft Access (*.mdb)};DriverId=25;FIL=MS Access;FILEDSN=C:\Program Files\Jeopardy\jeopardy.dsn;MaxBufferSize=2048;MaxScanRows=8;PageTimeout=5;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"
''   Set cnn = DataEnvironment1.Connection1
'
'
'   Set cnn = New ADODB.Connection

  Call dbConnection
'      '************************************
'      With cnn
'       .Provider = "Microsoft.Jet.OLEDB.4.0"
'       .ConnectionString = "User ID=Admin;password= ;" & " Data Source=" & App.Path & "\jeopardy.mdb;"
'       .CursorLocation = adUseClient
'       .Open
'      End With
'
      
      '************************************
'   cnn.Open
  Set cmd = New ADODB.Command
   Set cmd.ActiveConnection = cnn
   cmd.CommandText = "select distinct gamenumber from jeopardy "
   Set rs = cmd.Execute()
     maxgames = rs.RecordCount
   rs.Close
   cnn.Close
   Call mySounds("backgroundmusic")
  Dim rs1 As ADODB.Recordset
'   Set cnn = DataEnvironment1.Connection1
    Call dbConnection
   Set cmd = New ADODB.Command
'   cnn.Open
   Set cmd.ActiveConnection = cnn
   cmd.CommandText = "select * from jeopardy where gameNumber = 1 and DBLJ = 0"
   Set rs1 = cmd.Execute()
     mygameName = Trim(rs1!gameName)
     Text10.Text = mygameName
   rs1.Close
   cnn.Close
   someoneClicked = False
   
   'Set up pictures right from the start
     picNum1 = 1
     picNum2 = 2
     picNum3 = 3
     Call setPix
     Image14.Picture = LoadPicture(App.Path + myArrayOfPix(picNum1))
     Image15.Picture = LoadPicture(App.Path + myArrayOfPix(picNum2))
     Image16.Picture = LoadPicture(App.Path + myArrayOfPix(picNum3))
     Image14.Visible = True
     Image15.Visible = True
     Image16.Visible = True
End Sub
Public Function NumLockOn() As Boolean
'    Dim iKeyState As Integer
'    iKeyState = GetKeyState(vbKeyNumlock)
'    NumLockOn = (iKeyState = 1 Or iKeyState = -127)
'    If iKeyState = 1 Or iKeyState = -127 Then 'prompt user to hit numlock key
'       Shape7.Visible = True
'    Else
'       Shape7.Visible = False
'    End If
End Function

Private Sub setupKeycodes()
'use this function to change which keys you want for which contestant
'DEFAULTS:
'Had problems with the numpads because of numlock--going back to SHIFT-SPACE-CTRL Keys
'    CONTESTANT1KEYCODE = 97 'this is the '1' key on the keypad
'    CONTESTANT2KEYCODE = 98 'this is the '2' key on the keypad
'    CONTESTANT3KEYCODE = 99 'this is the '3' key on the keypad
    'or you can use the setupkeycode menu option to use SHIFT-SPACE-CTRL
    CONTESTANT1KEYCODE = 16 'this is the 'SHIFT' key on the keypad
    CONTESTANT2KEYCODE = 32 'this is the 'SPACE' key on the keypad
    CONTESTANT3KEYCODE = 17 'this is the 'CTRL' key on the keypad
    'others are:
'    Contestant1keycode = 96 'this is the '0' key on the keypad
'    Contestant1keycode = 100 'this is the '4' key on the keypad
'    Contestant2Keycode = 101 'this is the '5' key on the keypad
'    Contestant3Keycode = 102 'this is the '6' key on the keypad
'    Contestant1keycode = 103 'this is the '7' key on the keypad
'    Contestant2Keycode = 104 'this is the '8' key on the keypad
'    Contestant3Keycode = 105 'this is the '9' key on the keypad
End Sub

Private Sub getGameName(numGame, DJ)

Dim rs1 As ADODB.Recordset
'   Set cnn = DataEnvironment1.Connection1
Call dbConnection  'located in Module1

   Set cmd = New ADODB.Command
'   cnn.Open
   Set cmd.ActiveConnection = cnn
   cmd.CommandText = "select * from jeopardy where gamenumber = " + Str(numGame)
   cmd.CommandText = cmd.CommandText + " and DBLJ = " + Str(DJ)

   Set rs1 = cmd.Execute()
     mygameName = Trim(rs1!gameName)
     Text10.Text = mygameName
   rs1.Close
   cnn.Close
End Sub



Private Sub Frame9_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label15_Click(Index As Integer)
      'this sets new contestant to start game instead of #1 always
      If Index = 1 Then
       Shape4(1).Visible = True
       Shape4(2).Visible = False
       Shape4(3).Visible = False
      End If
      If Index = 2 Then
       Shape4(1).Visible = False
       Shape4(2).Visible = True
       Shape4(3).Visible = False
      End If
      If Index = 3 Then
       Shape4(1).Visible = False
       Shape4(2).Visible = False
       Shape4(3).Visible = True
      End If
      
       
End Sub

Private Sub Label16_Click()
   Call mySounds("backgroundmusic")
End Sub

Private Sub Label20_Click()
   theWager = Val(Trim(Text12.Text))
   If Len(Trim(Text12.Text)) = 0 Then
      MsgBox ("Please Enter Your Wager.")
   Else
      Image3.Visible = True
      Image6.Visible = True
      Frame5.Visible = True
      frame12.Visible = False
      player1guessed = False
      player2guessed = False
      player3guessed = False
      
   End If
End Sub

Private Sub Label29_Click()
'   Call chooseKey(Label29.Caption, X)
   Label37.Caption = "SHIFT_L"
   CONTESTANT1KEYCODE = 16
End Sub
Private Sub chooseKey(keyName, asciiNum)
  If Option1(1).Value = True Then
     Label32.Caption = keyName
     Label36.Caption = Str(asciiNum)
  End If
  If Option1(2).Value = True Then
      Label33.Caption = keyName
      Label37.Caption = Str(asciiNum)
  End If
  If Option1(3).Value = True Then
      Label34.Caption = keyName
      Label38.Caption = Str(asciiNum)
  End If
    
    
End Sub

Private Sub Label30_Click()
   Label35.Caption = "CTRL"
   CONTESTANT3KEYCODE = 17
End Sub

Private Sub Label31_Click()
   Label36.Caption = "SPACE"
   CONTESTANT2KEYCODE = 32
End Sub

Private Sub Label38_Click()
   Call chooseKey(Label38.Caption, X)
End Sub

Private Sub Label39_Click()
   Label37.Caption = 1
   CONTESTANT1KEYCODE = 97
   
End Sub

Private Sub Label40_Click()
   Label36.Caption = 2
   CONTESTANT2KEYCODE = 98
End Sub

Private Sub Label41_Click()
   Label35.Caption = 3
   CONTESTANT3KEYCODE = 99
End Sub

Private Sub mnuChangeScores_Click()
     Frame11.Visible = True
     Label19(1).Caption = Label1.Caption
     Label19(2).Caption = Label3.Caption
     Label19(3).Caption = Label6.Caption
     Text11(1).Text = Label10.Caption
     Text11(2).Text = Label11.Caption
     Text11(3).Text = Label12.Caption
     
     
End Sub

Private Sub mnuDJ_Click()
        DJ = 1
        For X = 31 To 36             'cover up categories
            Image1(X).Visible = True
        Next X
        Call getGameName(myGameNumber, DJ)

       'let the lead contestant go first...
        If Val(Trim(Label10.Caption)) > Val(Trim(Label11.Caption)) Then
                If Val(Trim(Label10.Caption)) > Val(Trim(Label12.Caption)) Then
                      'contestant 1 is highest
'                      Shape5.Left = 690  'move white dot to indicate who goes next
                       Shape4(1).Visible = True
                       Shape4(2).Visible = False
                       Shape4(3).Visible = False
                Else
                      'contestant 3 is highest
'                      Shape5.Left = 3975
                       Shape4(1).Visible = False
                       Shape4(2).Visible = False
                       Shape4(3).Visible = True
                End If
        Else    'label 11 is more than label 10
                If Val(Trim(Label11.Caption)) > Val(Trim(Label12.Caption)) Then
                    ' contestant 2 is highest
'                    Shape5.Left = 2340
                       Shape4(1).Visible = False
                       Shape4(2).Visible = True
                       Shape4(3).Visible = False
                Else
                    ' contestant 3 Is highest
'                    Shape5.Left = 3975
                       Shape4(1).Visible = False
                       Shape4(2).Visible = False
                       Shape4(3).Visible = True
                End If
        End If
       
'          Frame9.Visible = False 'Final Jeopardy Screen
          Label22.Visible = True     'the word double
          Label23.Visible = True     'superimposed two images
'      If bDoubleJeopardy = False Then
        For X = 1 To 30             'set values of each column entries to twice normal game
           Label4(X).Caption = Str(Val(Label4(X).Caption) * 2)
        Next X
           Label4(5).FontSize = 0.75 * Label4(1).FontSize  'last row of numbers is too big  (normally 26)
           Label4(10).FontSize = 0.75 * Label4(1).FontSize
           Label4(15).FontSize = 0.75 * Label4(1).FontSize
           Label4(20).FontSize = 0.75 * Label4(1).FontSize
           Label4(25).FontSize = 0.75 * Label4(1).FontSize
           Label4(30).FontSize = 0.75 * Label4(1).FontSize
'     End If
      bDoubleJeopardy = True
      Call Command4_Click
'      mnuDJ.Enabled = False
      mnuFinalJeopardy.Enabled = True
End Sub
Private Sub command14_KeyDown(KeyCode As Integer, Shift As Integer)   'use this to set up to capture who clicked in first
                                                  ' with ONE keyboard, or even three--one per player
  Timer4.Enabled = True
  Timer2.Enabled = False

 If KeyCode = CONTESTANT1KEYCODE Then        '  (this is for contestant #1 )
      If ok2click_1 = True Then
        If someoneClicked = False Then
              someoneClicked = True
              player1guessed = True
              Shape5.Visible = True
              Shape5.Left = 690
              Call ding

              Call Label1_Click
              'set it so contestant #1  cannot buzz in until next new question is revealed
              ok2click_1 = False
              Shape6(1).BackColor = vbRed
              
         End If
'         Command14.SetFocus        'sets focus so that next key press is not recognized in text9
       End If
   End If
   
   
'      If KeyCode = 32 Then       ' SPACEBar (32)  pressed  (this is for contestant # 2)
      If KeyCode = CONTESTANT2KEYCODE Then    'this is for contestant #2
'      Label39.Caption = Str(someoneClicked)
        If ok2click_2 = True Then
             If someoneClicked = False Then
              someoneClicked = True
               player2guessed = True
              Shape5.Visible = True
              Shape5.Left = 2340
                   Call ding
              Call Label3_Click
              'set it so contestant #2  cannot buzz in until new question is revealed
                   ok2click_2 = False
                   Shape6(2).BackColor = vbRed
'             Else
'                ok2click_2 = True
'                Shape6(2).BackColor = &H8000&
            End If
'         Command14.SetFocus
       End If
   End If
'     If KeyCode = 17 Then        ' "CTRL" KEY (17) pressed  (this is for contestant #3 )
      If KeyCode = CONTESTANT3KEYCODE Then
'     Label39.Caption = Str(someoneClicked)
        If ok2click_3 = True Then
            If someoneClicked = False Then
                  someoneClicked = True
                   player3guessed = True
                  Shape5.Visible = True
                  Shape5.Left = 3975
                       Call ding

                   Call Label6_Click
              'set it so contestant #3  cannot buzz in until new question is revealed
                   ok2click_3 = False
                   Shape6(3).BackColor = vbRed
'             Else
'                ok2click_3 = True
'                Shape6(3).BackColor = &H8000&
             End If
'         Command14.SetFocus
         End If
      End If
   
End Sub
Private Sub Image1_Click(Index As Integer)
  If Index < 31 Then
  ' do nothing---this is only for categories at top of screen--images 31 through 36
  Else
    Image1(31).Visible = False
    Call ding
    zz = 32
    Timer3.Enabled = True
End If
End Sub

Private Sub Image14_Click()
   If picNum1 < 7 Then
   picNum1 = picNum1 + 1
   Else
   picNum1 = 1
   End If
   Image14.Picture = LoadPicture(App.Path + myArrayOfPix(picNum1))
   Call whoosh
End Sub

Private Sub Image15_Click()
   
   If picNum2 < 7 Then
   picNum2 = picNum2 + 1
   Else
   picNum2 = 1
   End If
   Image15.Picture = LoadPicture(App.Path + myArrayOfPix(picNum2))
   Call whoosh
End Sub

Private Sub Image16_Click()
   
   If picNum3 < 7 Then
   picNum3 = picNum3 + 1
   Else
   picNum3 = 1
   End If
   Call whoosh
   Image16.Picture = LoadPicture(App.Path + myArrayOfPix(picNum3))
End Sub

Private Sub Image17_Click()
      Image17.Visible = False
      Call ding
      gameOver = True
End Sub

Private Sub Image3_Click()
'Label39.Caption = Str(someoneClicked)
If Label1.BackColor = vbWhite And Label3.BackColor = vbWhite And Label6.BackColor = vbWhite Then
  MsgBox ("Please Wait for a Contestant to buzz in.")
  Exit Sub
End If
   If dailyDouble = True Then
        valToAdd = 2 * theWager     'contestant gets TWICE what he/she bets
   Else
        valToAdd = Val(Label4(labelclicked).Caption)
   End If

   If Label1.BackColor = vbRed Then
       Contestant1NewValue = Val(Label10.Caption) + valToAdd
       Label10.Caption = Contestant1NewValue
       Shape4(1).Visible = True
       Shape4(2).Visible = False
       Shape4(3).Visible = False
   End If
   If Label3.BackColor = vbRed Then
       Contestant2NewValue = Val(Label11.Caption) + valToAdd
       Label11.Caption = Contestant2NewValue
       Shape4(1).Visible = False
       Shape4(2).Visible = True
       Shape4(3).Visible = False
   End If
   If Label6.BackColor = vbRed Then
       Contestant3NewValue = Val(Label12.Caption) + valToAdd
       Label12.Caption = Contestant3NewValue
       Shape4(1).Visible = False
       Shape4(2).Visible = False
       Shape4(3).Visible = True
   End If
' If numQleft = 0 Then
  If anyQs_Left = False Then
           Call endOfGameSound
   End If
   Frame5.Visible = False
   dailyDouble = False
      Label1.BackColor = vbWhite
   Label3.BackColor = vbWhite
   Label6.BackColor = vbWhite
   Label1.ForeColor = vbBlack
   Label3.ForeColor = vbBlack
   Label6.ForeColor = vbBlack
   Timer4.Enabled = False
   Timer6.Enabled = False
   Call correctwav
   someoneClicked = False
   Call Command11_Click
'   Frame14.Left = 500
'   Frame14.Width = 3000
'   Frame14.Height = 3000
   Frame14.Visible = False 'this is for the Pictures used in some questions
   For X = 1 To 3
      Shape6(X).BackColor = vbRed
   Next X
   Shape6(1).BackColor = vbRed
   ok2click_1 = False
   Shape6(2).BackColor = vbRed
   ok2click_2 = False
   Shape6(3).BackColor = vbRed
   ok2click_3 = False
   
End Sub

Private Sub Image6_Click()
If Label1.BackColor = vbWhite And Label3.BackColor = vbWhite And Label6.BackColor = vbWhite Then
  MsgBox ("Please Wait for a Contestant to buzz in.")
  Exit Sub
End If
   If dailyDouble = True Then
        valToSubtract = theWager    'contestant only subtracts value bet, not twice
        If contestant = 1 Then
            c1 = True 'variable to keep track if contestant 1 has already guessed and blew it
            ok2click_1 = False
            Shape6(1).BackColor = vbRed
            If c2 = False Then  'meaning contestant 2 has not yet guessed
               ok2click_2 = True
               Shape6(2).BackColor = vbGreen
            Else
               ok2click_2 = False
               Shape6(2).BackColor = vbRed
            End If
            If c3 = False Then   'meaning contesnt 3 has not yet guessed
               ok2click_3 = True
               Shape6(3).BackColor = vbGreen
            Else
               ok2click_3 = False
               Shape6(3).BackColor = vbRed
            End If
        End If
        If contestant = 2 Then
            c2 = True
            ok2click_2 = False
            Shape6(2).BackColor = vbRed
            If c1 = False Then  'meaning contestant 1 has not yet guessed
               ok2click_1 = True
               Shape6(1).BackColor = vbGreen
            Else
               ok2click_1 = False
               Shape6(1).BackColor = vbRed
            End If
            If c3 = False Then   'meaning contesnt 3 has not yet guessed
               ok2click_3 = True
               Shape6(3).BackColor = vbGreen
            Else
               ok2click_3 = False
               Shape6(3).BackColor = vbRed
            End If
        End If
        If contestant = 3 Then
            c3 = True
            ok2click_3 = False
            Shape6(3).BackColor = vbRed
            If c2 = False Then  'meaning contestant 2 has not yet guessed
               ok2click_2 = True
               Shape6(2).BackColor = vbGreen
            Else
               ok2click_2 = False
               Shape6(2).BackColor = vbRed
            End If
            If c1 = False Then   'meaning contestant 3 has not yet guessed
               ok2click_1 = True
               Shape6(1).BackColor = vbGreen
            Else
               ok2click_1 = False
               Shape6(1).BackColor = vbRed
            End If
        End If
        
   Else    'IF NOT DAILY DOUBLE THEN>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>.
        valToSubtract = Val(Label4(labelclicked).Caption)
   End If
   
   If Label1.BackColor = vbRed Then
       Contestant1NewValue = Val(Label10.Caption) - valToSubtract
       Label10.Caption = Contestant1NewValue
   End If
   If Label3.BackColor = vbRed Then
       Contestant2NewValue = Val(Label11.Caption) - valToSubtract
       Label11.Caption = Contestant2NewValue
   End If
   If Label6.BackColor = vbRed Then
       Contestant3NewValue = Val(Label12.Caption) - valToSubtract
       Label12.Caption = Contestant3NewValue
   End If
   If player1guessed = True Then
       Shape6(1).BackColor = vbRed
       ok2click_1 = False
   End If
      If player2guessed = True Then
       Shape6(2).BackColor = vbRed
       ok2click_2 = False
   End If
      If player3guessed = True Then
       Shape6(3).BackColor = vbRed
       ok2click_3 = False
   End If
   
   'gotta put here chances for the other one or two contestants to try guess if want to
   If player1guessed = False Or player2guessed = False Or player3guessed = False Then
        Call Command11_Click
        Timer4.Enabled = False
        Call wrongwav
        someoneClicked = False
        Exit Sub
   Else    'set all back to default
        Frame5.Visible = False
        dailyDouble = False
        Label1.BackColor = vbWhite
        Label3.BackColor = vbWhite
        Label6.BackColor = vbWhite
        Label1.ForeColor = vbBlack
        Label3.ForeColor = vbBlack
        Label6.ForeColor = vbBlack
        Timer4.Enabled = False
        Call wrongwav
        someoneClicked = False
        Call Command11_Click
        If anyQs_Left = False Then
           Call endOfGameSound
        End If
   End If
   Command14.SetFocus
   Timer6.Enabled = False
   Frame14.Visible = False 'this is for the Pictures used in some questions
End Sub

Private Sub Image8_Click()
'Label39.Caption = Str(someoneClicked)
   Frame5.Visible = False
   dailyDouble = False
   Label1.BackColor = vbWhite
   Label3.BackColor = vbWhite
   Label6.BackColor = vbWhite
   Label1.ForeColor = vbBlack
   Label3.ForeColor = vbBlack
   Label6.ForeColor = vbBlack
   Timer4.Enabled = False
   Timer6.Enabled = False
   Shape5.Visible = False
   Call exitwav
   someoneClicked = False
   Call Command11_Click
        If anyQs_Left = False Then
           Call endOfGameSound
        End If
   Frame14.Visible = False 'this is for the Pictures used in some questions
   ok2click_1 = False
   Shape6(1).BackColor = vbRed
   ok2click_1 = False
   Shape6(2).BackColor = vbRed
   ok2click_2 = False
   Shape6(3).BackColor = vbRed
   ok2click_3 = False
   
End Sub

Private Sub Label1_Click()
    Timer2.Enabled = False
    cPlayer = Trim(Label1.Caption)
    contestant = 1
    Label1.BackColor = vbRed
    Label1.ForeColor = vbWhite
    Label3.BackColor = vbWhite
    Label3.ForeColor = vbBlack
    Label6.BackColor = vbWhite
    Label6.ForeColor = vbBlack
    Call mySounds("endofround")
    Timer6.Enabled = True
    Shape6(1).BackColor = vbRed
End Sub
Private Sub Label3_Click()
    Timer2.Enabled = False
    cPlayer = Trim(Label3.Caption)
    contestant = 2
    Label1.BackColor = vbWhite
    Label1.ForeColor = vbBlack
    Label3.BackColor = vbRed
    Label3.ForeColor = vbWhite
    Label6.BackColor = vbWhite
    Label6.ForeColor = vbBlack
    Call mySounds("endofround")
    Timer6.Enabled = True
    Shape6(2).BackColor = vbRed
End Sub
Private Sub Label6_Click()
    Timer2.Enabled = False
    cPlayer = Trim(Label6.Caption)
    contestant = 3
    Label1.BackColor = vbWhite
    Label1.ForeColor = vbBlack
    Label3.BackColor = vbWhite
    Label3.ForeColor = vbBlack
    Label6.BackColor = vbRed
    Label6.ForeColor = vbWhite
    Call mySounds("endofround")
    Timer6.Enabled = True
    Shape6(3).BackColor = vbRed
End Sub
Private Sub mySounds(sndName)
     soundname = App.Path + "/sounds/" + sndName + ".wav"
     gbResults = PlaySound(soundname, 0, SND_ASYNC)
End Sub

Private Sub ding()
     soundname = App.Path + "/sounds/ding.wav"
     gbResults = PlaySound(soundname, 0, SND_ASYNC)
End Sub
Private Sub endOfGameSound()
     soundname = App.Path + "/sounds/endofround.wav"
     gbResults = PlaySound(soundname, 0, SND_ASYNC)
End Sub
Private Sub whoosh()
     soundname = App.Path + "/sounds/whoosh.wav"
     gbResults = PlaySound(soundname, 0, SND_ASYNC)
End Sub
Private Sub smclick()
     soundname = App.Path + "/sounds/NIMClick.wav"
     gbResults = PlaySound(soundname, 0, SND_ASYNC)
End Sub
Private Sub wrongwav()
     soundname = App.Path + "/sounds/error.wav"
     gbResults = PlaySound(soundname, 0, SND_ASYNC)
End Sub
Private Sub exitwav()
     soundname = App.Path + "/sounds/exit.wav"
     gbResults = PlaySound(soundname, 0, SND_ASYNC)
End Sub
Private Sub correctwav()
     soundname = App.Path + "/sounds/entry.wav"
     gbResults = PlaySound(soundname, 0, SND_ASYNC)
End Sub

Private Sub Label14_Click()
'MsgBox theLabelClicked
   theWager = Val(Trim(Text5.Text))
   If Len(Trim(Text5.Text)) = 0 Then
      MsgBox ("Please Enter Your Wager.")
   Else
      c1 = False
      c2 = False
      c3 = False
      Image3.Visible = True
      Image6.Visible = True
      Frame5.Visible = True
      Frame8.Visible = False
      player1guessed = False
      player2guessed = False
      player3guessed = False
      If gamenum = 4 And theLabelClicked = 21 And DJ = 1 Then   'show pictures in new frame to go with question
           Call loadPix("piercebrosnan.jpg")
        End If
        If gamenum = 4 And theLabelClicked = 22 And DJ = 1 Then   'show pictures in new frame to go with question
           Call loadPix("maureenosullivan.jpg")
        End If
        If gamenum = 4 And theLabelClicked = 23 And DJ = 1 Then   'show pictures in new frame to go with question
           Call loadPix("peterotoole.jpg")
        End If
        If gamenum = 4 And theLabelClicked = 24 And DJ = 1 Then   'show pictures in new frame to go with question
           Call loadPix("maureenohara.jpg")
        End If
        If gamenum = 4 And theLabelClicked = 25 And DJ = 1 Then   'show pictures in new frame to go with question
           Call loadPix("fionnulaflanagan.jpg")
        End If
        
        If gamenum = 6 And theLabelClicked = 1 And DJ = 0 Then  'show pictures in new frame to go with question
           Call loadPix("daytonaspeedway.jpg")
        End If
        If gamenum = 6 And theLabelClicked = 2 And DJ = 0 Then  'show pictures in new frame to go with question
           Call loadPix("indianapolisspeedway.jpg")
        End If
        If gamenum = 6 And theLabelClicked = 3 And DJ = 0 Then   'show pictures in new frame to go with question
           Call loadPix("jeffgordon-24.jpg")
        End If
        If gamenum = 6 And theLabelClicked = 4 And DJ = 0 Then   'show pictures in new frame to go with question
           Call loadPix("atlantaraceway.jpg")
        End If
        If gamenum = 6 And theLabelClicked = 5 And DJ = 0 Then  'show pictures in new frame to go with question
           Call loadPix("kurtbusch.jpg")
        End If
        
        If gamenum = 6 And theLabelClicked = 11 And DJ = 1 Then   'show pictures in new frame to go with question
           Call loadPix("monalisa-vinci.jpg")
        End If
        If gamenum = 6 And theLabelClicked = 12 And DJ = 1 Then   'show pictures in new frame to go with question
           Call loadPix("peasantman-gogh.jpg")
        End If
        If gamenum = 6 And theLabelClicked = 13 And DJ = 1 Then   'show pictures in new frame to go with question
           Call loadPix("paulconversion-mich.jpg")
        End If
        If gamenum = 6 And theLabelClicked = 14 And DJ = 1 Then   'show pictures in new frame to go with question
           Call loadPix("thepeasant-cezanne.jpg")
        End If
        If gamenum = 6 And theLabelClicked = 15 And DJ = 1 Then   'show pictures in new frame to go with question
           Call loadPix("paulgauguin.jpg")
        End If
        
        If gamenum = 7 And theLabelClicked = 16 And DJ = 0 Then   'show pictures in new frame to go with question
           Call mySounds("bingcrosby-HappyHoliday")
        End If
        If gamenum = 7 And theLabelClicked = 17 And DJ = 0 Then   'show pictures in new frame to go with question
           Call mySounds("natkingcole-ChristmasSong")
        End If
        If gamenum = 7 And theLabelClicked = 18 And DJ = 0 Then   'show pictures in new frame to go with question
           Call mySounds("FrankSinatra-JingleBells")
        End If
        If gamenum = 7 And theLabelClicked = 19 And DJ = 0 Then   'show pictures in new frame to go with question
           Call mySounds("JimmyBuffet-ChristmasIsland")
        End If
        If gamenum = 7 And theLabelClicked = 20 And DJ = 0 Then   'show pictures in new frame to go with question
           Call mySounds("AnnMurray-JoyToTheWorld")
        End If
        
       If gamenum = 10 And theLabelClicked = 21 And DJ = 0 Then   'show pictures in new frame to go with question
           Call loadPix("dr_bencasey.jpg")
        End If
        If gamenum = 10 And theLabelClicked = 22 And DJ = 0 Then   'show pictures in new frame to go with question
           Call loadPix("doogiehowsermd.jpg")
        End If
        If gamenum = 10 And theLabelClicked = 23 And DJ = 0 Then    'show pictures in new frame to go with question
           Call loadPix("dr_er.jpg")
        End If
        If gamenum = 10 And theLabelClicked = 24 And DJ = 0 Then   'show pictures in new frame to go with question
           Call loadPix("dr_house.jpg")
        End If
        If gamenum = 10 And theLabelClicked = 25 And DJ = 0 Then    'show pictures in new frame to go with question
           Call loadPix("dr_greys anatomy.jpg")
        End If
        
        
        If gamenum = 11 And theLabelClicked = 1 And DJ = 1 Then   'show pictures in new frame to go with question
           Call loadPix("alaska_outline.jpg")
        End If
        If gamenum = 11 And theLabelClicked = 2 And DJ = 1 Then   'show pictures in new frame to go with question
           Call loadPix("vermont_outline.jpg")
        End If
        If gamenum = 11 And theLabelClicked = 3 And DJ = 1 Then    'show pictures in new frame to go with question
           Call loadPix("kentucky_outline.jpg")
        End If
        If gamenum = 11 And theLabelClicked = 4 And DJ = 1 Then   'show pictures in new frame to go with question
           Call loadPix("mississippi_outline.jpg")
        End If
        If gamenum = 11 And theLabelClicked = 5 And DJ = 1 Then    'show pictures in new frame to go with question
           Call loadPix("iowa_outline.jpg")
        End If
        
        If gamenum = 10 And theLabelClicked = 26 And DJ = 1 Then 'show pictures in new frame to go with question
           Call mySounds("woodywoodpecker")
        End If
        If gamenum = 10 And theLabelClicked = 27 And DJ = 1 Then   'show pictures in new frame to go with question
           Call mySounds("porkypig")
        End If
        If gamenum = 10 And theLabelClicked = 28 And DJ = 1 Then    'show pictures in new frame to go with question
           Call mySounds("bullwinkle")
        End If
        If gamenum = 10 And theLabelClicked = 29 And DJ = 1 Then   'show pictures in new frame to go with question
           Call mySounds("georgeofthejungle")
        End If
        If gamenum = 10 And theLabelClicked = 30 And DJ = 1 Then    'show pictures in new frame to go with question
           Call mySounds("superchicken")
        End If
        
'        End If
            ok2click_1 = False
            ok2click_2 = False
            ok2click_3 = False
            Shape6(1).BackColor = vbRed
            Shape6(2).BackColor = vbRed
            Shape6(3).BackColor = vbRed
'        If player1guessed = True Then
'            ok2click_1 = False
'            ok2click_2 = True
'            ok2click_3 = True
'            Shape6(1).BackColor = vbRed
'            Shape6(2).BackColor = &H8000&
'            Shape6(3).BackColor = &H8000&
'        End If
'        If player2guessed = True Then
'            ok2click_1 = True
'            ok2click_2 = False
'            ok2click_3 = True
'            Shape6(1).BackColor = &H8000&
'            Shape6(2).BackColor = vbRed
'            Shape6(3).BackColor = &H8000&
'        End If
'        If player3guessed = True Then
'            ok2click_1 = True
'            ok2click_2 = True
'            ok2click_3 = False
'            Shape6(1).BackColor = &H8000&
'            Shape6(2).BackColor = &H8000&
'            Shape6(3).BackColor = vbRed
'        End If
   End If
End Sub



Private Sub Label2_Click(Index As Integer)
Timer1.Enabled = False
End Sub



Private Sub Label4_Click(Index As Integer)
   theLabelClicked = Index
   Label4(Index).Visible = False
   anyQs_Left = False
   For X = 1 To 30
      If Label4(X).Visible = True Then
          anyQs_Left = True
      End If
   Next X
   numQleft = numQleft - 1
   Call ding
    If bDoubleJeopardy = False Then
      myRandomNum = myRandomNum    'myRandomNum  is between 1 and 30
'      menuMac.Caption = Str(myRandomNum)
    Else
       If myRandomNum < 8 Then    ' in order for myRandomNumDJ to be greater than 0, am doing this
           myRandomNumDJ2 = myRandomNum + myRandomNumDJ  'mRandomNumDJ is between 1 and 7
        Else
           myRandomNumDJ2 = myRandomNum - myRandomNumDJ
        End If
'        menuMac.Caption = Str(myRandomNumDJ2)
    End If
    'stop players from clicking in until alex reads first 1 second of the question
    'this is also done when someone gets the right answer and Alex clicks Correct (Image3) AND
    'this is also done when no one gets the right answer and Alex clicks Jeopardy button (Image8)
ok2click_1 = False
Shape6(1).BackColor = vbRed
ok2click_2 = False
Shape6(2).BackColor = vbRed
ok2click_3 = False
Shape6(3).BackColor = vbRed
    'set it so players can now buzz in:
    Timer7.Enabled = True   'this function is done in timer7--1 second pause after frame5 is shown
 
If bDoubleJeopardy = False Then
'    myRandomNum = 24 'this is temporary
    If Index = myRandomNum Then  'do daily double
           If numQleft = 29 Then  'if this is the very first one that contestant actually chose dailydbl first thing
          Call popArray
       End If
        ' do the daily double thing!!!  here
        dailyDouble = True
        Label5.Caption = myArrayOfAnswers(myRandomNum)
        Text5.Text = ""
        Frame8.Visible = True
        soundname = App.Path + "/sounds/dailydbl.wav"
        gbResults = PlaySound(soundname, 0, SND_ASYNC)
         For X = 1 To 3
           If Shape4(X).Visible = True Then
              contestantNum = X
           End If
        Next X
        If contestantNum = 1 Then
           Label1.BackColor = vbRed
        End If
        If contestantNum = 2 Then
            Label3.BackColor = vbRed
        End If
        If contestantNum = 3 Then
            Label6.BackColor = vbRed
        End If
    Else     'show the regular question, NOT daily double
        valOfLabel = Val(Label4(Index).Caption)
        Frame5.Left = 2000
        Frame5.Top = 1600
        z = 0
        Image3.Visible = True
        Image6.Visible = True
        
        Frame5.Visible = True
        Timer2.Enabled = True  'give the contestants 15 seconds to answer
        player1guessed = False
        player2guessed = False
        player3guessed = False
        Label5.Caption = myArrayOfAnswers(Index)
        Label4(Index).Visible = False
        labelclicked = Index
     End If
    'pictures
    If dailyDouble = False Then  'can't have these pop up until dailydouble is clicked....
        If gamenum = 6 And Index = 1 Then   'show pictures in new frame to go with question
           Call loadPix("daytonaspeedway.jpg")
        End If
        If gamenum = 6 And Index = 2 Then   'show pictures in new frame to go with question
           Call loadPix("indianapolisspeedway.jpg")
        End If
        If gamenum = 6 And Index = 3 Then   'show pictures in new frame to go with question
           Call loadPix("jeffgordon-24.jpg")
        End If
        If gamenum = 6 And Index = 4 Then   'show pictures in new frame to go with question
           Call loadPix("atlantaraceway.jpg")
        End If
        If gamenum = 6 And Index = 5 Then   'show pictures in new frame to go with question
           Call loadPix("kurtbusch.jpg")
        End If
        If gamenum = 7 And theLabelClicked = 16 Then    'show pictures in new frame to go with question
           Call mySounds("bingcrosby-HappyHoliday")
        End If
        If gamenum = 7 And theLabelClicked = 17 Then    'show pictures in new frame to go with question
           Call mySounds("natkingcole-ChristmasSong")
        End If
        If gamenum = 7 And theLabelClicked = 18 Then    'show pictures in new frame to go with question
           Call mySounds("FrankSinatra-JingleBells")
        End If
        If gamenum = 7 And theLabelClicked = 19 Then    'show pictures in new frame to go with question
           Call mySounds("JimmyBuffet-ChristmasIsland")
        End If
        If gamenum = 7 And theLabelClicked = 20 Then    'show pictures in new frame to go with question
           Call mySounds("AnnMurray-JoyToTheWorld")
        End If
        If gamenum = 10 And theLabelClicked = 21 Then   'show pictures in new frame to go with question
           Call loadPix("dr_bencasey.jpg")
        End If
        If gamenum = 10 And theLabelClicked = 22 Then    'show pictures in new frame to go with question
           Call loadPix("doogiehowsermd.jpg")
        End If
        If gamenum = 10 And theLabelClicked = 23 Then     'show pictures in new frame to go with question
           Call loadPix("dr_er.jpg")
        End If
        If gamenum = 10 And theLabelClicked = 24 Then   'show pictures in new frame to go with question
           Call loadPix("dr_house.jpg")
        End If
        If gamenum = 10 And theLabelClicked = 25 Then    'show pictures in new frame to go with question
           Call loadPix("dr_greys anatomy.jpg")
        End If
 
    End If
Else   'double Jeopardy is TRUE

    If Index = myRandomNum Or Index = myRandomNumDJ2 Then
       If numQleft = 29 Then
          Call popArray
          If Index = myRandomNum Then
               Label5.Caption = myArrayOfAnswers(myRandomNum)
          End If
          If Index = myRandomNumDJ2 Then
               Label5.Caption = myArrayOfAnswers(myRandomNumDJ2)
'               MsgBox Str(myRandomNumDJ2) + "  - " + myArrayOfAnswers(myRandomNumDJ2)
          End If
       End If
        ' do the daily double thing!!!  here
        dailyDouble = True
        Label5.Caption = myArrayOfAnswers(Index)
        Text5.Text = ""
        Frame8.Visible = True
        soundname = App.Path + "/sounds/dailydbl.wav"
        gbResults = PlaySound(soundname, 0, SND_ASYNC)
         For X = 1 To 3
           If Shape4(X).Visible = True Then
              contestantNum = X
           End If
        Next X
        If contestantNum = 1 Then
           Label1.BackColor = vbRed
           Shape6(1).BackColor = &H8000&  'green
           Shape6(2).BackColor = vbRed
           Shape6(3).BackColor = vbRed
           ok2click_2 = False
           ok2click_3 = False
        End If
        If contestantNum = 2 Then
            Label3.BackColor = vbRed
           Shape6(1).BackColor = vbRed
           Shape6(2).BackColor = &H8000&  'green
           Shape6(3).BackColor = vbRed
           ok2click_1 = False
           ok2click_3 = False
        End If
        If contestantNum = 3 Then
            Label6.BackColor = vbRed
           Shape6(1).BackColor = vbRed
           Shape6(2).BackColor = vbRed
           Shape6(3).BackColor = &H8000&
           ok2click_1 = False
           ok2click_2 = False
        End If
    Else
        valOfLabel = Val(Label4(Index).Caption)
        Frame5.Left = 2000
        Frame5.Top = 1600
        z = 0
        Image3.Visible = True
        Image6.Visible = True
        Frame5.Visible = True
        Timer2.Enabled = True  'give the contestants 15 seconds to answer
        player1guessed = False
        player2guessed = False
        player3guessed = False
 
        Label5.Caption = myArrayOfAnswers(Index)
        Label4(Index).Visible = False
        labelclicked = Index

    End If
    If dailyDouble = False Then  'can't have these pop up until dailydouble is clicked....
        If gamenum = 4 And Index = 21 Then   'show pictures in new frame to go with question
           Call loadPix("piercebrosnan.jpg")
        End If
        If gamenum = 4 And Index = 22 Then   'show pictures in new frame to go with question
           Call loadPix("maureenosullivan.jpg")
        End If
        If gamenum = 4 And Index = 23 Then   'show pictures in new frame to go with question
           Call loadPix("peterotoole.jpg")
        End If
        If gamenum = 4 And Index = 24 Then   'show pictures in new frame to go with question
           Call loadPix("maureenohara.jpg")
        End If
        If gamenum = 4 And Index = 25 Then   'show pictures in new frame to go with question
           Call loadPix("fionnulaflanagan.jpg")
        End If
        If gamenum = 6 And theLabelClicked = 11 Then   'show pictures in new frame to go with question
           Call loadPix("monalisa-vinci.jpg")
        End If
        If gamenum = 6 And theLabelClicked = 12 Then   'show pictures in new frame to go with question
           Call loadPix("peasantman-gogh.jpg")
        End If
        If gamenum = 6 And theLabelClicked = 13 Then   'show pictures in new frame to go with question
           Call loadPix("paulconversion-mich.jpg")
        End If
        If gamenum = 6 And theLabelClicked = 14 Then   'show pictures in new frame to go with question
           Call loadPix("thepeasant-cezanne.jpg")
        End If
        If gamenum = 6 And theLabelClicked = 15 Then   'show pictures in new frame to go with question
           Call loadPix("paulgauguin.jpg")
        End If
        If gamenum = 10 And theLabelClicked = 21 Then    'show pictures in new frame to go with question
           Call loadPix("dr_bencasey.jpg")
        End If
        If gamenum = 10 And theLabelClicked = 22 Then    'show pictures in new frame to go with question
           Call loadPix("doogiehowsermd.jpg")
        End If
        If gamenum = 10 And theLabelClicked = 23 Then     'show pictures in new frame to go with question
           Call loadPix("dr_er.jpg")
        End If
        If gamenum = 10 And theLabelClicked = 24 Then    'show pictures in new frame to go with question
           Call loadPix("dr_house.jpg")
        End If
        If gamenum = 10 And theLabelClicked = 25 Then     'show pictures in new frame to go with question
           Call loadPix("dr_greys anatomy.jpg")
        End If
        If gamenum = 11 And theLabelClicked = 1 Then    'show pictures in new frame to go with question
           Call loadPix("alaska_outline.jpg")
        End If
        If gamenum = 11 And theLabelClicked = 2 Then     'show pictures in new frame to go with question
           Call loadPix("vermont_outline.jpg")
        End If
        If gamenum = 11 And theLabelClicked = 3 Then     'show pictures in new frame to go with question
           Call loadPix("kentucky_outline.jpg")
        End If
        If gamenum = 11 And theLabelClicked = 4 Then     'show pictures in new frame to go with question
           Call loadPix("mississippi_outline.jpg")
        End If
        If gamenum = 11 And theLabelClicked = 5 Then     'show pictures in new frame to go with question
           Call loadPix("iowa_outline.jpg")
        End If
        
        If gamenum = 10 And theLabelClicked = 26 Then  'show pictures in new frame to go with question
           Call mySounds("woodywoodpecker")
        End If
        If gamenum = 10 And theLabelClicked = 27 Then    'show pictures in new frame to go with question
           Call mySounds("porkypig")
        End If
        If gamenum = 10 And theLabelClicked = 28 Then     'show pictures in new frame to go with question
           Call mySounds("bullwinkle")
        End If
        If gamenum = 10 And theLabelClicked = 29 Then    'show pictures in new frame to go with question
           Call mySounds("georgeofthejungle")
        End If
        If gamenum = 10 And theLabelClicked = 30 Then     'show pictures in new frame to go with question
           Call mySounds("superchicken")
        End If
        
    End If
    
End If
Command14.SetFocus
End Sub
Private Sub loadPix(picName)
     Frame14.Visible = True
     Frame14.Top = 135
     Frame14.Left = 8000
     Image21.Picture = LoadPicture(App.Path + "\images\" + picName)
End Sub
Private Sub popArray()
      Dim rs As ADODB.Recordset
'   Set cnn = DataEnvironment1.Connection1
   Call dbConnection  'located in Module1
   Set cmd = New ADODB.Command
'   cnn.Open
   Set cmd.ActiveConnection = cnn
   cmd.CommandText = "select * from jeopardy "
   cmd.CommandText = cmd.CommandText + "where gameNumber = " & Str(gamenum) & " and DBLJ = " & Str(DJ)
   Set rs = cmd.Execute()
       myArrayOfAnswers(1) = Trim(rs!ans1)
       myArrayOfAnswers(2) = Trim(rs!ans2)
       myArrayOfAnswers(3) = Trim(rs!ans3)
       myArrayOfAnswers(4) = Trim(rs!ans4)
       myArrayOfAnswers(5) = Trim(rs!ans5)
       myArrayOfAnswers(6) = Trim(rs!ans6)
       myArrayOfAnswers(7) = Trim(rs!ans7)
       myArrayOfAnswers(8) = Trim(rs!ans8)
       myArrayOfAnswers(9) = Trim(rs!ans9)
       myArrayOfAnswers(10) = Trim(rs!ans10)
       myArrayOfAnswers(11) = Trim(rs!ans11)
       myArrayOfAnswers(12) = Trim(rs!ans12)
       myArrayOfAnswers(13) = Trim(rs!ans13)
       myArrayOfAnswers(14) = Trim(rs!ans14)
       myArrayOfAnswers(15) = Trim(rs!ans15)
       myArrayOfAnswers(16) = Trim(rs!ans16)
       myArrayOfAnswers(17) = Trim(rs!ans17)
       myArrayOfAnswers(18) = Trim(rs!ans18)
       myArrayOfAnswers(19) = Trim(rs!ans19)
       myArrayOfAnswers(20) = Trim(rs!ans20)
       myArrayOfAnswers(21) = Trim(rs!ans21)
       myArrayOfAnswers(22) = Trim(rs!ans22)
       myArrayOfAnswers(23) = Trim(rs!ans23)
       myArrayOfAnswers(24) = Trim(rs!ans24)
       myArrayOfAnswers(25) = Trim(rs!ans25)
       myArrayOfAnswers(26) = Trim(rs!ans26)
       myArrayOfAnswers(27) = Trim(rs!ans27)
       myArrayOfAnswers(28) = Trim(rs!ans28)
       myArrayOfAnswers(29) = Trim(rs!ans29)
       myArrayOfAnswers(30) = Trim(rs!ans30)
        If DJ = 0 Then  'use regular jeopardy finalquestion
           finalCategory = Trim(rs!finalcat)
           finalQuestion = Trim(rs!finaljeop)
       Else
           finalCategory = finalCategory
           finalQuestion = finalQuestion   'this was found when DJ was 0 (Regular Jeopardy!)
       End If
   rs.Close
   cnn.Close
End Sub



Private Sub mnuAssignNames_Click()
     Call outoftime
     picNum1 = 1
     picNum2 = 2
     picNum3 = 3
     Call setPix
     Image14.Picture = LoadPicture(App.Path + myArrayOfPix(picNum1))
     Image15.Picture = LoadPicture(App.Path + myArrayOfPix(picNum2))
     Image16.Picture = LoadPicture(App.Path + myArrayOfPix(picNum3))
     Frame7.Visible = True
     Image14.Visible = True
     Image15.Visible = True
     Image16.Visible = True
     
     
End Sub
Private Sub outoftime()
     soundname = App.Path + "/sounds/outoftime.wav"
'     soundname = App.Path + "/sounds/whoosh.wav"
     gbResults = PlaySound(soundname, 0, SND_ASYNC)
End Sub
Private Sub setPix()
   myArrayOfPix(1) = "/images/man1.jpg"
   myArrayOfPix(2) = "/images/man3.jpg"
   myArrayOfPix(3) = "/images/woman1.jpg"
   myArrayOfPix(4) = "/images/man2.jpg"
   myArrayOfPix(5) = "/images/man4.jpg"
   myArrayOfPix(6) = "/images/woman2.jpg"
   myArrayOfPix(7) = "/images/woman3.jpg"
End Sub
Private Sub mnuExit_Click()
    quitme = MsgBox("Are you sure you want to exit?", vbYesNoCancel, "MacJeopardy")
    If quitme = 6 Then
       Unload Me
    End If
End Sub

Private Sub mnuFinalJeopardy_Click()
'     gamenum = 1   'temporary
     If gamenum = 0 Then
         MsgBox ("You Must Play A GAME First")
     Else
         
         
'         MsgBox gamenum
         For X = 1 To 3
             Shape6(X).BackColor = vbRed
         Next X
'         ok2click_1 = False
'         ok2click_2 = False
'         ok2click_3 = False
'
         Call popArray
'         Frame9.Left = 360
'         Frame9.Top = 15
'         Frame9.Visible = True
'         Image17.Visible = True
'         Label15.Caption = finalCategory
'         Label16.Caption = Label1.Caption
'         Label17.Caption = Label3.Caption
'         Label18.Caption = Label6.Caption
'         For X = 1 To 3
'             Shape4(X).Visible = True
'         Next X
'         mnuDJ.Enabled = False
         mnuFinalJeopardy.Enabled = False
         mnuSelectGame.Enabled = True
'         bDoubleJeopardy = False
         frmFinalJ.Show vbModal
     End If
End Sub

Private Sub mnuHelpAbout_Click()
  MsgBox ("              MacJeopardy      " + vbNewLine + "               Version 1.0 " _
          + vbNewLine + _
          "    All rights Reserved - November 2004")
End Sub
Function Wait(numseconds As Long)  'this function is used to wait a certain number of seconds
       ' I have not tested this yet, but got it from the internet.
    Dim start As Variant, rightnow As Variant
    Dim HourDiff As Variant, MinuteDiff As Variant, SecondDiff As Variant
    Dim TotalMinDiff As Variant, TotalSecDiff As Variant
    start = Now
    While True
        rightnow = Now
        HourDiff = Hour(rightnow) - Hour(start)
        MinuteDiff = Minute(rightnow) - Minute(start)
        SecondDiff = Second(rightnow) - Second(start) + 1
        If SecondDiff = 60 Then
            MinuteDiff = MinuteDiff + 1 ' Add 1 To minute.
            SecondDiff = 0 ' Zero seconds.
        End If
        If MinuteDiff = 60 Then
            HourDiff = HourDiff + 1 ' Add 1 To hour.
            MinuteDiff = 0 ' Zero minutes.
        End If
        TotalMinDiff = (HourDiff * 60) + MinuteDiff ' Get totals.
        TotalSecDiff = (TotalMinDiff * 60) + SecondDiff
        If TotalSecDiff >= numseconds Then
            Exit Function
        End If
        DoEvents
            'Debug.Print rightnow
        Wend
    End Function


Private Sub mnuHelpGen_Click()
      frmHelp.Show
End Sub

Private Sub mnuPlayers_Click()

End Sub

Private Sub mnuInputKeys_Click()
     frame12.Visible = True
     frame12.Left = 50
     frame12.Top = 50
     frame12.Width = frmMain.Width
     frame12.Height = frmMain.Height
End Sub

Private Sub mnuReset_Click()
   Call Command11_Click
End Sub

Private Sub mnuSelectGame_Click()
        For X = 31 To 36             'cover up categories
            Image1(X).Visible = True
        Next X
    Frame6.Visible = True
'    Frame9.Visible = False 'Double Jeopardy Screen
    Image4.Visible = True
     Label22.Visible = False     'the word double
     Label23.Visible = False     'superimposed two images
    If bDoubleJeopardy = True Then   'reset some things that are diff in dblejeop
''        DJ = 1
        For X = 1 To 30   'need to set values of each label (are doubled in DoubleJeopardy)
           Label4(X).Visible = True   'reset all if returning from another game
           Label4(X).Caption = Str(Val(Label4(X).Caption) / 2)
        Next X
        Label4(5).FontSize = Label4(1).FontSize    'was set to 3/4 size in DoubleJeopardy so they would fit
        Label4(10).FontSize = Label4(1).FontSize   ' reset here back to normal
        Label4(15).FontSize = Label4(1).FontSize
        Label4(20).FontSize = Label4(1).FontSize
        Label4(25).FontSize = Label4(1).FontSize
        Label4(30).FontSize = Label4(1).FontSize
     Else
'        DJ = 0
     End If
     bDoubleJeopardy = False
     DJ = 0
     Image14.Visible = True    'Contestants Images
     Image15.Visible = True
     Image16.Visible = True
     Label10.Caption = "0"
     Label11.Caption = "0"
     Label12.Caption = "0"
     mnuSelectGame.Enabled = False
     mnuDJ.Enabled = True   'have to play regular Jeopardy first
End Sub

Private Sub mtest_Click()
    Form1.Show
    
End Sub

Private Sub Timer1_Timer()    'this is used to move the white rectangle around the board at startup
                            'this timer moves it to the right  (timer5 is used to move it back to the left)
       Shape3.Visible = True
       Shape3.Left = Label4(xx).Left
       Shape3.Top = Label4(xx).Top
       If xx < 28 Then
        xx = xx + 2
       Else
         Shape3.Visible = False
         Timer1.Enabled = False
         yy = 30
         Timer5.Enabled = True
       End If
End Sub

Private Sub Timer2_Timer()  'gives players 10 seconds to answer
     z = z + 1
     If z = 10 Then      'this equates to 10 seconds
       Call outoftime
     ok2click_1 = False
     Shape6(1).BackColor = vbRed
     ok2click_2 = False
     Shape6(2).BackColor = vbRed
     ok2click_3 = False
     Shape6(3).BackColor = vbRed
       Image3.Visible = False
       Image6.Visible = False
       Image8.Visible = True
       Timer2.Enabled = False
      End If
End Sub

Private Sub Timer3_Timer()   'timer3 is used to show the categories at the top of the screen.
If zz < 37 Then
    Image1(zz).Visible = False
    Call ding
    zz = zz + 1
Else
    Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()    ' contestant took too much time to answer after clicking in
                              ' gets 20 seconds   (20000 in timer4 interval)
    Call outoftime  'out of time sound
    If dailyDouble = True Then     'set how much to subtract from the contestant who clicked in
        valToSubtract = theWager
   Else
        valToSubtract = valOfLabel
   End If
   If contestant = 1 Then
       Contestant1NewValue = Val(Label10.Caption) - valToSubtract
       Label10.Caption = Contestant1NewValue
       player1guessed = True
       ok2click_1 = False
       Shape6(1).BackColor = vbRed
   End If
   If contestant = 2 Then
       Contestant2NewValue = Val(Label11.Caption) - valToSubtract
       Label11.Caption = Contestant2NewValue
       player2guessed = True
       ok2click_2 = False
       Shape6(2).BackColor = vbRed
   End If
   If contestant = 3 Then
       Contestant3NewValue = Val(Label12.Caption) - valToSubtract
       Label12.Caption = Contestant3NewValue
       player3guessed = True
       ok2click_3 = False
       Shape6(3).BackColor = vbRed
   End If
    Timer4.Enabled = False
Call Image6_Click
'    Call Image6_Click
'    Call Command11_Click
End Sub

Private Sub Timer5_Timer()    'this is used to move the white rectangle around the board at startup
                            'this timer moves it to the left  (timer1 is used to move it to the right)
       Shape3.Visible = True
       Shape3.Left = Label4(yy).Left
       Shape3.Top = Label4(yy).Top
       If yy > 3 Then
        yy = yy - 2
       Else
         Shape3.Visible = False
         Timer5.Enabled = False
       End If
End Sub



Private Sub Timer6_Timer()
myz = myz + 1
If myz = 20 Then
  Call ding
  myz = 0
  Timer6.Enabled = False
End If
End Sub

Private Sub Timer7_Timer()  'timer used to delay contestants from clicking in too early
     For X = 1 To 3
        Shape6(X).BackColor = vbGreen
     Next X
        ok2click_1 = True
        ok2click_2 = True
        ok2click_3 = True
     Timer7.Enabled = False
End Sub


