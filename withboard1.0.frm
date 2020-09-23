VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command17 
      Caption         =   "RM"
      Height          =   375
      Left            =   11880
      TabIndex        =   41
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      Caption         =   "BM"
      Height          =   375
      Left            =   11880
      TabIndex        =   40
      Top             =   5400
      Width           =   495
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Km"
      Height          =   255
      Left            =   11880
      TabIndex        =   39
      Top             =   6480
      Width           =   495
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Qm"
      Height          =   255
      Left            =   11880
      TabIndex        =   38
      Top             =   6840
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      Caption         =   "NM"
      Height          =   255
      Left            =   11880
      TabIndex        =   37
      Top             =   7200
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   5535
      Left            =   5640
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   36
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   600
      TabIndex        =   42
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   4770
      TabIndex        =   34
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   4170
      TabIndex        =   33
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   3570
      TabIndex        =   32
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label Label23 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   1770
      TabIndex        =   31
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label Label22 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   2370
      TabIndex        =   30
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label Label21 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   2970
      TabIndex        =   29
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label f 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   1170
      TabIndex        =   28
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label L 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   540
      TabIndex        =   27
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "  8 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   5160
      TabIndex        =   25
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "  7 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   5160
      TabIndex        =   24
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "  6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   5160
      TabIndex        =   23
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "  5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   5160
      TabIndex        =   22
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "  4 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   5160
      TabIndex        =   21
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "  3 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   5160
      TabIndex        =   20
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "  2 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   5160
      TabIndex        =   19
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "  1 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   5160
      TabIndex        =   18
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "  8 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "  7 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "  6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "  5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "  4 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "  3 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "  2 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "  1 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   5160
      Width           =   375
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   0
      Left            =   360
      Picture         =   "withboard1.0.frx":0000
      Stretch         =   -1  'True
      Tag             =   "0"
      Top             =   840
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   1
      Left            =   960
      Picture         =   "withboard1.0.frx":13F6
      Stretch         =   -1  'True
      Tag             =   "1"
      Top             =   840
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   2
      Left            =   1560
      Picture         =   "withboard1.0.frx":2830
      Stretch         =   -1  'True
      Tag             =   "2"
      Top             =   840
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   3
      Left            =   2160
      Picture         =   "withboard1.0.frx":3C26
      Stretch         =   -1  'True
      Tag             =   "3"
      Top             =   840
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   4
      Left            =   2760
      Picture         =   "withboard1.0.frx":5108
      Stretch         =   -1  'True
      Tag             =   "4"
      Top             =   840
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   5
      Left            =   3360
      Picture         =   "withboard1.0.frx":6542
      Stretch         =   -1  'True
      Tag             =   "5"
      Top             =   840
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   6
      Left            =   3960
      Picture         =   "withboard1.0.frx":7938
      Stretch         =   -1  'True
      Tag             =   "6"
      Top             =   840
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   7
      Left            =   4560
      Picture         =   "withboard1.0.frx":8D72
      Stretch         =   -1  'True
      Tag             =   "7"
      Top             =   840
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   15
      Left            =   4560
      Picture         =   "withboard1.0.frx":A168
      Stretch         =   -1  'True
      Tag             =   "15"
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   14
      Left            =   3960
      Picture         =   "withboard1.0.frx":B5A2
      Stretch         =   -1  'True
      Tag             =   "14"
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   13
      Left            =   3360
      Picture         =   "withboard1.0.frx":C9DC
      Stretch         =   -1  'True
      Tag             =   "13"
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   12
      Left            =   2760
      Picture         =   "withboard1.0.frx":DE16
      Stretch         =   -1  'True
      Tag             =   "12"
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   11
      Left            =   2160
      Picture         =   "withboard1.0.frx":F250
      Stretch         =   -1  'True
      Tag             =   "11"
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   10
      Left            =   1560
      Picture         =   "withboard1.0.frx":1068A
      Stretch         =   -1  'True
      Tag             =   "10"
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   9
      Left            =   960
      Picture         =   "withboard1.0.frx":11AC4
      Stretch         =   -1  'True
      Tag             =   "9"
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   8
      Left            =   360
      Picture         =   "withboard1.0.frx":12EFE
      Stretch         =   -1  'True
      Tag             =   "8"
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   31
      Left            =   4560
      Picture         =   "withboard1.0.frx":14338
      Stretch         =   -1  'True
      Tag             =   "55"
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   30
      Left            =   3960
      Picture         =   "withboard1.0.frx":158AA
      Stretch         =   -1  'True
      Tag             =   "54"
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   29
      Left            =   3360
      Picture         =   "withboard1.0.frx":16E1C
      Stretch         =   -1  'True
      Tag             =   "53"
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   28
      Left            =   2760
      Picture         =   "withboard1.0.frx":1838E
      Stretch         =   -1  'True
      Tag             =   "52"
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   27
      Left            =   2160
      Picture         =   "withboard1.0.frx":19900
      Stretch         =   -1  'True
      Tag             =   "51"
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   26
      Left            =   1560
      Picture         =   "withboard1.0.frx":1AE72
      Stretch         =   -1  'True
      Tag             =   "50"
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   25
      Left            =   960
      Picture         =   "withboard1.0.frx":1C3E4
      Stretch         =   -1  'True
      Tag             =   "49"
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   24
      Left            =   360
      Picture         =   "withboard1.0.frx":1D956
      Stretch         =   -1  'True
      Tag             =   "48"
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   16
      Left            =   4560
      Picture         =   "withboard1.0.frx":1EEC8
      Stretch         =   -1  'True
      Tag             =   "63"
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   20
      Left            =   2760
      Picture         =   "withboard1.0.frx":20346
      Stretch         =   -1  'True
      Tag             =   "60"
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   21
      Left            =   1560
      Picture         =   "withboard1.0.frx":21780
      Stretch         =   -1  'True
      Tag             =   "58"
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   22
      Left            =   960
      Picture         =   "withboard1.0.frx":22BBA
      Stretch         =   -1  'True
      Tag             =   "57"
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   17
      Left            =   3960
      Picture         =   "withboard1.0.frx":2412C
      Stretch         =   -1  'True
      Tag             =   "62"
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   18
      Left            =   3360
      Picture         =   "withboard1.0.frx":2569E
      Stretch         =   -1  'True
      Tag             =   "61"
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   19
      Left            =   2160
      Picture         =   "withboard1.0.frx":26AD8
      Stretch         =   -1  'True
      Tag             =   "59"
      Top             =   5040
      Width           =   615
   End
   Begin VB.Image PieceImg 
      Height          =   615
      Index           =   23
      Left            =   360
      Picture         =   "withboard1.0.frx":2804A
      Stretch         =   -1  'True
      Tag             =   "56"
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Label30 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   540
      TabIndex        =   7
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label31 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   1170
      TabIndex        =   6
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label32 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   2970
      TabIndex        =   5
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label33 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   2370
      TabIndex        =   4
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label34 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   1770
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label35 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   3570
      TabIndex        =   2
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label36 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   4170
      TabIndex        =   1
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label37 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   4770
      TabIndex        =   0
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   5535
      Left            =   5160
      TabIndex        =   26
      Top             =   480
      Width           =   375
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   7
      Left            =   4560
      Picture         =   "withboard1.0.frx":294C8
      Top             =   840
      Width           =   600
   End
   Begin VB.Label Label28 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   480
      Width           =   4815
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   56
      Left            =   360
      Picture         =   "withboard1.0.frx":2A7CA
      Top             =   5040
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   57
      Left            =   960
      Picture         =   "withboard1.0.frx":2BACC
      Top             =   5040
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   58
      Left            =   1560
      Picture         =   "withboard1.0.frx":2CDCE
      Top             =   5040
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   59
      Left            =   2160
      Picture         =   "withboard1.0.frx":2E0D0
      Top             =   5040
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   60
      Left            =   2760
      Picture         =   "withboard1.0.frx":2F3D2
      Top             =   5040
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   61
      Left            =   3360
      Picture         =   "withboard1.0.frx":306D4
      Top             =   5040
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   62
      Left            =   3960
      Picture         =   "withboard1.0.frx":319D6
      Top             =   5040
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   63
      Left            =   4560
      Picture         =   "withboard1.0.frx":32CD8
      Top             =   5040
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   48
      Left            =   360
      Picture         =   "withboard1.0.frx":33FDA
      Top             =   4440
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   49
      Left            =   960
      Picture         =   "withboard1.0.frx":352DC
      Top             =   4440
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   50
      Left            =   1560
      Picture         =   "withboard1.0.frx":365DE
      Top             =   4440
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   51
      Left            =   2160
      Picture         =   "withboard1.0.frx":378E0
      Top             =   4440
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   52
      Left            =   2760
      Picture         =   "withboard1.0.frx":38BE2
      Top             =   4440
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   53
      Left            =   3360
      Picture         =   "withboard1.0.frx":39EE4
      Top             =   4440
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   54
      Left            =   3960
      Picture         =   "withboard1.0.frx":3B1E6
      Top             =   4440
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   55
      Left            =   4560
      Picture         =   "withboard1.0.frx":3C4E8
      Top             =   4440
      Width           =   600
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   5535
      Left            =   0
      TabIndex        =   17
      Top             =   480
      Width           =   375
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   0
      Left            =   360
      Picture         =   "withboard1.0.frx":3D7EA
      Top             =   840
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   1
      Left            =   960
      Picture         =   "withboard1.0.frx":3EAEC
      Top             =   840
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   2
      Left            =   1560
      Picture         =   "withboard1.0.frx":3FDEE
      Top             =   840
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   3
      Left            =   2160
      Picture         =   "withboard1.0.frx":410F0
      Top             =   840
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   4
      Left            =   2760
      Picture         =   "withboard1.0.frx":423F2
      Top             =   840
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   5
      Left            =   3360
      Picture         =   "withboard1.0.frx":436F4
      Top             =   840
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   6
      Left            =   3960
      Picture         =   "withboard1.0.frx":449F6
      Top             =   840
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   15
      Left            =   4560
      Picture         =   "withboard1.0.frx":45CF8
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   14
      Left            =   3960
      Picture         =   "withboard1.0.frx":46FFA
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   13
      Left            =   3360
      Picture         =   "withboard1.0.frx":482FC
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   12
      Left            =   2760
      Picture         =   "withboard1.0.frx":495FE
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   11
      Left            =   2160
      Picture         =   "withboard1.0.frx":4A900
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   10
      Left            =   1560
      Picture         =   "withboard1.0.frx":4BC02
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   9
      Left            =   960
      Picture         =   "withboard1.0.frx":4CF04
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   8
      Left            =   360
      Picture         =   "withboard1.0.frx":4E206
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   16
      Left            =   360
      Picture         =   "withboard1.0.frx":4F508
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   17
      Left            =   960
      Picture         =   "withboard1.0.frx":5080A
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   18
      Left            =   1560
      Picture         =   "withboard1.0.frx":51B0C
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   19
      Left            =   2160
      Picture         =   "withboard1.0.frx":52E0E
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   20
      Left            =   2760
      Picture         =   "withboard1.0.frx":54110
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   21
      Left            =   3360
      Picture         =   "withboard1.0.frx":55412
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   22
      Left            =   3960
      Picture         =   "withboard1.0.frx":56714
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   23
      Left            =   4560
      Picture         =   "withboard1.0.frx":57A16
      Top             =   2040
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   24
      Left            =   360
      Picture         =   "withboard1.0.frx":58D18
      Top             =   2640
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   25
      Left            =   960
      Picture         =   "withboard1.0.frx":5A01A
      Top             =   2640
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   26
      Left            =   1560
      Picture         =   "withboard1.0.frx":5B31C
      Top             =   2640
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   27
      Left            =   2160
      Picture         =   "withboard1.0.frx":5C61E
      Top             =   2640
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   28
      Left            =   2760
      Picture         =   "withboard1.0.frx":5D920
      Top             =   2640
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   29
      Left            =   3360
      Picture         =   "withboard1.0.frx":5EC22
      Top             =   2640
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   30
      Left            =   3960
      Picture         =   "withboard1.0.frx":5FF24
      Top             =   2640
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   31
      Left            =   4560
      Picture         =   "withboard1.0.frx":61226
      Top             =   2640
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   32
      Left            =   360
      Picture         =   "withboard1.0.frx":62528
      Top             =   3240
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   33
      Left            =   960
      Picture         =   "withboard1.0.frx":6382A
      Top             =   3240
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   34
      Left            =   1560
      Picture         =   "withboard1.0.frx":64B2C
      Top             =   3240
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   35
      Left            =   2160
      Picture         =   "withboard1.0.frx":65E2E
      Top             =   3240
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   36
      Left            =   2760
      Picture         =   "withboard1.0.frx":67130
      Top             =   3240
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   37
      Left            =   3360
      Picture         =   "withboard1.0.frx":68432
      Top             =   3240
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   38
      Left            =   3960
      Picture         =   "withboard1.0.frx":69734
      Top             =   3240
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   39
      Left            =   4560
      Picture         =   "withboard1.0.frx":6AA36
      Top             =   3240
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   40
      Left            =   360
      Picture         =   "withboard1.0.frx":6BD38
      Top             =   3840
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   41
      Left            =   960
      Picture         =   "withboard1.0.frx":6D03A
      Top             =   3840
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   42
      Left            =   1560
      Picture         =   "withboard1.0.frx":6E33C
      Top             =   3840
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   43
      Left            =   2160
      Picture         =   "withboard1.0.frx":6F63E
      Top             =   3840
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   44
      Left            =   2760
      Picture         =   "withboard1.0.frx":70940
      Top             =   3840
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   45
      Left            =   3360
      Picture         =   "withboard1.0.frx":71C42
      Top             =   3840
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   46
      Left            =   3960
      Picture         =   "withboard1.0.frx":72F44
      Top             =   3840
      Width           =   600
   End
   Begin VB.Image SqImg 
      Height          =   600
      Index           =   47
      Left            =   4560
      Picture         =   "withboard1.0.frx":74246
      Top             =   3840
      Width           =   600
   End
   Begin VB.Label Label29 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   360
      TabIndex        =   35
      Top             =   5640
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim From, Mto, Ftag, Totag As Integer
Dim Tosq As Boolean
Dim side11, mobility11
Dim move_list1(100) As MoveList
Dim computerside As Integer
Dim movco1 As Integer
Sub GUI_MOVE(mstr As String)
Dim mov2 As MoveList
Dim indfrm, indto
StrToMov mstr, mov2
For i = 0 To 31
If PieceImg(i).Tag = mov2.From Then
indfrm = i
Exit For
End If
Next i
If mov2.PieceT <> 0 Then
For i = 0 To 31
If PieceImg(i).Tag = mov2.ToMov Then
indto = i
Exit For
End If
Next i
PieceImg(indfrm).Left = PieceImg(indto).Left
PieceImg(indfrm).Top = PieceImg(indto).Top
PieceImg(indto).Visible = False
PieceImg(indfrm).Tag = PieceImg(indto).Tag
PieceImg(indto).Tag = -1
Else
PieceImg(indfrm).Left = SqImg(mov2.ToMov).Left
PieceImg(indfrm).Top = SqImg(mov2.ToMov).Top
'PieceImg(Mov2.ToMov).Visible = False
PieceImg(indfrm).Tag = mov2.ToMov
End If
End Sub





Private Sub Form_Load()
movco1 = 1 ' moves counter
CurnTurn = 1 ' white side to move
Dim Mobility

Initialize_Board
Call BishopMove
Call QueenMove
Call RockMove
Call KingMove
Call KnightMove
Call PawnMove

Dim c1 As Long
dp = 4
c1 = AB2(dp, 1, -1000000000, 1000000000, 1) 'this is the main function it is starting the computer to search and pick the best move ! i hope
Text2.Text = Text2.Text & movco1 & " " & StrMove & " " ' TO record all moves done
Movecount1 = 0

Call GUI_MOVE(StrMove)  ' to make the move on the screen
Call Computer_Move(MoveToDo)


Dim Mobility1 As Integer

Call Generate_Moves(0, Mobility1, move_list1)

d = GetTickCount - prevs

End Sub

Private Sub PieceImg_Click(Index As Integer)



If Not Tosq Then
From = Index
Ftag = CInt(PieceImg(Index).Tag)
Tosq = True
Else
Mto = Index
Totag = CInt(PieceImg(Index).Tag)
If ValidMove(Ftag, Totag) Then
Caption = MoveNotation
Tosq = False
PieceImg(From).Left = PieceImg(Mto).Left
PieceImg(From).Top = PieceImg(Mto).Top
PieceImg(Mto).Visible = False
PieceImg(From).Tag = PieceImg(Mto).Tag
PieceImg(Mto).Tag = -1
Text2.Text = Text2.Text & "                " & MoveNotation & vbCrLf
movco1 = movco1 + 1
prevs = GetTickCount
dp = 4
c1 = AB2(dp, 1, -1000000000, 1000000000, 1)
Call Computer_Move(MoveToDo)


Movecount1 = 0
Call GUI_MOVE(StrMove)
Text2.Text = Text2.Text & movco1 & " " & StrMove

Dim Mobility1 As Integer

Call Generate_Moves(0, Mobility1, move_list1)

d = GetTickCount - prevs

End If
End If

End Sub

Private Sub SqImg_Click(Index As Integer)

If Tosq Then
Mto = Index
Tosq = False

Mto = Index

t = Mto
If ValidMove(Ftag, Mto) Then
Caption = MoveNotation
Tosq = False
PieceImg(From).Left = SqImg(Mto).Left
PieceImg(From).Top = SqImg(Mto).Top

PieceImg(From).Tag = Index

Text2.Text = Text2.Text & "                " & MoveNotation & vbCrLf
movco1 = movco1 + 1
prevs = GetTickCount
dp = 4
c1 = AB2(dp, 1, -1000000000, 1000000000, 1)
Text2.Text = Text2.Text & movco1 & " " & StrMove & " "
Movecount1 = 0

Call GUI_MOVE(StrMove)
Dim Mobility1 As Integer
Call Computer_Move(MoveToDo)
Call Generate_Moves(0, Mobility1, move_list1)








End If

End If

End Sub
