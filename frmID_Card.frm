VERSION 5.00
Begin VB.Form frmID_Card 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ID Card Maker"
   ClientHeight    =   10065
   ClientLeft      =   1365
   ClientTop       =   780
   ClientWidth     =   12960
   Icon            =   "frmID_Card.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   177.536
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   228.6
   Begin VB.CommandButton Command17 
      Height          =   735
      Left            =   9120
      Picture         =   "frmID_Card.frx":1002
      Style           =   1  'Graphical
      TabIndex        =   157
      ToolTipText     =   "LAUNCH WEBCAM PICTURE CAPTURE"
      Top             =   7200
      Width           =   735
   End
   Begin VB.CommandButton Command16 
      Height          =   735
      Left            =   10080
      Picture         =   "frmID_Card.frx":18CC
      Style           =   1  'Graphical
      TabIndex        =   156
      ToolTipText     =   "OPEN PRINT & PRINT PREVIEW"
      Top             =   7200
      Width           =   735
   End
   Begin VB.CommandButton Command15 
      Caption         =   "P ORANGE"
      Height          =   375
      Index           =   22
      Left            =   240
      TabIndex        =   155
      Top             =   9480
      Width           =   1095
   End
   Begin VB.CommandButton Command15 
      Caption         =   "L ORANGE"
      Height          =   375
      Index           =   21
      Left            =   240
      TabIndex        =   154
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton Command15 
      Caption         =   "L WOOD"
      Height          =   375
      Index           =   20
      Left            =   1440
      TabIndex        =   153
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton Command15 
      Caption         =   "P WOOD"
      Height          =   375
      Index           =   19
      Left            =   1440
      TabIndex        =   152
      Top             =   9480
      Width           =   1095
   End
   Begin VB.CommandButton Command15 
      Caption         =   "P BW CIRC"
      Height          =   375
      Index           =   18
      Left            =   2640
      TabIndex        =   151
      Top             =   9480
      Width           =   1095
   End
   Begin VB.CommandButton Command15 
      Caption         =   "L BW CIRC"
      Height          =   375
      Index           =   17
      Left            =   2640
      TabIndex        =   150
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton Command15 
      Caption         =   "P LEAF"
      Height          =   375
      Index           =   16
      Left            =   3840
      TabIndex        =   149
      Top             =   9480
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      Caption         =   "L LEAF"
      Height          =   375
      Index           =   15
      Left            =   3840
      TabIndex        =   148
      Top             =   9000
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      Caption         =   "L LITE"
      Height          =   375
      Index           =   14
      Left            =   7080
      TabIndex        =   147
      Top             =   9480
      Width           =   855
   End
   Begin VB.CommandButton Command15 
      Caption         =   "P LITE"
      Height          =   375
      Index           =   13
      Left            =   6000
      TabIndex        =   146
      Top             =   9480
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      Caption         =   "P CHECK"
      Height          =   375
      Index           =   12
      Left            =   4920
      TabIndex        =   145
      Top             =   9480
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      Caption         =   "L PAPER"
      Height          =   375
      Index           =   11
      Left            =   6000
      TabIndex        =   144
      Top             =   9000
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      Caption         =   "P PAPER"
      Height          =   375
      Index           =   10
      Left            =   7080
      TabIndex        =   143
      Top             =   9000
      Width           =   855
   End
   Begin VB.CommandButton Command15 
      Caption         =   "P STUCCO"
      Height          =   375
      Index           =   9
      Left            =   6000
      TabIndex        =   142
      Top             =   8520
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      Caption         =   "L STUCCO"
      Height          =   375
      Index           =   8
      Left            =   4920
      TabIndex        =   141
      Top             =   8520
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      Caption         =   "PLANETS"
      Height          =   375
      Index           =   7
      Left            =   7080
      TabIndex        =   140
      Top             =   8520
      Width           =   855
   End
   Begin VB.CommandButton Command15 
      Caption         =   "L CHECK"
      Height          =   375
      Index           =   6
      Left            =   4920
      TabIndex        =   139
      Top             =   9000
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      Caption         =   "P RED"
      Height          =   375
      Index           =   5
      Left            =   7080
      TabIndex        =   138
      Top             =   8040
      Width           =   855
   End
   Begin VB.CommandButton Command15 
      Caption         =   "P GREEN"
      Height          =   375
      Index           =   4
      Left            =   6000
      TabIndex        =   137
      Top             =   8040
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      Caption         =   "P BLUE"
      Height          =   375
      Index           =   3
      Left            =   4920
      TabIndex        =   136
      Top             =   8040
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      Caption         =   "L RED"
      Height          =   375
      Index           =   2
      Left            =   7080
      TabIndex        =   135
      Top             =   7560
      Width           =   855
   End
   Begin VB.CommandButton Command15 
      Caption         =   "L GREEN"
      Height          =   375
      Index           =   1
      Left            =   6000
      TabIndex        =   134
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      Caption         =   "L BLUE"
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   132
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      Height          =   735
      Left            =   11040
      Picture         =   "frmID_Card.frx":2596
      Style           =   1  'Graphical
      TabIndex        =   121
      Top             =   7200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command12 
      Height          =   735
      Left            =   12000
      Picture         =   "frmID_Card.frx":28A0
      Style           =   1  'Graphical
      TabIndex        =   120
      Top             =   7200
      Width           =   735
   End
   Begin VB.CommandButton Command11 
      Height          =   735
      Left            =   9120
      Picture         =   "frmID_Card.frx":2BAA
      Style           =   1  'Graphical
      TabIndex        =   119
      ToolTipText     =   "OPEN DESKTOP SHORTCUT DIALOGUE"
      Top             =   9120
      Width           =   735
   End
   Begin VB.Frame Frame4 
      Caption         =   "FILE SELECTORS"
      Height          =   2295
      Left            =   7680
      TabIndex        =   113
      Top             =   0
      Width           =   5175
      Begin VB.FileListBox File1 
         Height          =   1455
         Left            =   3120
         TabIndex        =   116
         Top             =   675
         Width           =   1935
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   120
         TabIndex        =   115
         Top             =   660
         Width           =   2895
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   114
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "*.jpg"
         Height          =   255
         Left            =   4680
         TabIndex        =   117
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.CommandButton Command10 
      Height          =   735
      Left            =   11040
      Picture         =   "frmID_Card.frx":2EB4
      Style           =   1  'Graphical
      TabIndex        =   112
      ToolTipText     =   "MINIMISE THIS SCREEN"
      Top             =   9120
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Height          =   735
      Left            =   12000
      Picture         =   "frmID_Card.frx":3006
      Style           =   1  'Graphical
      TabIndex        =   110
      ToolTipText     =   "CREATE TEMPLATE FOLDER AND INSTALL JPG CONVERTER"
      Top             =   8160
      Width           =   750
   End
   Begin VB.CheckBox Check7 
      Caption         =   "USE IMAGE AS BACKGROUND"
      Height          =   615
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   108
      ToolTipText     =   "SET THE SELECTED IMAGE AS CARD BACKGROUND"
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      Height          =   735
      Left            =   10080
      Picture         =   "frmID_Card.frx":3CD0
      Style           =   1  'Graphical
      TabIndex        =   107
      ToolTipText     =   "SHOW ABOUT AND LINK"
      Top             =   9120
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Height          =   735
      Left            =   8160
      Picture         =   "frmID_Card.frx":499A
      Style           =   1  'Graphical
      TabIndex        =   105
      ToolTipText     =   "SHOW HELP"
      Top             =   9120
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Height          =   735
      Left            =   10080
      Picture         =   "frmID_Card.frx":5664
      Style           =   1  'Graphical
      TabIndex        =   104
      ToolTipText     =   "PREVIEW CURRENT CARD DESIGN"
      Top             =   8160
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Height          =   735
      Left            =   11040
      Picture         =   "frmID_Card.frx":632E
      Style           =   1  'Graphical
      TabIndex        =   103
      ToolTipText     =   "CLEAR AND RESET ALL"
      Top             =   8160
      Width           =   735
   End
   Begin VB.CheckBox Check6 
      Height          =   735
      Left            =   9120
      Picture         =   "frmID_Card.frx":6FF8
      Style           =   1  'Graphical
      TabIndex        =   94
      ToolTipText     =   "CARD EDIT LOCK"
      Top             =   8160
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Caption         =   "TEXT DISPLAY OPTIONS"
      Height          =   1095
      Left            =   5040
      TabIndex        =   75
      Top             =   5880
      Width           =   7815
      Begin VB.OptionButton optLineText 
         Caption         =   "L 12"
         Height          =   255
         Index           =   11
         Left            =   7080
         TabIndex        =   88
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton optLineText 
         Caption         =   "L 11"
         Height          =   255
         Index           =   10
         Left            =   6440
         TabIndex        =   87
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton optLineText 
         Caption         =   "L 10"
         Height          =   255
         Index           =   9
         Left            =   5808
         TabIndex        =   86
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton optLineText 
         Caption         =   "L 9"
         Height          =   255
         Index           =   8
         Left            =   5176
         TabIndex        =   85
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   84
         Text            =   "Label Text"
         Top             =   720
         Width           =   7575
      End
      Begin VB.OptionButton optLineText 
         Caption         =   "L 1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   83
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton optLineText 
         Caption         =   "L 2"
         Height          =   255
         Index           =   1
         Left            =   752
         TabIndex        =   82
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton optLineText 
         Caption         =   "L 3"
         Height          =   255
         Index           =   2
         Left            =   1384
         TabIndex        =   81
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton optLineText 
         Caption         =   "L 4"
         Height          =   255
         Index           =   3
         Left            =   2016
         TabIndex        =   80
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton optLineText 
         Caption         =   "L 5"
         Height          =   255
         Index           =   4
         Left            =   2648
         TabIndex        =   79
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton optLineText 
         Caption         =   "L 6"
         Height          =   255
         Index           =   5
         Left            =   3280
         TabIndex        =   78
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton optLineText 
         Caption         =   "L 7"
         Height          =   255
         Index           =   6
         Left            =   3912
         TabIndex        =   77
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton optLineText 
         Caption         =   "L 8"
         Height          =   255
         Index           =   7
         Left            =   4544
         TabIndex        =   76
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      Height          =   735
      Left            =   6840
      Picture         =   "frmID_Card.frx":7302
      Style           =   1  'Graphical
      TabIndex        =   74
      ToolTipText     =   "MAKE NEW FOLDER"
      Top             =   120
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "FONT OPTIONS"
      Height          =   3015
      Left            =   10080
      TabIndex        =   59
      Top             =   2520
      Width           =   2775
      Begin VB.CommandButton Command4 
         Height          =   375
         Left            =   2280
         Picture         =   "frmID_Card.frx":760C
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Left            =   1848
         Picture         =   "frmID_Card.frx":7B96
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   2520
         Width           =   375
      End
      Begin VB.ComboBox cmbTxtColour 
         Height          =   315
         ItemData        =   "frmID_Card.frx":8120
         Left            =   1440
         List            =   "frmID_Card.frx":813C
         MouseIcon       =   "frmID_Card.frx":8177
         MousePointer    =   99  'Custom
         TabIndex        =   97
         Text            =   "Black"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CheckBox Check5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         Height          =   375
         Left            =   1416
         Picture         =   "frmID_Card.frx":82C9
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   2520
         Width           =   375
      End
      Begin VB.CheckBox Check4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   984
         Picture         =   "frmID_Card.frx":8853
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   2520
         Width           =   375
      End
      Begin VB.CheckBox Check3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   552
         Picture         =   "frmID_Card.frx":8DDD
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   2520
         Width           =   375
      End
      Begin VB.CheckBox Check2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Picture         =   "frmID_Card.frx":9367
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cmbFontSize 
         Height          =   315
         ItemData        =   "frmID_Card.frx":98F1
         Left            =   1800
         List            =   "frmID_Card.frx":9A18
         MouseIcon       =   "frmID_Card.frx":9B9B
         MousePointer    =   99  'Custom
         TabIndex        =   62
         Text            =   "8"
         Top             =   1170
         Width           =   855
      End
      Begin VB.ComboBox cmbFont 
         Height          =   315
         Left            =   1080
         MouseIcon       =   "frmID_Card.frx":9CED
         MousePointer    =   99  'Custom
         Sorted          =   -1  'True
         TabIndex        =   60
         Text            =   "MS Sans Serif"
         Top             =   690
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Select Font Colour:"
         Height          =   375
         Left            =   120
         TabIndex        =   100
         Top             =   1650
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Select Font Size:"
         Height          =   255
         Left            =   120
         TabIndex        =   99
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Select Font:"
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "No Of Available Fonts:"
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   255
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   -720
         X2              =   3360
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         X1              =   3360
         X2              =   -240
         Y1              =   2400
         Y2              =   2400
      End
   End
   Begin VB.CommandButton Command2 
      Height          =   735
      Left            =   12000
      Picture         =   "frmID_Card.frx":9E3F
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "EXIT THIS APPLICATION:- BY THE NEAREST DOOR!"
      Top             =   9120
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   480
      TabIndex        =   57
      ToolTipText     =   "CREATE NEW FOLDER"
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   56
      Text            =   "Type new folder name here"
      Top             =   480
      Width           =   5895
   End
   Begin VB.PictureBox VRuler 
      AutoRedraw      =   -1  'True
      Height          =   4252
      Left            =   120
      ScaleHeight     =   4185
      ScaleWidth      =   285
      TabIndex        =   6
      Top             =   1560
      Width           =   345
   End
   Begin VB.PictureBox HRuler 
      AutoRedraw      =   -1  'True
      Height          =   345
      Left            =   480
      ScaleHeight     =   285
      ScaleWidth      =   7020
      TabIndex        =   5
      Top             =   1200
      Width           =   7087
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   4252
      Left            =   480
      ScaleHeight     =   4185
      ScaleWidth      =   7020
      TabIndex        =   2
      Top             =   1560
      Width           =   7087
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   93
         Top             =   3840
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   92
         Top             =   3600
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   91
         Top             =   3360
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   90
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   69
         Top             =   2880
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   68
         Top             =   2640
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Index           =   5
         Left            =   840
         TabIndex        =   67
         Top             =   3840
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Index           =   4
         Left            =   840
         TabIndex        =   66
         Top             =   3600
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Index           =   3
         Left            =   840
         TabIndex        =   65
         Top             =   3360
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   64
         Top             =   3120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   61
         Top             =   2880
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   55
         Top             =   2640
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   1665
         Left            =   360
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label Node 
         BackColor       =   &H00C00000&
         Height          =   90
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   90
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ID CARD BACKGROUND"
      Height          =   2295
      Left            =   7800
      TabIndex        =   0
      Top             =   2520
      Width           =   2175
      Begin VB.PictureBox picBG_Col 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   48
         Left            =   1800
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   89
         ToolTipText     =   "DEFAULT BACKGROUND COLOUR"
         Top             =   1920
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00400040&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   47
         Left            =   1800
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   53
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00800080&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   46
         Left            =   1800
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   52
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00C000C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   45
         Left            =   1800
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   51
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   44
         Left            =   1800
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   50
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   43
         Left            =   1800
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   49
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   42
         Left            =   1800
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   48
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   41
         Left            =   1560
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   47
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   40
         Left            =   1560
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   46
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00C00000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   39
         Left            =   1560
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   45
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   38
         Left            =   1560
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   44
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   37
         Left            =   1560
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   43
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   36
         Left            =   1560
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   42
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00404000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   35
         Left            =   1320
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   41
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   34
         Left            =   1320
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   40
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00C0C000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   33
         Left            =   1320
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   39
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   32
         Left            =   1320
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   38
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   31
         Left            =   1320
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   37
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   30
         Left            =   1320
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   36
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00004000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   29
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   35
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   28
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   34
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   27
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   33
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   26
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   32
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   25
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   31
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   24
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   30
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00004040&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   23
         Left            =   840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   29
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00008080&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   22
         Left            =   840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   28
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H0000C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   21
         Left            =   840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   27
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   20
         Left            =   840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   26
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   19
         Left            =   840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   25
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   18
         Left            =   840
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   24
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00004080&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   17
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   23
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00404080&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   16
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   22
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H000040C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   15
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   21
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   14
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   20
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   13
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   19
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   12
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   18
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   11
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   17
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00000080&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   10
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   16
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H000000C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   9
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   15
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   8
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   14
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   7
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   13
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   6
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   12
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   5
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   11
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   10
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   9
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   8
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   7
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picBG_Col 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   1
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   1920
         Width           =   45
      End
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      Height          =   4252
      Left            =   600
      ScaleHeight     =   4185
      ScaleWidth      =   7020
      TabIndex        =   109
      Top             =   1320
      Visible         =   0   'False
      Width           =   7087
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   0
      Left            =   11880
      TabIndex        =   111
      Text            =   "No"
      Top             =   8400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   1575
      Left            =   3840
      TabIndex        =   122
      Top             =   2040
      Visible         =   0   'False
      Width           =   3015
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         LargeChange     =   5
         Left            =   120
         Max             =   255
         MouseIcon       =   "frmID_Card.frx":A149
         MousePointer    =   99  'Custom
         TabIndex        =   127
         Top             =   1200
         Width           =   2175
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   5
         Left            =   120
         Max             =   255
         MouseIcon       =   "frmID_Card.frx":A29B
         MousePointer    =   99  'Custom
         TabIndex        =   126
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2400
         TabIndex        =   125
         Text            =   "0"
         Top             =   1185
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2400
         TabIndex        =   124
         Text            =   "0"
         Top             =   705
         Width           =   495
      End
      Begin VB.CommandButton Command14 
         Height          =   735
         Left            =   2040
         Picture         =   "frmID_Card.frx":A3ED
         Style           =   1  'Graphical
         TabIndex        =   123
         ToolTipText     =   "ADD GRADIENT BACKGROUNG"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CARD GRADIENT FILL OPTIONS"
         Height          =   255
         Left            =   120
         TabIndex        =   131
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Adjust Top Colour"
         Height          =   255
         Left            =   120
         TabIndex        =   130
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Adjust Lower Colour"
         Height          =   255
         Left            =   120
         TabIndex        =   129
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label14 
         Caption         =   "ADD THE GRADIENT AS THE LAST  STEP IN THE CARD DESIGN"
         Height          =   615
         Left            =   120
         TabIndex        =   128
         Top             =   720
         Width           =   1935
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      Index           =   20
      X1              =   158.75
      X2              =   158.75
      Y1              =   141.817
      Y2              =   124.883
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      Index           =   22
      X1              =   2.117
      X2              =   2.117
      Y1              =   175.684
      Y2              =   156.633
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      Index           =   18
      X1              =   2.117
      X2              =   84.667
      Y1              =   156.633
      Y2              =   156.633
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      Index           =   16
      X1              =   2.117
      X2              =   141.817
      Y1              =   175.684
      Y2              =   175.684
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      Index           =   14
      X1              =   84.667
      X2              =   84.667
      Y1              =   156.633
      Y2              =   124.883
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PRESET  BACKGROUNDS"
      Height          =   255
      Left            =   5130
      TabIndex        =   133
      Top             =   7215
      Width           =   2865
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      Index           =   12
      X1              =   84.667
      X2              =   141.817
      Y1              =   124.883
      Y2              =   124.883
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      Index           =   10
      X1              =   141.817
      X2              =   226.484
      Y1              =   141.817
      Y2              =   141.817
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      Index           =   8
      X1              =   158.75
      X2              =   158.75
      Y1              =   175.684
      Y2              =   158.75
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      Index           =   6
      X1              =   175.684
      X2              =   175.684
      Y1              =   175.684
      Y2              =   124.883
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      Index           =   4
      X1              =   209.55
      X2              =   209.55
      Y1              =   175.684
      Y2              =   124.883
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      Index           =   2
      X1              =   141.817
      X2              =   226.484
      Y1              =   158.75
      Y2              =   158.75
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      Index           =   0
      X1              =   192.617
      X2              =   192.617
      Y1              =   175.684
      Y2              =   124.883
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000003&
      Height          =   2895
      Index           =   0
      Left            =   8040
      Top             =   7080
      Width           =   4815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Height          =   2895
      Index           =   1
      Left            =   8040
      Top             =   7080
      Width           =   4815
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Label10"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3480
      TabIndex        =   118
      Top             =   960
      Width           =   570
   End
   Begin VB.Label Label9 
      Height          =   255
      Left            =   480
      TabIndex        =   106
      Top             =   600
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.Image imgBuffer 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1665
      Left            =   7320
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lock Card Edit"
      Height          =   630
      Left            =   8400
      TabIndex        =   95
      Top             =   8212
      Width           =   420
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   6255
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   1
      X1              =   192.617
      X2              =   192.617
      Y1              =   175.684
      Y2              =   124.883
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   3
      X1              =   141.817
      X2              =   226.484
      Y1              =   158.75
      Y2              =   158.75
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   9
      X1              =   158.75
      X2              =   158.75
      Y1              =   175.684
      Y2              =   158.75
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   7
      X1              =   175.684
      X2              =   175.684
      Y1              =   175.684
      Y2              =   124.883
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   5
      X1              =   209.55
      X2              =   209.55
      Y1              =   175.684
      Y2              =   124.883
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   11
      X1              =   141.817
      X2              =   226.484
      Y1              =   141.817
      Y2              =   141.817
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   13
      X1              =   84.667
      X2              =   141.817
      Y1              =   124.883
      Y2              =   124.883
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   15
      X1              =   84.667
      X2              =   84.667
      Y1              =   156.633
      Y2              =   124.883
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   17
      X1              =   2.117
      X2              =   141.817
      Y1              =   175.684
      Y2              =   175.684
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   19
      X1              =   2.117
      X2              =   84.667
      Y1              =   156.633
      Y2              =   156.633
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   23
      X1              =   2.117
      X2              =   2.117
      Y1              =   175.684
      Y2              =   156.633
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   21
      X1              =   158.75
      X2              =   158.75
      Y1              =   141.817
      Y2              =   124.883
   End
End
Attribute VB_Name = "frmID_Card"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Dim blues, greens, reds, colours

Dim X1 As Single
Dim Y1 As Single

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private CurX As Double
Private CurY As Double

'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 
         Private Const WM_PAINT = &HF
        Private Const WM_PRINT = &H317
        Private Const PRF_CLIENT = &H4&    ' Draw the window's client area
        Private Const PRF_CHILDREN = &H10& ' Draw all visible child windows
        Private Const PRF_OWNED = &H20&    ' Draw all owned windows
  Const myerrfilepath = 53
  Const errfilepath = 75
    Const rterror = 68
    
Private Sub Check1_Click()
'*******************************************************
'*        enables the make new folder button           *
'*******************************************************
If Check1.Value = 1 Then
Command1.Enabled = True
Else: Command1.Enabled = False
End If
End Sub

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*              set the check button cursor            *
'*******************************************************
Check1.MousePointer = 99
Check1.MouseIcon = LoadResPicture(105, vbResCursor)
End Sub

Private Sub Check2_Click()
'*******************************************************
'*        highlight the selected label in bold         *
'*******************************************************
If optLineText(0).Value = True Then
If Check2.Value = 1 Then
lblText(0).FontBold = True
Else: lblText(0).FontBold = False
End If
End If
If optLineText(1).Value = True Then
If Check2.Value = 1 Then
lblText(1).FontBold = True
Else: lblText(1).FontBold = False
End If
End If
If optLineText(2).Value = True Then
If Check2.Value = 1 Then
lblText(2).FontBold = True
Else: lblText(2).FontBold = False
End If
End If
If optLineText(3).Value = True Then
If Check2.Value = 1 Then
lblText(3).FontBold = True
Else: lblText(3).FontBold = False
End If
End If
If optLineText(4).Value = True Then
If Check2.Value = 1 Then
lblText(4).FontBold = True
Else: lblText(4).FontBold = False
End If
End If
If optLineText(5).Value = True Then
If Check2.Value = 1 Then
lblText(5).FontBold = True
Else: lblText(5).FontBold = False
End If
End If
If optLineText(6).Value = True Then
If Check2.Value = 1 Then
lblText(6).FontBold = True
Else: lblText(6).FontBold = False
End If
End If
If optLineText(7).Value = True Then
If Check2.Value = 1 Then
lblText(7).FontBold = True
Else: lblText(7).FontBold = False
End If
End If
If optLineText(8).Value = True Then
If Check2.Value = 1 Then
lblText(8).FontBold = True
Else: lblText(8).FontBold = False
End If
End If
If optLineText(9).Value = True Then
If Check2.Value = 1 Then
lblText(9).FontBold = True
Else: lblText(9).FontBold = False
End If
End If
If optLineText(10).Value = True Then
If Check2.Value = 1 Then
lblText(10).FontBold = True
Else: lblText(10).FontBold = False
End If
End If
If optLineText(11).Value = True Then
If Check2.Value = 1 Then
lblText(11).FontBold = True
Else: lblText(11).FontBold = False
End If
End If
End Sub

Private Sub Check2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*              set the check button cursor            *
'*******************************************************
Check2.MousePointer = 99
Check2.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Check3_Click()
'*******************************************************
'*       highlight the selected label in italic        *
'*******************************************************
If optLineText(0).Value = True Then
If Check3.Value = 1 Then
lblText(0).FontItalic = True
Else: lblText(0).FontItalic = False
End If
End If
If optLineText(1).Value = True Then
If Check3.Value = 1 Then
lblText(1).FontItalic = True
Else: lblText(1).FontItalic = False
End If
End If
If optLineText(2).Value = True Then
If Check3.Value = 1 Then
lblText(2).FontItalic = True
Else: lblText(2).FontItalic = False
End If
End If
If optLineText(3).Value = True Then
If Check3.Value = 1 Then
lblText(3).FontItalic = True
Else: lblText(3).FontItalic = False
End If
End If
If optLineText(4).Value = True Then
If Check3.Value = 1 Then
lblText(4).FontItalic = True
Else: lblText(4).FontItalic = False
End If
End If
If optLineText(5).Value = True Then
If Check3.Value = 1 Then
lblText(5).FontItalic = True
Else: lblText(5).FontItalic = False
End If
End If
If optLineText(6).Value = True Then
If Check3.Value = 1 Then
lblText(6).FontItalic = True
Else: lblText(6).FontItalic = False
End If
End If
If optLineText(7).Value = True Then
If Check3.Value = 1 Then
lblText(7).FontItalic = True
Else: lblText(7).FontItalic = False
End If
End If
If optLineText(8).Value = True Then
If Check3.Value = 1 Then
lblText(8).FontItalic = True
Else: lblText(8).FontItalic = False
End If
End If
If optLineText(9).Value = True Then
If Check3.Value = 1 Then
lblText(9).FontItalic = True
Else: lblText(9).FontItalic = False
End If
End If
If optLineText(10).Value = True Then
If Check3.Value = 1 Then
lblText(10).FontItalic = True
Else: lblText(10).FontItalic = False
End If
End If
If optLineText(11).Value = True Then
If Check3.Value = 1 Then
lblText(11).FontItalic = True
Else: lblText(11).FontItalic = False
End If
End If
End Sub

Private Sub Check3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*              set the check button cursor            *
'*******************************************************
Check3.MousePointer = 99
Check3.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Check4_Click()
'*******************************************************
'*     highlight the selected label in underline       *
'*******************************************************
If optLineText(0).Value = True Then
If Check4.Value = 1 Then
lblText(0).FontUnderline = True
ElseIf Check4.Value = 0 Then
lblText(0).FontUnderline = False
End If
End If
If optLineText(1).Value = True Then
If Check4.Value = 1 Then
lblText(1).FontUnderline = True
ElseIf Check4.Value = 0 Then
lblText(1).FontUnderline = False
End If
End If
If optLineText(2).Value = True Then
If Check4.Value = 1 Then
lblText(2).FontUnderline = True
ElseIf Check4.Value = 0 Then
lblText(2).FontUnderline = False
End If
End If
If optLineText(3).Value = True Then
If Check4.Value = 1 Then
lblText(3).FontUnderline = True
ElseIf Check4.Value = 0 Then
lblText(3).FontUnderline = False
End If
End If
If optLineText(4).Value = True Then
If Check4.Value = 1 Then
lblText(4).FontUnderline = True
ElseIf Check4.Value = 0 Then
lblText(4).FontUnderline = False
End If
End If
If optLineText(5).Value = True Then
If Check4.Value = 1 Then
lblText(5).FontUnderline = True
ElseIf Check4.Value = 0 Then
lblText(5).FontUnderline = False
End If
End If
If optLineText(6).Value = True Then
If Check4.Value = 1 Then
lblText(6).FontUnderline = True
ElseIf Check4.Value = 0 Then
lblText(6).FontUnderline = False
End If
End If
If optLineText(7).Value = True Then
If Check4.Value = 1 Then
lblText(7).FontUnderline = True
ElseIf Check4.Value = 0 Then
lblText(7).FontUnderline = False
End If
End If
If optLineText(8).Value = True Then
If Check4.Value = 1 Then
lblText(8).FontUnderline = True
ElseIf Check4.Value = 0 Then
lblText(8).FontUnderline = False
End If
End If
If optLineText(9).Value = True Then
If Check4.Value = 1 Then
lblText(9).FontUnderline = True
ElseIf Check4.Value = 0 Then
lblText(9).FontUnderline = False
End If
End If
If optLineText(10).Value = True Then
If Check4.Value = 1 Then
lblText(10).FontUnderline = True
ElseIf Check4.Value = 0 Then
lblText(10).FontUnderline = False
End If
End If
If optLineText(11).Value = True Then
If Check4.Value = 1 Then
lblText(11).FontUnderline = True
ElseIf Check4.Value = 0 Then
lblText(11).FontUnderline = False
End If
End If
End Sub

Private Sub Check4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*              set the check button cursor            *
'*******************************************************
Check4.MousePointer = 99
Check4.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Check5_Click()
'*******************************************************
'*     highlight the selected label in strikethru      *
'*******************************************************
If optLineText(0).Value = True Then
If Check5.Value = 1 Then
lblText(0).FontStrikethru = True
ElseIf Check5.Value = 0 Then
lblText(0).FontStrikethru = False
End If
End If
If optLineText(1).Value = True Then
If Check5.Value = 1 Then
lblText(1).FontStrikethru = True
ElseIf Check5.Value = 0 Then
lblText(1).FontStrikethru = False
End If
End If
If optLineText(2).Value = True Then
If Check5.Value = 1 Then
lblText(2).FontStrikethru = True
ElseIf Check5.Value = 0 Then
lblText(2).FontStrikethru = False
End If
End If
If optLineText(3).Value = True Then
If Check5.Value = 1 Then
lblText(3).FontStrikethru = True
ElseIf Check5.Value = 0 Then
lblText(3).FontStrikethru = False
End If
End If
If optLineText(4).Value = True Then
If Check5.Value = 1 Then
lblText(4).FontStrikethru = True
ElseIf Check5.Value = 0 Then
lblText(4).FontStrikethru = False
End If
End If
If optLineText(5).Value = True Then
If Check5.Value = 1 Then
lblText(5).FontStrikethru = True
ElseIf Check5.Value = 0 Then
lblText(5).FontStrikethru = False
End If
End If
If optLineText(6).Value = True Then
If Check5.Value = 1 Then
lblText(6).FontStrikethru = True
ElseIf Check5.Value = 0 Then
lblText(6).FontStrikethru = False
End If
End If
If optLineText(7).Value = True Then
If Check5.Value = 1 Then
lblText(7).FontStrikethru = True
ElseIf Check5.Value = 0 Then
lblText(7).FontStrikethru = False
End If
End If
If optLineText(8).Value = True Then
If Check5.Value = 1 Then
lblText(8).FontStrikethru = True
ElseIf Check5.Value = 0 Then
lblText(8).FontStrikethru = False
End If
End If
If optLineText(9).Value = True Then
If Check5.Value = 1 Then
lblText(9).FontStrikethru = True
ElseIf Check5.Value = 0 Then
lblText(9).FontStrikethru = False
End If
End If
If optLineText(10).Value = True Then
If Check5.Value = 1 Then
lblText(10).FontStrikethru = True
ElseIf Check5.Value = 0 Then
lblText(10).FontStrikethru = False
End If
End If
If optLineText(11).Value = True Then
If Check5.Value = 1 Then
lblText(11).FontStrikethru = True
ElseIf Check5.Value = 0 Then
lblText(11).FontStrikethru = False
End If
End If
End Sub

Private Sub Check5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*              set the check button cursor            *
'*******************************************************
Check5.MousePointer = 99
Check5.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Check6_Click()
'*******************************************************
'*     lock or unlock the picture box for editing      *
'*    also lock the colour picker and font options     *
'*     and the preset background selector buttons      *
'*******************************************************
Dim B As Integer

If Check6.Value = 0 Then
Picture1.Enabled = True
Check6.Picture = LoadResPicture(102, vbResIcon)
Label1.Caption = "Lock Card Edit"
Frame4.Enabled = True
Frame3.Enabled = True
Frame2.Enabled = True
Frame1.Enabled = True
If Command12.Visible = True Then
Command12.Enabled = True
End If
If Command13.Visible = True Then
Command13.Enabled = True
End If
HideNodes
Check7.Enabled = True
Command14.Enabled = True
For B = 0 To 20
Command15(B).Enabled = True
Next B
Else
Picture1.Enabled = False
Check6.Picture = LoadResPicture(103, vbResIcon)
Label1.Caption = "Open Card Edit"
Frame4.Enabled = False
Frame3.Enabled = False
Frame2.Enabled = False
Frame1.Enabled = False
If Command12.Visible = True Then
Command12.Enabled = False
End If
If Command13.Visible = True Then
Command13.Enabled = False
End If
HideNodes
Check7.Enabled = False
Command14.Enabled = False

For B = 0 To 20
Command15(B).Enabled = False
Next B

End If
End Sub

Private Sub Check6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*              set the check button cursor            *
'*******************************************************
Check6.MousePointer = 99
Check6.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Check7_Click()
'*******************************************************
'*                set the background image             *
'*******************************************************
End Sub

Private Sub Check7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*              set the check button cursor            *
'*******************************************************
Check7.MousePointer = 99
Check7.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub cmbTxtColour_Click()
'*******************************************************
'*        change the selected label font colour        *
'*******************************************************
If optLineText(0).Value = True Then
If cmbTxtColour.Text = "Black" Then
lblText(0).ForeColor = vbBlack
ElseIf cmbTxtColour.Text = "Red" Then
lblText(0).ForeColor = vbRed
ElseIf cmbTxtColour.Text = "Green" Then
lblText(0).ForeColor = vbGreen
ElseIf cmbTxtColour.Text = "Yellow" Then
lblText(0).ForeColor = vbYellow
ElseIf cmbTxtColour.Text = "Blue" Then
lblText(0).ForeColor = vbBlue
ElseIf cmbTxtColour.Text = "Magenta" Then
lblText(0).ForeColor = vbMagenta
ElseIf cmbTxtColour.Text = "Cyan" Then
lblText(0).ForeColor = vbCyan
ElseIf cmbTxtColour.Text = "White" Then
lblText(0).ForeColor = vbWhite
End If
ElseIf optLineText(1).Value = True Then '****************************************
If cmbTxtColour.Text = "Black" Then
lblText(1).ForeColor = vbBlack
ElseIf cmbTxtColour.Text = "Red" Then
lblText(1).ForeColor = vbRed
ElseIf cmbTxtColour.Text = "Green" Then
lblText(1).ForeColor = vbGreen
ElseIf cmbTxtColour.Text = "Yellow" Then
lblText(1).ForeColor = vbYellow
ElseIf cmbTxtColour.Text = "Blue" Then
lblText(1).ForeColor = vbBlue
ElseIf cmbTxtColour.Text = "Magenta" Then
lblText(1).ForeColor = vbMagenta
ElseIf cmbTxtColour.Text = "Cyan" Then
lblText(1).ForeColor = vbCyan
ElseIf cmbTxtColour.Text = "White" Then
lblText(0).ForeColor = vbWhite
End If
ElseIf optLineText(2).Value = True Then '****************************************
If cmbTxtColour.Text = "Black" Then
lblText(2).ForeColor = vbBlack
ElseIf cmbTxtColour.Text = "Red" Then
lblText(2).ForeColor = vbRed
ElseIf cmbTxtColour.Text = "Green" Then
lblText(2).ForeColor = vbGreen
ElseIf cmbTxtColour.Text = "Yellow" Then
lblText(2).ForeColor = vbYellow
ElseIf cmbTxtColour.Text = "Blue" Then
lblText(2).ForeColor = vbBlue
ElseIf cmbTxtColour.Text = "Magenta" Then
lblText(2).ForeColor = vbMagenta
ElseIf cmbTxtColour.Text = "Cyan" Then
lblText(2).ForeColor = vbCyan
ElseIf cmbTxtColour.Text = "White" Then
lblText(2).ForeColor = vbWhite
End If
ElseIf optLineText(3).Value = True Then '****************************************
If cmbTxtColour.Text = "Black" Then
lblText(3).ForeColor = vbBlack
ElseIf cmbTxtColour.Text = "Red" Then
lblText(3).ForeColor = vbRed
ElseIf cmbTxtColour.Text = "Green" Then
lblText(3).ForeColor = vbGreen
ElseIf cmbTxtColour.Text = "Yellow" Then
lblText(3).ForeColor = vbYellow
ElseIf cmbTxtColour.Text = "Blue" Then
lblText(3).ForeColor = vbBlue
ElseIf cmbTxtColour.Text = "Magenta" Then
lblText(3).ForeColor = vbMagenta
ElseIf cmbTxtColour.Text = "Cyan" Then
lblText(3).ForeColor = vbCyan
ElseIf cmbTxtColour.Text = "White" Then
lblText(3).ForeColor = vbWhite
End If
ElseIf optLineText(4).Value = True Then '****************************************
If cmbTxtColour.Text = "Black" Then
lblText(4).ForeColor = vbBlack
ElseIf cmbTxtColour.Text = "Red" Then
lblText(4).ForeColor = vbRed
ElseIf cmbTxtColour.Text = "Green" Then
lblText(4).ForeColor = vbGreen
ElseIf cmbTxtColour.Text = "Yellow" Then
lblText(4).ForeColor = vbYellow
ElseIf cmbTxtColour.Text = "Blue" Then
lblText(4).ForeColor = vbBlue
ElseIf cmbTxtColour.Text = "Magenta" Then
lblText(4).ForeColor = vbMagenta
ElseIf cmbTxtColour.Text = "Cyan" Then
lblText(4).ForeColor = vbCyan
ElseIf cmbTxtColour.Text = "White" Then
lblText(4).ForeColor = vbWhite
End If
ElseIf optLineText(5).Value = True Then '****************************************
If cmbTxtColour.Text = "Black" Then
lblText(5).ForeColor = vbBlack
ElseIf cmbTxtColour.Text = "Red" Then
lblText(5).ForeColor = vbRed
ElseIf cmbTxtColour.Text = "Green" Then
lblText(5).ForeColor = vbGreen
ElseIf cmbTxtColour.Text = "Yellow" Then
lblText(5).ForeColor = vbYellow
ElseIf cmbTxtColour.Text = "Blue" Then
lblText(5).ForeColor = vbBlue
ElseIf cmbTxtColour.Text = "Magenta" Then
lblText(5).ForeColor = vbMagenta
ElseIf cmbTxtColour.Text = "Cyan" Then
lblText(5).ForeColor = vbCyan
ElseIf cmbTxtColour.Text = "White" Then
lblText(5).ForeColor = vbWhite
End If
ElseIf optLineText(6).Value = True Then '****************************************
If cmbTxtColour.Text = "Black" Then
lblText(6).ForeColor = vbBlack
ElseIf cmbTxtColour.Text = "Red" Then
lblText(6).ForeColor = vbRed
ElseIf cmbTxtColour.Text = "Green" Then
lblText(6).ForeColor = vbGreen
ElseIf cmbTxtColour.Text = "Yellow" Then
lblText(6).ForeColor = vbYellow
ElseIf cmbTxtColour.Text = "Blue" Then
lblText(6).ForeColor = vbBlue
ElseIf cmbTxtColour.Text = "Magenta" Then
lblText(6).ForeColor = vbMagenta
ElseIf cmbTxtColour.Text = "Cyan" Then
lblText(6).ForeColor = vbCyan
ElseIf cmbTxtColour.Text = "White" Then
lblText(6).ForeColor = vbWhite
End If
ElseIf optLineText(7).Value = True Then '****************************************
If cmbTxtColour.Text = "Black" Then
lblText(7).ForeColor = vbBlack
ElseIf cmbTxtColour.Text = "Red" Then
lblText(7).ForeColor = vbRed
ElseIf cmbTxtColour.Text = "Green" Then
lblText(7).ForeColor = vbGreen
ElseIf cmbTxtColour.Text = "Yellow" Then
lblText(7).ForeColor = vbYellow
ElseIf cmbTxtColour.Text = "Blue" Then
lblText(7).ForeColor = vbBlue
ElseIf cmbTxtColour.Text = "Magenta" Then
lblText(7).ForeColor = vbMagenta
ElseIf cmbTxtColour.Text = "Cyan" Then
lblText(7).ForeColor = vbCyan
ElseIf cmbTxtColour.Text = "White" Then
lblText(7).ForeColor = vbWhite
End If
ElseIf optLineText(8).Value = True Then '****************************************
If cmbTxtColour.Text = "Black" Then
lblText(8).ForeColor = vbBlack
ElseIf cmbTxtColour.Text = "Red" Then
lblText(8).ForeColor = vbRed
ElseIf cmbTxtColour.Text = "Green" Then
lblText(8).ForeColor = vbGreen
ElseIf cmbTxtColour.Text = "Yellow" Then
lblText(8).ForeColor = vbYellow
ElseIf cmbTxtColour.Text = "Blue" Then
lblText(8).ForeColor = vbBlue
ElseIf cmbTxtColour.Text = "Magenta" Then
lblText(8).ForeColor = vbMagenta
ElseIf cmbTxtColour.Text = "Cyan" Then
lblText(8).ForeColor = vbCyan
ElseIf cmbTxtColour.Text = "White" Then
lblText(8).ForeColor = vbWhite
End If
ElseIf optLineText(9).Value = True Then '****************************************
If cmbTxtColour.Text = "Black" Then
lblText(9).ForeColor = vbBlack
ElseIf cmbTxtColour.Text = "Red" Then
lblText(9).ForeColor = vbRed
ElseIf cmbTxtColour.Text = "Green" Then
lblText(9).ForeColor = vbGreen
ElseIf cmbTxtColour.Text = "Yellow" Then
lblText(9).ForeColor = vbYellow
ElseIf cmbTxtColour.Text = "Blue" Then
lblText(9).ForeColor = vbBlue
ElseIf cmbTxtColour.Text = "Magenta" Then
lblText(9).ForeColor = vbMagenta
ElseIf cmbTxtColour.Text = "Cyan" Then
lblText(9).ForeColor = vbCyan
ElseIf cmbTxtColour.Text = "White" Then
lblText(9).ForeColor = vbWhite
End If
ElseIf optLineText(10).Value = True Then '****************************************
If cmbTxtColour.Text = "Black" Then
lblText(10).ForeColor = vbBlack
ElseIf cmbTxtColour.Text = "Red" Then
lblText(10).ForeColor = vbRed
ElseIf cmbTxtColour.Text = "Green" Then
lblText(10).ForeColor = vbGreen
ElseIf cmbTxtColour.Text = "Yellow" Then
lblText(10).ForeColor = vbYellow
ElseIf cmbTxtColour.Text = "Blue" Then
lblText(10).ForeColor = vbBlue
ElseIf cmbTxtColour.Text = "Magenta" Then
lblText(10).ForeColor = vbMagenta
ElseIf cmbTxtColour.Text = "Cyan" Then
lblText(10).ForeColor = vbCyan
ElseIf cmbTxtColour.Text = "White" Then
lblText(10).ForeColor = vbWhite
End If
ElseIf optLineText(11).Value = True Then '****************************************
If cmbTxtColour.Text = "Black" Then
lblText(11).ForeColor = vbBlack
ElseIf cmbTxtColour.Text = "Red" Then
lblText(11).ForeColor = vbRed
ElseIf cmbTxtColour.Text = "Green" Then
lblText(11).ForeColor = vbGreen
ElseIf cmbTxtColour.Text = "Yellow" Then
lblText(11).ForeColor = vbYellow
ElseIf cmbTxtColour.Text = "Blue" Then
lblText(11).ForeColor = vbBlue
ElseIf cmbTxtColour.Text = "Magenta" Then
lblText(11).ForeColor = vbMagenta
ElseIf cmbTxtColour.Text = "Cyan" Then
lblText(11).ForeColor = vbCyan
ElseIf cmbTxtColour.Text = "White" Then
lblText(11).ForeColor = vbWhite
End If
End If
End Sub

Private Sub Command1_Click()
'*******************************************************
'*           make new folder command button            *
'*******************************************************
MkDir Label4.Caption & "\" & Text1.Text & "\"

Label9.Caption = Label4.Caption & "\" & Text1.Text

Check1.Value = 0
End Sub

Private Sub Command10_Click()
'*******************************************************
'*                minimise main form                   *
'*******************************************************
Me.WindowState = vbMinimized
End Sub

Private Sub Command10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command10.MousePointer = 99
Command10.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command11_Click()
'********************************************************
'*      show place shortcut on the desktop screen       *
'********************************************************
frmDTI.Show
End Sub

Private Sub Command11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command11.MousePointer = 99
Command11.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command12_Click()
'*******************************************************
'*             set to portrait orientation             *
'*******************************************************
Picture1.Width = 75                 '*********************************************************
Picture1.Height = 125               '*********************************************************
HRuler.Width = Picture1.Width       '*  set the new card dimensions, adjust the H & V rulers *
VRuler.Height = Picture1.Height     '*                                                       *
                                    '*   change the buffer pic to match the new card size    *
picBuffer.Width = Picture1.Width    '*********************************************************
picBuffer.Height = Picture1.Height  '*********************************************************


Command12.Visible = False ' set card to portrait button hidden
Command13.Visible = True  ' reset card to landscape button visible

' show the size of the label
Label10.Caption = "Current Card Size: " & Round((Picture1.Width), 0) & " mm" & " X " & Round((Picture1.Height), 0) & " mm"

Call ssReset ' reset image & label positions

Command15(3).Enabled = True    '************************************************
Command15(4).Enabled = True    '************************************************
Command15(5).Enabled = True    '*                                              *
Command15(9).Enabled = True    '*                                              *
Command15(19).Enabled = True   '*                                              *
Command15(18).Enabled = True   '*                                              *
Command15(16).Enabled = True   '*                                              *
Command15(12).Enabled = True   '*                                              *
Command15(13).Enabled = True   '*                                              *
Command15(10).Enabled = True   '*                                              *
Command15(22).Enabled = True   '*                                              *
                               '*  enable the preselected  background buttons  *
Command15(0).Enabled = False   '*                                              *
Command15(1).Enabled = False   '*                                              *
Command15(2).Enabled = False   '*                                              *
Command15(6).Enabled = False   '*                                              *
Command15(7).Enabled = False   '*                                              *
Command15(8).Enabled = False   '*                                              *
Command15(11).Enabled = False  '*                                              *
Command15(14).Enabled = False  '*                                              *
Command15(15).Enabled = False  '*                                              *
Command15(17).Enabled = False  '*                                              *
Command15(20).Enabled = False  '************************************************
Command15(21).Enabled = False  '************************************************
End Sub

Private Sub Command12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command12.MousePointer = 99
Command12.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command13_Click()
'*******************************************************
'*             set to landscape orientation            *
'*******************************************************
Picture1.Width = 125                 '*********************************************************
Picture1.Height = 75                 '*********************************************************
HRuler.Width = Picture1.Width        '*  set the new card dimensions, adjust the H & V rulers *
VRuler.Height = Picture1.Height      '*                                                       *
                                     '*   change the buffer pic to match the new card size    *
picBuffer.Width = Picture1.Width     '*********************************************************
picBuffer.Height = Picture1.Height   '*********************************************************


Command13.Visible = False ' set card to portrait button visible
Command12.Visible = True  ' reset card to landscape button hidden

' show the size of the label
Label10.Caption = "Current Card Size: " & Round((Picture1.Width), 0) & " mm" & " X " & Round((Picture1.Height), 0) & " mm"

Call ssReset ' reset image & label positions

Command15(3).Enabled = False     '************************************************
Command15(4).Enabled = False     '************************************************
Command15(5).Enabled = False     '*                                              *
Command15(9).Enabled = False     '*                                              *
Command15(19).Enabled = False    '*                                              *
Command15(18).Enabled = False    '*                                              *
Command15(16).Enabled = False    '*                                              *
Command15(12).Enabled = False    '*                                              *
Command15(13).Enabled = False    '*                                              *
Command15(10).Enabled = False    '*                                              *
Command15(22).Enabled = False    '*                                              *
                                 '*  enable the preselected  background buttons  *
Command15(0).Enabled = True      '*                                              *
Command15(1).Enabled = True      '*                                              *
Command15(2).Enabled = True      '*                                              *
Command15(6).Enabled = True      '*                                              *
Command15(7).Enabled = True      '*                                              *
Command15(8).Enabled = True      '*                                              *
Command15(11).Enabled = True     '*                                              *
Command15(14).Enabled = True     '*                                              *
Command15(15).Enabled = True     '*                                              *
Command15(17).Enabled = True     '*                                              *
Command15(20).Enabled = True     '************************************************
Command15(21).Enabled = True     '************************************************
End Sub

Private Sub Command13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command13.MousePointer = 99
Command13.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

'Private Sub Command14_Click()
''*******************************************************
''*         sets a very basic vertical gradient         *
''*******************************************************
'Dim i As Integer, Y As Integer
'
'Picture1.AutoRedraw = True
'    Picture1.DrawStyle = 6
'    Picture1.DrawMode = 13
'    Picture1.DrawWidth = 13
'    Picture1.ScaleMode = 6
'   Picture1.ScaleHeight = 256
'
'    For i = 0 To 300
'        Picture1.Line (0, Y)-(Picture1.Width, Y + 1), RGB(i, Text6.Text, Text5.Text), BF
'        Y = Y + 1
'    Next i
'End Sub

Private Sub Command14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command14.MousePointer = 99
Command14.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub


Private Sub Command15_Click(Index As Integer)
'*******************************************************
'*      load the gradients from resource file          *
'*******************************************************
Select Case Index
Case 0
Picture1.Picture = LoadResPicture(101, vbResBitmap)
Case 1
Picture1.Picture = LoadResPicture(103, vbResBitmap)
Case 2
Picture1.Picture = LoadResPicture(105, vbResBitmap)
Case 3
Picture1.Picture = LoadResPicture(102, vbResBitmap)
Case 4
Picture1.Picture = LoadResPicture(104, vbResBitmap)
Case 5
Picture1.Picture = LoadResPicture(106, vbResBitmap)
Case 6
Picture1.Picture = LoadResPicture(112, vbResBitmap)
Case 7
Picture1.Picture = LoadResPicture(107, vbResBitmap)
Case 8
Picture1.Picture = LoadResPicture(108, vbResBitmap)
Case 9
Picture1.Picture = LoadResPicture(109, vbResBitmap)
Case 10
Picture1.Picture = LoadResPicture(111, vbResBitmap)
Case 11
Picture1.Picture = LoadResPicture(110, vbResBitmap)
Case 12
Picture1.Picture = LoadResPicture(113, vbResBitmap)
Case 13
Picture1.Picture = LoadResPicture(115, vbResBitmap)
Case 14
Picture1.Picture = LoadResPicture(114, vbResBitmap)
Case 15
Picture1.Picture = LoadResPicture(116, vbResBitmap)
Case 16
Picture1.Picture = LoadResPicture(117, vbResBitmap)
Case 17
Picture1.Picture = LoadResPicture(118, vbResBitmap)
Case 18
Picture1.Picture = LoadResPicture(119, vbResBitmap)
Case 19
Picture1.Picture = LoadResPicture(121, vbResBitmap)
Case 20
Picture1.Picture = LoadResPicture(120, vbResBitmap)
Case 21
Picture1.Picture = LoadResPicture(122, vbResBitmap)
Case 22
Picture1.Picture = LoadResPicture(123, vbResBitmap)
End Select
End Sub

Private Sub Command15_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command15(Index).MousePointer = 99
Command15(Index).MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command16_Click()
'**********************************************************
'*  show card printer & preview & minimise design screen  *
'**********************************************************
frmID_Card_Printer.Show

Me.WindowState = vbMinimized

End Sub

Private Sub Command16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command16.MousePointer = 99
Command16.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command17_Click()
'**********************************************************
'*          show webcam picture capture screen            *
'**********************************************************
frmID_ICMWIC.Show

Me.WindowState = vbMinimized
End Sub

Private Sub Command17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command17.MousePointer = 99
Command17.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command3.MousePointer = 99
Command3.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub cmbFont_Click()
'*******************************************************
'*            change the selected label font           *
'*******************************************************
If optLineText(0).Value = True Then
lblText(0).Font = cmbFont.Text
ElseIf optLineText(1).Value = True Then
lblText(1).Font = cmbFont.Text
ElseIf optLineText(2).Value = True Then
lblText(2).Font = cmbFont.Text
ElseIf optLineText(3).Value = True Then
lblText(3).Font = cmbFont.Text
ElseIf optLineText(4).Value = True Then
lblText(4).Font = cmbFont.Text
ElseIf optLineText(5).Value = True Then
lblText(5).Font = cmbFont.Text
ElseIf optLineText(6).Value = True Then
lblText(6).Font = cmbFont.Text
ElseIf optLineText(7).Value = True Then
lblText(7).Font = cmbFont.Text
ElseIf optLineText(8).Value = True Then
lblText(8).Font = cmbFont.Text
ElseIf optLineText(9).Value = True Then
lblText(9).Font = cmbFont.Text
ElseIf optLineText(10).Value = True Then
lblText(10).Font = cmbFont.Text
ElseIf optLineText(11).Value = True Then
lblText(11).Font = cmbFont.Text
End If
End Sub

Private Sub cmbFontSize_Click()
'*******************************************************
'*         change the selected label font size         *
'*******************************************************
If optLineText(0).Value = True Then
lblText(0).FontSize = cmbFontSize.Text
ElseIf optLineText(1).Value = True Then
lblText(1).FontSize = cmbFontSize.Text
ElseIf optLineText(2).Value = True Then
lblText(2).FontSize = cmbFontSize.Text
ElseIf optLineText(3).Value = True Then
lblText(3).FontSize = cmbFontSize.Text
ElseIf optLineText(4).Value = True Then
lblText(4).FontSize = cmbFontSize.Text
ElseIf optLineText(5).Value = True Then
lblText(5).FontSize = cmbFontSize.Text
ElseIf optLineText(6).Value = True Then
lblText(6).FontSize = cmbFontSize.Text
ElseIf optLineText(7).Value = True Then
lblText(7).FontSize = cmbFontSize.Text
ElseIf optLineText(8).Value = True Then
lblText(8).FontSize = cmbFontSize.Text
ElseIf optLineText(9).Value = True Then
lblText(9).FontSize = cmbFontSize.Text
ElseIf optLineText(10).Value = True Then
lblText(10).FontSize = cmbFontSize.Text
ElseIf optLineText(11).Value = True Then
lblText(11).FontSize = cmbFontSize.Text
End If
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*           set the command button cursor             *
'*              make new folder button                 *
'*******************************************************
Command1.MousePointer = 99
Command1.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command2_Click()
'*******************************************************
'*                  exit application                   *
'*******************************************************
Unload Me
Unload frmID_Preview
Unload frmID_Card_Printer
Unload frmID_Help
Unload frmID_ICMWIC
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*           set the command button cursor             *
'*******************************************************
Command2.MousePointer = 99
Command2.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command3_Click()
'*******************************************************
'*    convert the selected label text to uppercase     *
'*******************************************************
If optLineText(0).Value = True Then
lblText(0).Caption = UCase$(lblText(0).Caption)
End If
If optLineText(1).Value = True Then
lblText(1).Caption = UCase$(lblText(1).Caption)
End If
If optLineText(2).Value = True Then
lblText(2).Caption = UCase$(lblText(2).Caption)
End If
If optLineText(3).Value = True Then
lblText(3).Caption = UCase$(lblText(3).Caption)
End If
If optLineText(4).Value = True Then
lblText(4).Caption = UCase$(lblText(4).Caption)
End If
If optLineText(5).Value = True Then
lblText(5).Caption = UCase$(lblText(5).Caption)
End If
If optLineText(6).Value = True Then
lblText(6).Caption = UCase$(lblText(6).Caption)
End If
If optLineText(7).Value = True Then
lblText(7).Caption = UCase$(lblText(7).Caption)
End If
If optLineText(8).Value = True Then
lblText(8).Caption = UCase$(lblText(8).Caption)
End If
If optLineText(9).Value = True Then
lblText(9).Caption = UCase$(lblText(9).Caption)
End If
If optLineText(10).Value = True Then
lblText(10).Caption = UCase$(lblText(10).Caption)
End If
If optLineText(11).Value = True Then
lblText(11).Caption = UCase$(lblText(11).Caption)
End If
End Sub

Private Sub Command4_Click()
'*******************************************************
'*    convert the selected label text to lowercase     *
'*******************************************************
If optLineText(0).Value = True Then
lblText(0).Caption = LCase$(lblText(0).Caption)
End If
If optLineText(1).Value = True Then
lblText(1).Caption = LCase$(lblText(1).Caption)
End If
If optLineText(2).Value = True Then
lblText(2).Caption = LCase$(lblText(2).Caption)
End If
If optLineText(3).Value = True Then
lblText(3).Caption = LCase$(lblText(3).Caption)
End If
If optLineText(4).Value = True Then
lblText(4).Caption = LCase$(lblText(4).Caption)
End If
If optLineText(5).Value = True Then
lblText(5).Caption = LCase$(lblText(5).Caption)
End If
If optLineText(6).Value = True Then
lblText(6).Caption = LCase$(lblText(6).Caption)
End If
If optLineText(7).Value = True Then
lblText(7).Caption = LCase$(lblText(0).Caption)
End If
If optLineText(8).Value = True Then
lblText(8).Caption = LCase$(lblText(8).Caption)
End If
If optLineText(9).Value = True Then
lblText(9).Caption = LCase$(lblText(9).Caption)
End If
If optLineText(10).Value = True Then
lblText(10).Caption = LCase$(lblText(10).Caption)
End If
If optLineText(11).Value = True Then
lblText(11).Caption = LCase$(lblText(11).Caption)
End If
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command4.MousePointer = 99
Command4.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command5_Click()
'*******************************************************
'*               clear all command button              *
'*******************************************************
Dim i As Integer
For i = 0 To 11
lblText(i).Caption = ""       '* clear   hide    labels
lblText(i).Visible = False    '*      and    text

lblText(i).Font = "MS Sans Serif"   '* reset
lblText(i).FontSize = 8             '*      the
lblText(i).FontBold = False         '*         label
lblText(i).FontItalic = False       '*              default
lblText(i).FontUnderline = False    '*                     values
lblText(i).FontStrikethru = False   '*                           and
lblText(i).ForeColor = vbBlack      '*                              properties

lblText(i).Left = 240 '* reset   label
lblText(i).Top = 2640 '*      the     placement

optLineText(i).Value = False  '* reset line selectors
Next i
Picture1.Picture = picBuffer.Picture '* clear any background image
Picture1.BackColor = &H8000000F '* default background

Text2.Text = "Label Text" '* reset textbox

Image1.Picture = imgBuffer.Picture '* clear image
Image1.Top = 360
Image1.Left = 360      '*      the            placement
Image1.Width = 1425    '*         image    and
Image1.Height = 1665   '*              size


cmbFont.Text = "MS Sans Serif" '* text font
cmbFontSize.Text = "8"         '*          text size
cmbTxtColour.Text = "Black"    '*                   text colour

Check2.Value = 0 '* bold
Check3.Value = 0 '* italic
Check4.Value = 0 '* underline
Check5.Value = 0 '* strikethru
Check6.Value = 0 '* reset lock edit
Check7.Value = 0 '* set background image option

HScroll1.Value = 0
HScroll2.Value = 0

HideNodes

End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command5.MousePointer = 99
Command5.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command6_Click()
'*******************************************************
'*          show the card preview window               *
'*        minimises the main design window             *
'*******************************************************
Dim rv As Long

frmID_Preview.Show
frmID_Preview.Label9.Caption = Label9.Caption


Me.WindowState = vbMinimized
         
'* this came straight from ms technet!
frmID_Preview.Picture2.AutoRedraw = True
rv = SendMessage(Picture1.hwnd, WM_PAINT, frmID_Preview.Picture2.hdc, 0)
rv = SendMessage(Picture1.hwnd, WM_PRINT, frmID_Preview.Picture2.hdc, PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)
frmID_Preview.Picture2.Picture = frmID_Preview.Picture2.image
frmID_Preview.Picture2.AutoRedraw = False


End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command6.MousePointer = 99
Command6.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command7_Click()
'*******************************************************
'*                     mayday mayday                   *
'********************************************************
frmID_Help.Show
frmID_Help.Command1.Enabled = True
End Sub

Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command7.MousePointer = 99
Command7.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command8_Click()
'*******************************************************
'*                          about                      *
'*******************************************************
frmLink.Show
End Sub

Private Sub Command8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command8.MousePointer = 99
Command8.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command9_Click()
'**********************************************
'*         create folder, extract dll         *
'*            save disable button             *
'**********************************************
MkDir App.Path & "\Template\"
Call inJpgConvert
Text4(0).Text = "Yes"
Call Saverz
Command9.Enabled = False
End Sub
Private Sub inJpgConvert()
'**********************************************
'*   extract the dll file from the resource   *
'**********************************************
On Error Resume Next
Dim ff As Integer
ff = FreeFile
  MkDir App.Path & "\"
   On Error GoTo 0
  Dim bytResourceData()   As Byte
  bytResourceData = LoadResData(101, "CUSTOM")
   Open App.Path & "\ijl11.dll" For Binary Shared As #ff
  Put #ff, 1, bytResourceData
  Close
End Sub

Private Sub Command9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command9.MousePointer = 99
Command9.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

 Private Sub Dir1_Change()
 File1.Path = Dir1.Path
 End Sub

 Private Sub Drive1_Change()
 '*******************************************************
'*                 first select your drive              *
'********************************************************
 On Error GoTo fubar
  Dim msg As String
  
 Dir1.Path = Drive1.Drive
 
fubar:
      If (Err.Number = rterror) Then
        msg = "DRIVE NOT READY TRY ANOTHER"
        If MsgBox(msg) = vbOK Then
          frmID_Card.SetFocus
        End If
      End If
      Exit Sub
 End Sub
 Private Sub File1_Click()
'********************************************************
'*         opens the file into the image box            *
'*            or as the background image                *
'********************************************************
If Check7.Value = 1 Then
Picture1.Picture = LoadPicture(File1.Path & "\" & File1.FileName) '* background image
Else
Image1.Picture = LoadPicture(File1.Path & "\" & File1.FileName)   '* image box image
End If
 End Sub

Private Sub HScroll1_Scroll()
'*******************************************************
'*              set the scrolling cursor               *
'*******************************************************
HScroll1.MousePointer = 99
HScroll1.MouseIcon = LoadResPicture(106, vbResCursor)
End Sub

Private Sub HScroll2_Scroll()
'*******************************************************
'*              set the scrolling cursor               *
'*******************************************************
HScroll2.MousePointer = 99
HScroll2.MouseIcon = LoadResPicture(106, vbResCursor)
End Sub

Private Sub Image1_DblClick()
'*******************************************************
'*             hide the grab handles(nodes)            *
'*******************************************************
 For i = 0 To 7
        Node(i).Visible = False
    Next
End Sub

Private Sub HideNodes()
'*******************************************************
'*             hide the grab handles(nodes)            *
'*******************************************************
    For i = 0 To 7
        Node(i).Visible = False
    Next
End Sub
Private Sub SizeObject(NodeIndex As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    With Image1
        Select Case NodeIndex
            Case 0
                .Width = .Width + (.Left - X)
                .Height = .Height + (.Top - Y)
                .Left = X
                .Top = Y
            Case 1
                .Width = .Width + (.Left - X)
                .Left = X
            Case 2
                .Width = (.Left - X) + .Width
                .Height = Y - .Top
                .Left = X
            Case 3
                .Height = .Height + .Top - Y
                .Top = Y
            Case 4
                .Height = Y - .Top
            Case 5
                .Width = X - .Left
                .Height = .Height + .Top - Y
                .Top = Y
            Case 6
                .Width = X - .Left
            Case 7
                .Width = X - .Left
                .Height = Y - .Top
        End Select
    End With
    'KeyEdit = "Move"
End Sub
Private Sub SetNodes(SelectedControl As Control)
    With SelectedControl
        For i = 0 To 7
            Select Case i
                'Left Top
                Case 0
                    Node(i).Left = .Left - Node(i).Width
                    Node(i).Top = .Top - Node(i).Height
                'Left center
                Case 1
                    Node(i).Left = .Left - Node(i).Width
                    Node(i).Top = .Top + ((.Height - Node(i).Height) / 2)
                'Left bottom
                Case 2
                    Node(i).Left = .Left - Node(i).Width
                    Node(i).Top = .Top + .Height
                'Center Top
                Case 3
                    Node(i).Left = .Left + ((.Width + Node(i).Width) / 2)
                    Node(i).Top = .Top - Node(i).Height
                'Center Bottom
                Case 4
                    Node(i).Left = .Left + ((.Width + Node(i).Width) / 2)
                    Node(i).Top = .Top + .Height
                 'Right Top
                Case 5
                    Node(i).Left = .Left + .Width
                    Node(i).Top = .Top - Node(i).Height
                'Right Center
                Case 6
                    Node(i).Left = .Left + .Width
                    Node(i).Top = .Top + ((.Height - Node(i).Height) / 2)
                'Right Bottom
                Case 7
                    Node(i).Left = .Left + .Width
                    Node(i).Top = .Top + .Height
            End Select
            Node(i).Visible = True
        Next
    End With
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*               set the select drag cursor            *
'*******************************************************
CurX = X
CurY = Y
Image1.MousePointer = 99
Image1.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*                set the default cursor               *
'*******************************************************
Image1.MousePointer = 0
If Button = 2 Then
  HideNodes
End If
End Sub

Private Sub lblText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*               set the select drag cursor            *
'*******************************************************
CurX = X
CurY = Y
lblText(Index).MousePointer = 99
lblText(Index).MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub lblText_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*        set cursor & drag the labels around          *
'*******************************************************
If Button = 2 Then
lblText(Index).Move lblText(Index).Left + (X - CurX), lblText(Index).Top + (Y - CurY)
lblText(Index).MousePointer = 99
lblText(Index).MouseIcon = LoadResPicture(101, vbResCursor)
End If
End Sub

Private Sub lblText_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*                set the default cursor               *
'*******************************************************
lblText(Index).MousePointer = 0
End Sub

Private Sub optLineText_Click(Index As Integer)
'*******************************************************
'*        reset the font options for each label        *
'*******************************************************
If Check2.Value = 1 Then ' bold
Check2.Value = 0
End If
If Check3.Value = 1 Then ' italic
Check3.Value = 0
End If
If Check4.Value = 1 Then ' underline
Check4.Value = 0
End If
If Check5.Value = 1 Then ' strikethru
Check5.Value = 0
End If
cmbFont.Text = "MS Sans Serif" '* text font
cmbFontSize.Text = "8"         '*          text size
cmbTxtColour.Text = "Black"    '*                   text colour

If lblText(Index).Caption <> "Label1" Then
cmbFont.Text = lblText(Index).FontName        '*************************************************
cmbFontSize.Text = lblText(Index).FontSize    '*************************************************
Text2.Text = lblText(Index).Caption           '*                                              **
If lblText(Index).FontBold = True Then        '*                                              **
Check2.Value = 1                              '*                                              **
End If                                        '*       a refinement of the label editing      **
If lblText(Index).FontItalic = True Then      '*                                              **
Check3.Value = 1                              '*      displays the label text and formatting  **
End If                                        '*                                              **
If lblText(Index).FontUnderline = True Then   '*         when the opt button is selected      **
Check4.Value = 1                              '*                                              **
End If                                        '*      all changes noted except label colour   **
If lblText(Index).FontStrikethru = True Then  '*                                              **
Check5.Value = 1                              '*                                              **
End If                                        '*                                              **
Else: Text2.Text = ""                         '*************************************************
End If                                        '*************************************************
End Sub

Private Sub optLineText_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*       set the select line radio button cursor       *
'*******************************************************
optLineText(Index).MousePointer = 99
optLineText(Index).MouseIcon = LoadResPicture(104, vbResCursor)
 
For Index = 0 To 11
optLineText(Index).ToolTipText = "SELECT AND MAKE LINE LABEL VISIBLE"
Next Index
End Sub

Private Sub picBG_Col_Click(Index As Integer)
'*******************************************************
'*        select the id card background colour         *
'*******************************************************
Picture1.BackColor = picBG_Col(Index).BackColor
End Sub

Private Sub picBG_Col_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*      set the dropper cursor, call the rgb sub       *
'*******************************************************

picBG_Col(Index).MousePointer = 99
picBG_Col(Index).MouseIcon = LoadResPicture(102, vbResCursor)

colours = picBG_Col(Index).BackColor
Call sVColour
Label2.Caption = "R:" & reds & " G:" & greens & " B:" & blues

End Sub

Private Sub Picture1_DragDrop(Source As Control, X As Single, Y As Single)
    SetNodes Image1
End Sub

Private Sub Picture1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    SizeObject Source.Index, X, Y
End Sub

Private Sub Form_Load()
'********************************************
'*       intialise nodes(grab handles)      *
'*            image file pattern            *
'*     enumerate fonts & add to combo box   *
'********************************************
Dim G As String
   Dim F   ' Declare variable.
   For F = 0 To Printer.FontCount - 1  ' Determine number of fonts.
      cmbFont.AddItem Printer.Fonts(F)    ' Put each font into list box.
   Next F
   G = cmbFont.ListCount
   Text3.Text = G
   
Label4.Caption = App.Path

File1.Pattern = Label3.Caption ' *.jpg - jpeg files as used by digital cameras
    If Node.count = 1 Then
        For i = 1 To 7
            Load Node(i)
        Next
    End If
    
 Call Openerz 'load the convos.jcf file
    
If Text4(0).Text = "No" Then       '************************************************
Command9.Enabled = True            '*  just a holder for the convos.jcf file text  *
ElseIf Text4(0).Text = "Yes" Then  '*     the check file for the dll & folder      *
Command9.Enabled = False           '*                installation                  *
End If                             '************************************************


' show the size of the label
Label10.Caption = "Current Card Size: " & Round((Picture1.Width), 0) & " mm" & " X " & Round((Picture1.Height), 0) & " mm"

Command15(3).Enabled = False
Command15(4).Enabled = False
Command15(5).Enabled = False
Command15(9).Enabled = False
Command15(19).Enabled = False
Command15(18).Enabled = False
Command15(16).Enabled = False
Command15(12).Enabled = False
Command15(13).Enabled = False
Command15(10).Enabled = False
Command15(22).Enabled = False

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 '****************************************************
'*     follow the cursor in the H & V ruler pics     *
'*****************************************************
   Me.MousePointer = 0
    HRuler.DrawMode = 6
    HRuler.Line (X, 0)-(X, HRuler.ScaleHeight)
    If X1 > 0 Then
        HRuler.Line (X1, 0)-(X1, HRuler.ScaleHeight)
    End If
    HRuler.DrawMode = 13
    X1 = X
    VRuler.DrawMode = 6
        VRuler.Line (0, Y)-(VRuler.ScaleWidth, Y)
    If Y1 > 0 Then
        VRuler.Line (0, Y1)-(VRuler.ScaleWidth, Y1)
    End If
       VRuler.DrawMode = 13
    Y1 = Y
End Sub

Private Sub Image1_Click()
'****************************************************
'*  show the nodes or grab handles to resize image  *
'****************************************************
    If Node.count = 1 Then
        For i = 1 To 7
            Load Node(i)
        Next
    End If
    SetNodes Image1

End Sub

Private Sub Image1_DragDrop(Source As Control, X As Single, Y As Single)
    SetNodes Image1
End Sub

Private Sub Image1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'****************************************************
'*         SizeObject Source.Index, X, Y            *
'****************************************************
    Me.MousePointer = 0
    With Image1
        Select Case Source.Index
            Case 0
                .Top = .Top + Y
                .Left = .Left + X
                .Width = .Width - X
                .Height = .Height - Y
            Case 1
                .Left = .Left + X
                .Width = .Width - X
            Case 2
                .Width = .Width - X
                .Height = Y
                .Left = .Left + X
            Case 3
                .Height = .Height - Y
                .Top = .Top + Y
            Case 4
                .Height = Y
            Case 5
                .Width = X
                .Height = .Height - Y
                .Top = .Top + Y
            Case 6
                .Width = X
            Case 7
                .Width = X
                .Height = Y
        End Select
    End With
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*       set cursor & drag the picture around          *
'*******************************************************
If Button = 2 Then
Image1.MousePointer = 99
Image1.MouseIcon = LoadResPicture(101, vbResCursor)
Image1.Move Image1.Left + (X - CurX), Image1.Top + (Y - CurY)
  HideNodes
End If
End Sub

Private Sub node_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   HideNodes
   Node(Index).Drag
End Sub

Private Sub node_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*  sets the cursor used to drag the grab handle(node) *
'*******************************************************
    Select Case Index
        Case 0, 7
            MousePointer = 8
        Case 1, 6
            MousePointer = 9
        Case 2, 5
            MousePointer = 6
        Case 3, 4
            MousePointer = 7
    End Select
End Sub
Private Sub sVColour()
'*******************************************************
'*        split the colour into its rgb parts          *
'*******************************************************
blues = Int(colours / 65536)
greens = Int((colours - (65536 * blues)) / 256)
reds = colours - (blues * 65536) - (greens * 256)
End Sub

Private Sub Text2_Change()
'*******************************************************
'*      select the label in which to insert text       *
'*******************************************************
If optLineText(0).Value = True Then
lblText(0).Visible = True
lblText(0).Caption = Text2.Text
ElseIf optLineText(1).Value = True Then
lblText(1).Visible = True
lblText(1).Caption = Text2.Text
ElseIf optLineText(2).Value = True Then
lblText(2).Visible = True
lblText(2).Caption = Text2.Text
ElseIf optLineText(3).Value = True Then
lblText(3).Visible = True
lblText(3).Caption = Text2.Text
ElseIf optLineText(4).Value = True Then
lblText(4).Visible = True
lblText(4).Caption = Text2.Text
ElseIf optLineText(5).Value = True Then
lblText(5).Visible = True
lblText(5).Caption = Text2.Text
ElseIf optLineText(6).Value = True Then
lblText(6).Visible = True
lblText(6).Caption = Text2.Text
ElseIf optLineText(7).Value = True Then
lblText(7).Visible = True
lblText(7).Caption = Text2.Text
ElseIf optLineText(8).Value = True Then
lblText(8).Visible = True
lblText(8).Caption = Text2.Text
ElseIf optLineText(9).Value = True Then
lblText(9).Visible = True
lblText(9).Caption = Text2.Text
ElseIf optLineText(10).Value = True Then
lblText(10).Visible = True
lblText(10).Caption = Text2.Text
ElseIf optLineText(11).Value = True Then
lblText(11).Visible = True
lblText(11).Caption = Text2.Text
End If
End Sub
Private Sub Saverz()
'*******************************************************
'*           saves the disable button file             *
'*******************************************************
On Error GoTo snafufubar
 
  Dim msg As String
  Dim Filehandle As Integer
  Dim X As Integer

  Filehandle = FreeFile

   Open App.Path & "\convos.jcf" For Output As Filehandle
        
          Write #Filehandle, frmID_Card.Text4(0);

      
      Close #Filehandle

snafufubar:
      If (Err.Number = myerrfilepath) Then
        msg = "you must save a file"
        If MsgBox(msg) = vbOK Then
          frmID_Card.SetFocus
        End If
      End If
      Exit Sub
End Sub
Private Sub Openerz()
'*******************************************************
'*           opens the disable button file             *
'*******************************************************
On Error GoTo fubar
  
  Dim msg As String
  Dim box1 As String

  Dim Filenumber As Integer

  Filenumber = FreeFile

    Open App.Path & "\convos.jcf" For Input As #Filenumber

        Do While Not EOF(Filenumber)
          Input #Filenumber, box1
          Text4(0).Text = box1
                Loop
      Close #Filenumber

      Exit Sub
fubar:
      If (Err.Number = errfilepath) Then
        msg = "you must select a file to open"
        If MsgBox(msg) = vbOK Then
          frmID_Card.SetFocus
        End If
      End If
      Exit Sub
End Sub
Private Sub ssReset()
'*******************************************************
'*        reset the image and label postions           *
'*          when changing card orientation             *
'*******************************************************
Dim i As Integer

For i = 0 To 11
lblText(i).Left = 240 '* reset   label
lblText(i).Top = 2640 '*      the     placement
Next i

Image1.Top = 360
Image1.Left = 360      '*      the            placement
Image1.Width = 1425    '*         image    and
Image1.Height = 1665   '*              size

End Sub

'Private Sub HScroll1_Change()
''*******************************************************
''*          set the hand pointing cursor               *
''*******************************************************
'HScroll1.MousePointer = 99
'HScroll1.MouseIcon = LoadResPicture(103, vbResCursor)
''*******************************************************
''*             changes the top value colur             *
''*******************************************************
'Text6.Text = HScroll1.Value
'End Sub
'
'Private Sub HScroll2_Change()
''*******************************************************
''*          set the hand pointing cursor               *
''*******************************************************
'HScroll2.MousePointer = 99
'HScroll2.MouseIcon = LoadResPicture(103, vbResCursor)
''*******************************************************
''*           changes the lower value colur             *
''*******************************************************
'Text5.Text = HScroll2.Value
'End Sub
'
'Private Sub Text5_Change()
''*******************************************************
''*             scrolls the top value colur             *
''*******************************************************
'HScroll2.Value = Text5.Text
'End Sub
'
'Private Sub Text6_Change()
''*******************************************************
''*           scrolls the lower value colur             *
''*******************************************************
'HScroll1.Value = Text6.Text
'End Sub
