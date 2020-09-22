VERSION 5.00
Begin VB.Form frmCardPreviewHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " CARD PREVIEW HELP["
   ClientHeight    =   8175
   ClientLeft      =   210
   ClientTop       =   1665
   ClientWidth     =   13095
   Icon            =   "frmCardPreviewHelp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   13095
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   12600
      Picture         =   "frmCardPreviewHelp.frx":1002
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "EXIT THIS SCREEN:- BY THE NEAREST DOOR!"
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   2175
      Index           =   8
      Left            =   6960
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   7
      Left            =   7200
      TabIndex        =   9
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   6
      Left            =   7200
      TabIndex        =   8
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   5
      Left            =   7200
      TabIndex        =   7
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   4
      Left            =   7200
      TabIndex        =   6
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   3
      Left            =   7200
      TabIndex        =   5
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   7815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   7815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   5775
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   4095
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   8280
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   7935
      Left            =   120
      Picture         =   "frmCardPreviewHelp.frx":158C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "frmCardPreviewHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command4_Click()
'*******************************************************
'*                  exit application                   *
'*******************************************************
Unload Me
Unload frmID_Help
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*           set the command button cursor             *
'*******************************************************
Command4.MousePointer = 99
Command4.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*                   reset the cursor                  *
'*******************************************************
Image1.MousePointer = 0
Label1.Caption = ""
End Sub

Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*                set the label cursor                 *
'*******************************************************
Label2(Index).MousePointer = 99
Label2(Index).MouseIcon = LoadResPicture(103, vbResCursor)
'*******************************************************
'*                  show help text                     *
'*******************************************************
Select Case Index
Case 0
Label1.Caption = "CARD PREVIEW WINDOW"
Case 1
Label1.Caption = "ENTER THE FILE NAME WITH WHICH TO SAVE THE CARD"
Case 2
Label1.Caption = "IF YOU HAVE INSTALLED THE JPG CONVERTER THEN YOU CAN ENTER A TEMPLATE FILE NAME"
Case 3
Label1.Caption = "OPEN THE PRINT PREVIEW SCREEN"
Case 4
Label1.Caption = "PIC TO JPG BUTTON - TO CREATE THE TEMPLATE IF YOU HAVE INSTALLED THE JPG CONVERTER "
Case 5
Label1.Caption = "SAVE FILE ENABLED WHEN YOU ENTER A FILE NAME"
Case 6
Label1.Caption = "SHOW THE HELP MENU"
Case 7
Label1.Caption = "EXIT THE CARD PREVIEW SCREEN AND RETURN TO THE MAIN DESIGN SCREEN"
Case 8
Label1.Caption = "JPG QUALITY ENTER A NUMBER BETWEEN 1 AND 100. 1 BEING THE POOREST AND 100 BEING THE BEST THE DEFAULT VALUE IS 90"
End Select
End Sub
