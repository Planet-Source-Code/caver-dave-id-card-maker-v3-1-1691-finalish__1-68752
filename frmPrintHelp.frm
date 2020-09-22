VERSION 5.00
Begin VB.Form frmPrintHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " PRINT PREVIEW HELP"
   ClientHeight    =   7605
   ClientLeft      =   225
   ClientTop       =   1680
   ClientWidth     =   12945
   Icon            =   "frmPrintHelp.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   12945
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   12480
      Picture         =   "frmPrintHelp.frx":1002
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "EXIT THIS SCREEN:- BY THE NEAREST DOOR!"
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   8
      Left            =   8400
      TabIndex        =   10
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   7
      Left            =   8400
      TabIndex        =   9
      Top             =   6240
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   6
      Left            =   7800
      TabIndex        =   8
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   5
      Left            =   5400
      TabIndex        =   7
      Top             =   6120
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   4
      Left            =   4080
      TabIndex        =   6
      Top             =   6600
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   3
      Left            =   4680
      TabIndex        =   5
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   2
      Left            =   4080
      TabIndex        =   4
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   1815
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   5520
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   5055
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   8655
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
      Height          =   6855
      Left            =   9120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   7335
      Left            =   120
      Picture         =   "frmPrintHelp.frx":158C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   8895
   End
End
Attribute VB_Name = "frmPrintHelp"
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
Label1.Caption = "PRINT PREVIEW" & vbCrLf & vbCrLf & "Shows the layout of the page to be printed" & vbCrLf & "Either 3 designs the same, 3 different designs or 2 the same and 1 different"
Case 1
Label1.Caption = "THE FILE SELECTOR -" & vbCrLf & "SELECT THE DRIVE - " & vbCrLf & "SELECT THE FOLDER" & vbCrLf & "THEN CLICK ON THE IMAGE THAT YOU TO INSERT EITHER AS THE ID PICTURE OR BACKGROUND IMAGE"
Case 2
Label1.Caption = "DOC 1" & vbCrLf & vbCrLf & "Displays out the selected card image into the top place holder"
Case 3
Label1.Caption = "DOC 2" & vbCrLf & vbCrLf & "Displays out the selected card image into the middle place holder"
Case 4
Label1.Caption = "DOC 3" & vbCrLf & vbCrLf & "Displays out the selected card image into the bottom place holder"
Case 5
Label1.Caption = "DOC 1,2,3" & vbCrLf & vbCrLf & "Select this option and then select the card image and the program will display the selected card image 3 times"
Case 6
Label1.Caption = "PRINT" & vbCrLf & vbCrLf & "Print the currently previewed design on the system default printer"
Case 7
Label1.Caption = "SHOW HELP MENU" & vbCrLf & vbCrLf & "YOU'RE HERE SO YOU WORKED IT OUT"
Case 8
Label1.Caption = "EXIT PRINT PREVEIW" & vbCrLf & vbCrLf & "Exit the print and print preview screen and return to the main design screen"
End Select
End Sub



