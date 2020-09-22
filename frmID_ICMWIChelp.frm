VERSION 5.00
Begin VB.Form frmID_ICMWIChelp 
   Caption         =   "Form1"
   ClientHeight    =   6975
   ClientLeft      =   210
   ClientTop       =   1650
   ClientWidth     =   13410
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   13410
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   12960
      Picture         =   "frmID_ICMWIChelp.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "EXIT THIS SCREEN:- BY THE NEAREST DOOR!"
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   525
      Index           =   10
      Left            =   8145
      TabIndex        =   12
      Top             =   6210
      Width           =   540
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   540
      Index           =   9
      Left            =   7440
      TabIndex        =   11
      Top             =   6195
      Width           =   555
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   525
      Index           =   8
      Left            =   6060
      TabIndex        =   10
      Top             =   6015
      Width           =   540
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   525
      Index           =   7
      Left            =   5445
      TabIndex        =   9
      Top             =   6030
      Width           =   555
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   525
      Index           =   6
      Left            =   4575
      TabIndex        =   8
      Top             =   6030
      Width           =   540
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   525
      Index           =   5
      Left            =   3990
      TabIndex        =   7
      Top             =   6030
      Width           =   510
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   555
      Index           =   4
      Left            =   1260
      TabIndex        =   6
      Top             =   6000
      Width           =   540
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   540
      Index           =   3
      Left            =   495
      TabIndex        =   5
      Top             =   6015
      Width           =   555
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   3060
      TabIndex        =   4
      Top             =   5490
      Width           =   5640
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   5055
      Index           =   1
      Left            =   4470
      TabIndex        =   3
      Top             =   375
      Width           =   4215
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
      Height          =   6255
      Left            =   8880
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   5055
      Index           =   0
      Left            =   225
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   6735
      Left            =   120
      Picture         =   "frmID_ICMWIChelp.frx":058A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "frmID_ICMWIChelp"
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
Label1.Caption = "LIVE PREVIEW" & vbCrLf & vbCrLf & "Allows you to line up the persons face in order that you can capture their picture for the ID card. Using the right mouse button you can drag the preview image to line it up as you wish."
Case 1
Label1.Caption = "CAPTURE PREVIEW" & vbCrLf & vbCrLf & "Shows you the captured face shot before you save it."
Case 2
Label1.Caption = "SAVE LOCATION AND NAME" & vbCrLf & vbCrLf & "Saves the picture to the folder with the file name that you choose and type in. If the folder does not exist then it will be created e.g on first use of the program."
Case 3
Label1.Caption = "WEBCAM SETTINGS" & vbCrLf & vbCrLf & "Calls your webcam settings dialogue. Only enabled when preview is active"
Case 4
Label1.Caption = "WEBCAM SETTINGS 2" & vbCrLf & vbCrLf & "Calls your webcam picture settings and allows you to choose which cam to use if you have more than 1. Only enabled when preview is active"
Case 5
Label1.Caption = "START LIVE PREVIEW" & vbCrLf & vbCrLf & "Opens the webcam for you to compose the ID card picture and allows you to adjust the webcam settings."
Case 6
Label1.Caption = "STOP LIVE PREVIEW" & vbCrLf & vbCrLf & "Effectively takes the ID card picture and allows it to be captured prior to being saved."
Case 7
Label1.Caption = "CAPTURE PICTURE" & vbCrLf & vbCrLf & "Captures the picture and shows it in the window nect to the preview. The picture must be captured before it can be saved."
Case 8
Label1.Caption = "SAVE PICTURE" & vbCrLf & vbCrLf & "Saves the captured picture in the JPG format so that it can be used in the ID card main design screen."
Case 9
Label1.Caption = "SHOW THE HELP MENU"
Case 10
Label1.Caption = "EXIT THE WEBCAM CAPTURE SCREEN AND RETURN TO THE MAIN DESIGN SCREEN"
End Select
End Sub
