VERSION 5.00
Begin VB.Form frmID_Help 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ID Card Maker help"
   ClientHeight    =   615
   ClientLeft      =   210
   ClientTop       =   615
   ClientWidth     =   7995
   Icon            =   "frmID_Help.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   7995
   Begin VB.CommandButton Command5 
      Caption         =   "WEBCAM CAPTURE"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      ToolTipText     =   "SHOW PRINT PREVIEW SCREEN HELP"
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   7560
      Picture         =   "frmID_Help.frx":1002
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "EXIT THIS SCREEN:- BY THE NEAREST DOOR!"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PRINT PREVIEW"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      ToolTipText     =   "SHOW PRINT PREVIEW SCREEN HELP"
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CARD PREVIEW"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      ToolTipText     =   "SHOW CARD PREVIEW SCREEN HELP"
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MAIN SCREEN"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "SHOW MAIN SCREEN HELP"
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmID_Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'*******************************************************
'*                show the help screen                 *
'*******************************************************
frmMainHelp.Show
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*           set the command button cursor             *
'*******************************************************
Command1.MousePointer = 99
Command1.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command2_Click()
'*******************************************************
'*                show the help screen                 *
'*******************************************************
frmCardPreviewHelp.Show
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
'*                show the help screen                 *
'*******************************************************
frmPrintHelp.Show
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*           set the command button cursor             *
'*******************************************************
Command3.MousePointer = 99
Command3.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command4_Click()
'*******************************************************
'*                  exit application                   *
'*******************************************************
Unload Me
Unload frmCardPreviewHelp
Unload frmMainHelp
Unload frmPrintHelp
Unload frmID_ICMWIChelp
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*           set the command button cursor             *
'*******************************************************
Command4.MousePointer = 99
Command4.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command5_Click()
'*******************************************************
'*                show the help screen                 *
'*******************************************************
frmID_ICMWIChelp.Show
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*           set the command button cursor             *
'*******************************************************
Command5.MousePointer = 99
Command5.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub
