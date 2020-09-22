VERSION 5.00
Begin VB.Form frmID_Preview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ID Card Preview"
   ClientHeight    =   9480
   ClientLeft      =   4440
   ClientTop       =   765
   ClientWidth     =   8355
   Icon            =   "frmID_Preview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   167.217
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   147.373
   Begin VB.CommandButton Command7 
      Height          =   735
      Left            =   7410
      Picture         =   "frmID_Preview.frx":1002
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "SHOW HELP"
      Top             =   7095
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Height          =   735
      Left            =   7410
      Picture         =   "frmID_Preview.frx":1CCC
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "GO TO PRINT PREVIEW"
      Top             =   4560
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   7560
      Picture         =   "frmID_Preview.frx":2996
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   8880
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Top             =   945
      Width           =   6135
   End
   Begin VB.TextBox txtQuality 
      Height          =   315
      Left            =   7320
      TabIndex        =   6
      Text            =   "90"
      Top             =   1680
      Width           =   915
   End
   Begin VB.CommandButton Command3 
      Enabled         =   0   'False
      Height          =   735
      Left            =   7410
      Picture         =   "frmID_Preview.frx":2F20
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "SAVE PICTURE AS A JPG TEMPLATE"
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Enabled         =   0   'False
      Height          =   735
      Left            =   7410
      Picture         =   "frmID_Preview.frx":37EA
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "SAVE CARD DESIGN"
      Top             =   6255
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   585
      Width           =   6135
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   4252
      Left            =   120
      ScaleHeight     =   4185
      ScaleWidth      =   7020
      TabIndex        =   1
      Top             =   1320
      Width           =   7087
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   7410
      Picture         =   "frmID_Preview.frx":44B4
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "EXIT THIS SCREEN:- BY THE NEAREST DOOR!"
      Top             =   7950
      Width           =   735
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Label10"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   360
      TabIndex        =   16
      Top             =   8520
      Width           =   570
   End
   Begin VB.Label Label4 
      Caption         =   $"frmID_Preview.frx":517E
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   8880
      Width           =   7965
   End
   Begin VB.Label Label3 
      Caption         =   "TEMPLATE FILE NAME"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "FILE NAME"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Image Quality"
      Height          =   255
      Left            =   7320
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblQuality 
      Caption         =   "Quality -  1 - 100.   1 is lowest quality/ highest   compression.  Use values > 50 for reasonable results."
      Height          =   2055
      Left            =   7320
      TabIndex        =   7
      Top             =   2040
      Width           =   855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label9 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "frmID_Preview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_cDib As New cDIBSection
Private Sub Command1_Click()
'*******************************************************
'*    exit card preview and restore design window      *
'*******************************************************
Unload Me
frmID_Card.WindowState = vbNormal
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command1.MousePointer = 99
Command1.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command2_Click()
'*******************************************************
'*        save file routine if no path specified       *
'*            then the default path is used            *
'*******************************************************
If Label9.Caption = "" Then
SavePicture Picture2.image, App.Path & "\" & Text1.Text & ".bmp"
Else
SavePicture Picture2.image, Label9.Caption & "\" & Text1.Text & ".bmp"
End If
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command2.MousePointer = 99
Command2.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command3_Click()
'*******************************************************
'*        save file routine if no path specified       *
'*            then the default path is used            *
'*******************************************************
Dim sI As String
Dim c As New cDIBSection
Dim i As Long

   Set c = New cDIBSection
   c.CreateFromPicture Picture2.Picture
   
   sI = App.Path & "\Template\" & Text2.Text & ".jpg"
   'If VBGetSaveFileName(sI, , , "JPEG Files (*.JPG)|*.JPG|All Files (*.*)|*.*", 1, , , "JPG", Me.hwnd) Then
      If SaveJPG(c, sI, plQuality()) Then
         ' OK!
      Else
         MsgBox "Failed to save the picture to the file: '" & sI & "'", vbExclamation
      End If
'   End If
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command3.MousePointer = 99
Command3.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command4_Click()
'*******************************************************
'*      go to the card selector and print preview      *
'*******************************************************
Unload Me
frmID_Card_Printer.Show
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command4.MousePointer = 99
Command4.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command7_Click()
'*******************************************************
'*                     mayday mayday                   *
'********************************************************
frmID_Help.Show
frmID_Help.Command2.Enabled = True
End Sub

Private Sub Form_Load()

 Picture2.Width = frmID_Card.Picture1.Width


Picture2.Height = frmID_Card.Picture1.Height
     
Label10.Caption = "Current Card Size: " & Round((Picture2.Width), 0) & " mm" & " X " & Round((Picture2.Height), 0) & " mm"
If frmID_Card.Command9.Enabled = True Then
Label3.Enabled = False
Text2.Enabled = False
Else
Label3.Enabled = True
Text2.Enabled = True
End If
End Sub

Private Sub Text1_Change()
'*******************************************************
'*     checks to see if file name has been entered     *
'*******************************************************
If Text1.Text <> "" Then
Command2.Enabled = True
Else: Command2.Enabled = False
End If
End Sub

Private Sub Text2_Change()
'*******************************************************
'*     checks to see if file name has been entered     *
'*******************************************************
If Text2.Text <> "" Then
Command3.Enabled = True
Else: Command3.Enabled = False
End If
End Sub

Private Sub txtQuality_KeyPress(KeyAscii As Integer)
   If KeyAscii = 8 Then
   ElseIf KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
   Else
      KeyAscii = 0
   End If
End Sub
Private Function plQuality() As Long
   On Error Resume Next
   plQuality = CLng(txtQuality.Text)
   If Not Err.Number = 0 Then
      txtQuality.Text = "90"
      plQuality = 90
   End If
End Function


Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command7.MousePointer = 99
Command7.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub
