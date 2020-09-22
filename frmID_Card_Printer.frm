VERSION 5.00
Begin VB.Form frmID_Card_Printer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ID CARD SELECTOR & PRINTER"
   ClientHeight    =   10200
   ClientLeft      =   960
   ClientTop       =   540
   ClientWidth     =   13185
   Icon            =   "frmID_Card_Printer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10200
   ScaleWidth      =   13185
   Begin VB.CommandButton Command5 
      Height          =   735
      Left            =   11520
      Picture         =   "frmID_Card_Printer.frx":1002
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "CLEAR AND RESET ALL"
      Top             =   8520
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      Height          =   735
      Index           =   1
      Left            =   9120
      Picture         =   "frmID_Card_Printer.frx":1CCC
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "PRINT 4 CARDS IN PORTRAIT FORMAT"
      Top             =   9120
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      Height          =   735
      Index           =   0
      Left            =   9120
      Picture         =   "frmID_Card_Printer.frx":1FD6
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "PRINT 3 CARDS IN LANDSCAPE FORMAT"
      Top             =   8280
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Height          =   735
      Left            =   12360
      Picture         =   "frmID_Card_Printer.frx":22E0
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "SHOW HELP"
      Top             =   8520
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "FILE SELECTORS"
      Height          =   2655
      Left            =   120
      TabIndex        =   12
      Top             =   7440
      Width           =   5535
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   5295
      End
      Begin VB.DirListBox Dir1 
         Height          =   1665
         Left            =   120
         TabIndex        =   14
         Top             =   780
         Width           =   2895
      End
      Begin VB.FileListBox File1 
         Height          =   1845
         Left            =   3120
         TabIndex        =   13
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "*.bmp"
         Height          =   255
         Left            =   2520
         TabIndex        =   16
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   11760
      TabIndex        =   11
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Height          =   735
      Left            =   11520
      Picture         =   "frmID_Card_Printer.frx":2FAA
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "PRINT THIS PAGE"
      Top             =   9360
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "CARD PRINTING OPTIONS"
      Height          =   2175
      Left            =   5760
      TabIndex        =   5
      Top             =   7680
      Width           =   3135
      Begin VB.OptionButton Option1 
         Height          =   735
         Index           =   3
         Left            =   1080
         Picture         =   "frmID_Card_Printer.frx":3C74
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "ADD CARD TO LAST PLACE"
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Height          =   735
         Left            =   2040
         Picture         =   "frmID_Card_Printer.frx":3F7E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "PRINT ONE DESIGN PER SHEET (3 CARDS)"
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Height          =   735
         Index           =   0
         Left            =   120
         Picture         =   "frmID_Card_Printer.frx":4288
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "ADD CARD TO TOP PLACE"
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Height          =   735
         Index           =   1
         Left            =   1080
         Picture         =   "frmID_Card_Printer.frx":4592
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "ADD CARD TO MIDDLE PLACE"
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Height          =   735
         Index           =   2
         Left            =   120
         Picture         =   "frmID_Card_Printer.frx":489C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "ADD CARD TO LAST PLACE"
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   12360
      Picture         =   "frmID_Card_Printer.frx":4BA6
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "EXIT THIS SCREEN:- BY THE NEAREST DOOR!"
      Top             =   9360
      Width           =   735
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7215
      Left            =   12720
      Max             =   10340
      MouseIcon       =   "frmID_Card_Printer.frx":5870
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox picPPview1 
      BackColor       =   &H00C00000&
      Height          =   7215
      Left            =   120
      ScaleHeight     =   7155
      ScaleWidth      =   12555
      TabIndex        =   0
      Top             =   120
      Width           =   12615
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   16834
         Left            =   360
         ScaleHeight     =   16800
         ScaleWidth      =   11880
         TabIndex        =   2
         Top             =   1080
         Width           =   11904
         Begin VB.Image Image1 
            Height          =   7080
            Index           =   6
            Left            =   6360
            Top             =   8400
            Width           =   4275
         End
         Begin VB.Image Image1 
            Height          =   7080
            Index           =   5
            Left            =   1320
            Top             =   8400
            Width           =   4275
         End
         Begin VB.Image Image1 
            Height          =   7080
            Index           =   4
            Left            =   6360
            Top             =   960
            Width           =   4275
         End
         Begin VB.Image Image1 
            Height          =   7080
            Index           =   3
            Left            =   1320
            Top             =   960
            Width           =   4275
         End
         Begin VB.Image Image1 
            Height          =   4275
            Index           =   2
            Left            =   2400
            Top             =   10800
            Width           =   7080
         End
         Begin VB.Image Image1 
            Height          =   4275
            Index           =   1
            Left            =   2400
            Top             =   6000
            Width           =   7080
         End
         Begin VB.Image Image1 
            Height          =   4275
            Index           =   0
            Left            =   2400
            Top             =   1185
            Width           =   7080
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   16834
         Left            =   360
         ScaleHeight     =   16800
         ScaleWidth      =   11880
         TabIndex        =   3
         Top             =   1080
         Width           =   11904
      End
   End
   Begin VB.Image imgPBuff 
      Height          =   4275
      Left            =   1080
      Top             =   1200
      Visible         =   0   'False
      Width           =   7080
   End
   Begin VB.Image imgLBuff 
      Height          =   7080
      Left            =   2160
      Top             =   240
      Visible         =   0   'False
      Width           =   4275
   End
   Begin VB.Label Label1 
      Caption         =   "Print Layout Options"
      Height          =   495
      Left            =   9120
      TabIndex        =   21
      Top             =   7680
      Width           =   855
   End
End
Attribute VB_Name = "frmID_Card_Printer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 
         Private Const WM_PAINT = &HF
        Private Const WM_PRINT = &H317
        Private Const PRF_CLIENT = &H4&    ' Draw the window's client area
        Private Const PRF_CHILDREN = &H10& ' Draw all visible child windows
        Private Const PRF_OWNED = &H20&    ' Draw all owned windows

Private Sub Check1_Click()
'*******************************************************
'*            set the check button action              * Picture1.Picture = LoadPicture(File1.Path + "\" + File1.FileName)
'*******************************************************
If Check1.Value = 1 Then
Option1(0).Value = False
Option1(1).Value = False
Option1(2).Value = False
Option1(3).Value = False
End If
End Sub

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*             set the check button cursor             *
'*******************************************************
Check1.MousePointer = 99
Check1.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command1_Click()
'*******************************************************
'*    exit print preview and restore design window     *
'*******************************************************
Unload Me
frmID_Card.WindowState = vbNormal
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
'*                   print the cards                   *
'*******************************************************

         Dim rv As Long
       Picture1.SetFocus  ' So that the button doesn't look pressed
       Picture2.AutoRedraw = True
       rv = SendMessage(Picture1.hwnd, WM_PAINT, Picture2.hdc, 0)
       rv = SendMessage(Picture1.hwnd, WM_PRINT, Picture2.hdc, _
          PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)
       Picture2.Picture = Picture2.image
       Picture2.AutoRedraw = False
       Printer.Orientation = vbPRORPortrait   ' 1
       Printer.Print ""
       Printer.PaintPicture Picture2.Picture, 0, 0
       Printer.EndDoc
       Command2.SetFocus  ' Return Focu
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*           set the command button cursor             *
'*******************************************************
Command2.MousePointer = 99
Command2.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command5_Click()
'*******************************************************
'*               clear all command button              *
'*******************************************************
VScroll1.Value = 0

Image1(0).Picture = imgPBuff.Picture '* clear image
Image1(1).Picture = imgPBuff.Picture '* clear image
Image1(2).Picture = imgPBuff.Picture '* clear image

Image1(3).Picture = imgLBuff.Picture '* clear image
Image1(4).Picture = imgLBuff.Picture '* clear image
Image1(5).Picture = imgLBuff.Picture '* clear image
Image1(6).Picture = imgLBuff.Picture '* clear image

End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command5.MousePointer = 99
Command5.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command7_Click()
'*******************************************************
'*                     mayday mayday                   *
'********************************************************
frmID_Help.Show
frmID_Help.Command3.Enabled = True
End Sub

Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command7.MousePointer = 99
Command7.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Form_Load()
File1.Pattern = Label3.Caption ' *.bmp - bitmap files as used by this app
End Sub

Private Sub Option1_Click(Index As Integer)
'*******************************************************
'*            set the check button action              *
'*******************************************************
Select Case Index
Case 0
If Check1.Value = 1 Then
Check1.Value = 0
End If
Case 1
If Check1.Value = 1 Then
Check1.Value = 0
End If
Case 2
If Check1.Value = 1 Then
Check1.Value = 0
End If
Case 3
If Option1(3).Visible = True Then
If Check1.Value = 1 Then
Check1.Value = 0
End If
End If
End Select
End Sub

Private Sub Option1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the option button cursor             *
'*******************************************************
Option1(Index).MousePointer = 99
Option1(Index).MouseIcon = LoadResPicture(103, vbResCursor)
End Sub
Private Sub Dir1_Change()
 File1.Path = Dir1.Path
 End Sub

 Private Sub Drive1_Change()
 Dir1.Path = Drive1.Drive
 End Sub
 Private Sub File1_Click()
'********************************************************
'*         opens the file into the image box            *
'*            or as the background image                *
'********************************************************
If Option2(0).Value = True Then '* 3 up cards in landscape format
If Check1.Value = 1 Then
Image1(0).Picture = LoadPicture(File1.Path & "\" & File1.FileName) '* top image
Image1(1).Picture = LoadPicture(File1.Path & "\" & File1.FileName) '* middle image
Image1(2).Picture = LoadPicture(File1.Path & "\" & File1.FileName) '* bottom image
End If
If Option1(0).Value = True Then
Image1(0).Picture = LoadPicture(File1.Path & "\" & File1.FileName) '* top image
End If
If Option1(1).Value = True Then
Image1(1).Picture = LoadPicture(File1.Path & "\" & File1.FileName) '* middle image
End If
If Option1(2).Value = True Then
Image1(2).Picture = LoadPicture(File1.Path & "\" & File1.FileName) '* bottom image
End If
End If
'^*****^*****^*****^*****^*****^*****^*****^*****^*****^*****^*****^*****^*****^*****^*****^*****^
'v*****v*****v*****v***********v*****v*****v*****v*****v*****v*****v*****v*****v*****v*****v*****v
If Option2(1).Value = True Then '* 4 up cards in portrait format
If Check1.Value = 1 Then
Image1(3).Picture = LoadPicture(File1.Path & "\" & File1.FileName) '* top left image
Image1(4).Picture = LoadPicture(File1.Path & "\" & File1.FileName) '* top right image
Image1(5).Picture = LoadPicture(File1.Path & "\" & File1.FileName) '* bottom left image
Image1(6).Picture = LoadPicture(File1.Path & "\" & File1.FileName) '* bottom right image
End If
If Option1(0).Value = True Then
Image1(3).Picture = LoadPicture(File1.Path & "\" & File1.FileName) '* top left image
End If
If Option1(1).Value = True Then
Image1(4).Picture = LoadPicture(File1.Path & "\" & File1.FileName) '* top right image
End If
If Option1(2).Value = True Then
Image1(5).Picture = LoadPicture(File1.Path & "\" & File1.FileName) '* bottom left image
End If
If Option1(3).Value = True Then
Image1(6).Picture = LoadPicture(File1.Path & "\" & File1.FileName) '* bottom right image
End If
End If
 End Sub

Private Sub Option2_Click(Index As Integer)
'*******************************************************
'*    option buttons to adjust the printing options    *
'*******************************************************
Select Case Index
Case 0
If Option2(0).Value = True Then
Check1.Picture = LoadResPicture(108, vbResIcon)
Option1(3).Visible = False
Check1.Value = 0
Image1(3).Picture = imgLBuff.Picture '* clear image
Image1(4).Picture = imgLBuff.Picture '* clear image
Image1(5).Picture = imgLBuff.Picture '* clear image
Image1(6).Picture = imgLBuff.Picture '* clear image
End If
Case 1
If Option2(1).Value = True Then
Check1.Picture = LoadResPicture(109, vbResIcon)
Option1(3).Visible = True
Check1.Value = 0
Image1(0).Picture = imgPBuff.Picture '* clear image
Image1(1).Picture = imgPBuff.Picture '* clear image
Image1(2).Picture = imgPBuff.Picture '* clear image
End If
End Select
End Sub

Private Sub Option2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the option button cursor             *
'*******************************************************
Option2(Index).MousePointer = 99
Option2(Index).MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub VScroll1_Change()
'*******************************************************
'*          set the hand pointing cursor               *
'*******************************************************
VScroll1.MousePointer = 99
VScroll1.MouseIcon = LoadResPicture(103, vbResCursor)
'********************************************************
'*            scrolls the print preview page            *
'********************************************************
Picture1.Top = -VScroll1.Value
If VScroll1.Value = 0 Then
Picture1.Top = 1080
End If
Picture2.Top = -VScroll1.Value
If VScroll1.Value = 0 Then
Picture2.Top = 1080
End If
Text1.Text = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
'*******************************************************
'*              set the scrolling cursor               *
'*******************************************************
VScroll1.MousePointer = 99
VScroll1.MouseIcon = LoadResPicture(106, vbResCursor)
End Sub
