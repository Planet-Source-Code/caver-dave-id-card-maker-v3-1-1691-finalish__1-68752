VERSION 5.00
Begin VB.Form frmID_ICMWIC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ID CARD MAKER WEBCAM IMAGE CAPTURE"
   ClientHeight    =   9135
   ClientLeft      =   1890
   ClientTop       =   1215
   ClientWidth     =   11850
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   609
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   790
   Begin VB.CommandButton Command8 
      Height          =   735
      Left            =   10080
      Picture         =   "Form1.frx":1002
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "SHOW HELP"
      Top             =   8280
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "WEBCAM CAPTURE  && SAVE OPTIONS"
      Height          =   1335
      Left            =   5160
      TabIndex        =   11
      Top             =   7680
      Width           =   3975
      Begin VB.CommandButton Command2 
         Enabled         =   0   'False
         Height          =   735
         Left            =   3000
         Picture         =   "Form1.frx":1CCC
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "SAVE CAPTURED PICTURE AS A JPEG"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Enabled         =   0   'False
         Height          =   735
         Left            =   960
         Picture         =   "Form1.frx":2596
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "STOP WEBCAM "
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Height          =   735
         Left            =   120
         Picture         =   "Form1.frx":28A0
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "START WEBCAM CAPTURE"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Enabled         =   0   'False
         Height          =   735
         Left            =   2160
         Picture         =   "Form1.frx":2BAA
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "CAPTURE STOPPED WEBCAM IMAGE"
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "WEBCAM SETUP OPTIONS"
      Height          =   1335
      Left            =   240
      TabIndex        =   8
      Top             =   7680
      Width           =   2415
      Begin VB.CommandButton Command6 
         Height          =   735
         Left            =   240
         Picture         =   "Form1.frx":3474
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "WEBCAM SET UP OPTIONS"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         Height          =   735
         Left            =   1320
         Picture         =   "Form1.frx":377E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "WEBCAM IMAGE SIZE AND SELECT CAM OPTIONS"
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5040
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   7238
      Width           =   6735
   End
   Begin VB.PictureBox Picture3 
      Height          =   6975
      Left            =   6000
      ScaleHeight     =   6915
      ScaleWidth      =   5715
      TabIndex        =   6
      Top             =   120
      Width           =   5775
   End
   Begin VB.PictureBox Picture2 
      Height          =   6975
      Left            =   120
      ScaleHeight     =   6915
      ScaleWidth      =   5715
      TabIndex        =   5
      Top             =   120
      Width           =   5775
      Begin VB.Image Image1 
         Height          =   5280
         Left            =   240
         Top             =   0
         Width           =   4800
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   11040
      Picture         =   "Form1.frx":4448
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "EXIT THIS SCREEN:- BY THE NEAREST DOOR!"
      Top             =   8280
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   885
      Left            =   9600
      ScaleHeight     =   825
      ScaleWidth      =   1305
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   570
      Pattern         =   "*.bmp"
      TabIndex        =   1
      Top             =   1290
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Timer Timer1 
      Left            =   6000
      Top             =   6720
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Each Picture saved in folder ..\myPic\"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   210
      TabIndex        =   2
      Top             =   7200
      Width           =   4665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   3720
      Width           =   945
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      Picture         =   "Form1.frx":5112
      Top             =   3720
      Width           =   930
   End
End
Attribute VB_Name = "frmID_ICMWIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************
'
'   collected,Converted and Edited by :
'        Mohammed Samir Fayed
'              10/2004
'
'******************************************

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

    Private m_TimeToCapture_milliseconds As Integer
    
    Private m_Width As Long
    Private m_Height As Long
    
    Private mCapHwnd As Long
   
    Private bStopped As Boolean
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Function ReleaseCapture Lib "user32" () As Long
Private CurX As Double
Private CurY As Double

Private Const WM_PAINT = &HF
Private Const WM_PRINT = &H317
Private Const PRF_CLIENT = &H4&    ' Draw the window's client area
Private Const PRF_CHILDREN = &H10& ' Draw all visible child windows
Private Const PRF_OWNED = &H20&    ' Draw all owned windows

Private Sub Command1_Click()
Timer1.Enabled = False
    If mCapHwnd <> 0 Then StopWork
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
On Error Resume Next
DoEvents
If Dir(App.Path & "\myPic", vbDirectory) = "" Then MkDir (App.Path & "\myPic")

File1.Path = App.Path & "\myPic"
File1.Pattern = "*.jpg"
File1.Refresh


    Picture1.Picture = Picture3.Picture 'Image1.Picture
    SAVEJPEG App.Path & "\myPic\" & Text1.Text & ".jpg", 100, Me.Picture1
  DoEvents
  
Command2.Enabled = False
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command2.MousePointer = 99
Command2.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command3.MousePointer = 99
Command3.MouseIcon = LoadResPicture(103, vbResCursor)
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
'*          show the card preview window               *
'*        minimises the main design window             *
'*******************************************************
Dim rv As Long

'* this came straight from ms technet!
Picture3.AutoRedraw = True
rv = SendMessage(Picture2.hwnd, WM_PAINT, Picture3.hdc, 0)
rv = SendMessage(Picture2.hwnd, WM_PRINT, Picture3.hdc, PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)
Picture3.Picture = Picture3.image
Picture3.AutoRedraw = False

Command5.Enabled = False
Command2.Enabled = True
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command5.MousePointer = 99
Command5.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command6_Click()
On Error Resume Next
  If mCapHwnd = 0 Then Exit Sub

    Call SendMessage(mCapHwnd, WM_CAP_DLG_VIDEOSOURCE, 0, 0)
    DoEvents
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command6.MousePointer = 99
Command6.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command7_Click()
On Error Resume Next
    
    If mCapHwnd = 0 Then Exit Sub

    Call SendMessage(mCapHwnd, WM_CAP_DLG_VIDEOFORMAT, 0, 0)
    DoEvents
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
'*                      show help                      *
'*******************************************************
frmID_Help.Show
frmID_Help.Command5.Enabled = True
End Sub

Private Sub Command8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command8.MousePointer = 99
Command8.MouseIcon = LoadResPicture(103, vbResCursor)
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
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*       set cursor & drag the picture around          *
'*******************************************************
If Button = 2 Then
Image1.Move Image1.Left + (X - CurX), Image1.Top + (Y - CurY)
Image1.MousePointer = 99
Image1.MouseIcon = LoadResPicture(101, vbResCursor)
End If
End Sub
Private Sub Command3_Click()
StopWork
    Command6.Enabled = False
    Command7.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = True
    Command5.Enabled = True
End Sub

Private Sub Command4_Click()
 Start
 Command6.Enabled = True
Command7.Enabled = True
Command3.Enabled = True
Command4.Enabled = False
End Sub

Private Sub Form_Load()
On Error Resume Next
    m_TimeToCapture_milliseconds = 100
    m_Width = 352
    m_Height = 288
    bStopped = True
    mCapHwnd = 0
    
End Sub

Public Sub Start()
    On Error Resume Next
    If mCapHwnd <> 0 Then Exit Sub
    FrameNum = 0
    
    Timer1.Interval = m_TimeToCapture_milliseconds

    ' for safety, call stop, just in case we are already running
    Me.Timer1.Enabled = False

    ' setup a capture window
    mCapHwnd = capCreateCaptureWindowA("WebCap", 0, 0, 0, m_Width, m_Height, Me.hwnd, 0)
    DoEvents
    
    ' connect to the capture device
    Call SendMessage(mCapHwnd, WM_CAP_CONNECT, 0, 0)
    DoEvents
    
    Call SendMessage(mCapHwnd, WM_CAP_SET_PREVIEW, 0, 0)

    ' set the timer information
    bStopped = False
    Me.Timer1.Enabled = True
        

End Sub
    
Public Sub StopWork()
    On Error Resume Next
    ' stop the timer
    bStopped = True
    Timer1.Enabled = False

    ' disconnect from the video source
    DoEvents

    Call SendMessage(mCapHwnd, WM_CAP_DISCONNECT, 0, 0)
    mCapHwnd = 0

End Sub


Private Sub Timer1_Timer()
On Error Resume Next

    ' pause the timer
    Timer1.Enabled = False

    ' get the next frame;
    Call SendMessage(mCapHwnd, WM_CAP_GET_FRAME, 0, 0)

    ' copy the frame to the clipboard
    Call SendMessage(mCapHwnd, WM_CAP_COPY, 0, 0)

    ' For some reason, the API is not resizing the video
    ' feed to the width and height provided when the video
    ' feed was started, so we must resize the image here
    ' Image1.Stretch = True
            
    ' get from the clipboard
    Image1.Picture = Clipboard.GetData
         
         
    ' restart the timer
    DoEvents
    If Not bStopped Then
        Timer1.Enabled = True
    End If

End Sub
