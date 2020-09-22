VERSION 5.00
Begin VB.Form frmDTI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ShortCut 2 Desktop"
   ClientHeight    =   2055
   ClientLeft      =   3705
   ClientTop       =   4560
   ClientWidth     =   8730
   Icon            =   "frmDTI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   8730
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "frmDTI.frx":1002
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   13
      Top             =   1440
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "frmDTI.frx":130C
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   12
      Top             =   780
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "frmDTI.frx":1616
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "CREATE DESKTOP SHORTCUT"
      Height          =   495
      Left            =   840
      TabIndex        =   10
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox txtFileAssociation 
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Top             =   4305
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Height          =   615
      Left            =   8040
      Picture         =   "frmDTI.frx":1920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txtShortCutName 
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   105
      Width           =   5895
   End
   Begin VB.TextBox txtIconLocation 
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Top             =   825
      Width           =   5895
   End
   Begin VB.TextBox txtTargetPath 
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   465
      Width           =   5895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CREATE DESKTOP SHORTCUT"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4440
      TabIndex        =   0
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Default File Extension"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Default ShortCut Name"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Default Icon Location"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Default Target Path"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmDTI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wShell As New IWshShell_Class
Dim wShortcut As IWshShortcut_Class

Private Sub Check1_Click()
'*******************************************************
'*     fail safe in case of erroneous button click     *
'*         enables the create shortcut button          *
'*                default is disabled                  *
'*******************************************************
If Check1.Value = 1 Then
Command1.Enabled = True
Else
Command1.Enabled = False
End If
End Sub

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*              set the check button cursor            *
'*******************************************************
Check1.MousePointer = 99
Check1.MouseIcon = LoadResPicture(105, vbResCursor)
End Sub

Private Sub Command1_Click()
'*******************************************************
'*            call the windows script proc             *
'*******************************************************
Call CreateIcon
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command1.MousePointer = 99
Command1.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Command4_Click()
'*******************************************************
'*                   exit this screen                  *
'*******************************************************
Unload Me
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*            set the command button cursor            *
'*******************************************************
Command4.MousePointer = 99
Command4.MouseIcon = LoadResPicture(103, vbResCursor)
End Sub

Private Sub Form_Load()
'*******************************************************
'*              hard coded flexible paths              *
'*                 and shortcut name                   *
'*******************************************************
txtTargetPath.Text = App.Path & "\ID_Card.exe"
txtIconLocation.Text = App.Path & "\ID_Card.exe"
txtShortCutName.Text = "ID Card Maker v 1.2.1691.lnk"

End Sub

Private Sub CreateIcon()
'*******************************************************
'*   windows script object target, path & icon proc    *
'*******************************************************
   Set wShortcut = wShell.CreateShortcut(wShell.SpecialFolders.Item(0) & "\" & txtShortCutName.Text)
    wShortcut.TargetPath = txtTargetPath.Text
    wShortcut.IconLocation = txtIconLocation.Text
    wShortcut.Save
End Sub

