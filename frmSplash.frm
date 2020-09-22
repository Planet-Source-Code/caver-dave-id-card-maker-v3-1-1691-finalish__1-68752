VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2190
   ClientLeft      =   5190
   ClientTop       =   4635
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   360
      Top             =   1440
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "by caver dave"
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ID CARD MAKER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   20
      Height          =   2175
      Left            =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Label2.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
Timer1.Enabled = True
Timer1.Interval = 3500
End Sub

Private Sub Timer1_Timer()
If Timer1.Interval = 3500 Then
Timer1.Enabled = False
Unload Me
frmID_Card.Show
End If
End Sub
