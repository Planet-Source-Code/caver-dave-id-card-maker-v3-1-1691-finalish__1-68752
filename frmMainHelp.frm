VERSION 5.00
Begin VB.Form frmMainHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " MAIN HELP"
   ClientHeight    =   8745
   ClientLeft      =   210
   ClientTop       =   1650
   ClientWidth     =   14040
   Icon            =   "frmMainHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   14040
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   13560
      Picture         =   "frmMainHelp.frx":1002
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "EXIT THIS SCREEN:- BY THE NEAREST DOOR!"
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   16
      Left            =   7680
      TabIndex        =   18
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   15
      Left            =   9120
      TabIndex        =   17
      Top             =   7800
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   14
      Left            =   9120
      TabIndex        =   16
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   13
      Left            =   8400
      TabIndex        =   15
      Top             =   7800
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   12
      Left            =   7680
      TabIndex        =   14
      Top             =   7800
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   11
      Left            =   6240
      TabIndex        =   13
      Top             =   7800
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   10
      Left            =   8400
      TabIndex        =   12
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   9
      Left            =   7680
      TabIndex        =   11
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   735
      Index           =   8
      Left            =   6120
      TabIndex        =   10
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   495
      Index           =   7
      Left            =   5880
      TabIndex        =   9
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   1935
      Index           =   6
      Left            =   7680
      TabIndex        =   8
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   1935
      Index           =   5
      Left            =   5880
      TabIndex        =   7
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   1935
      Index           =   4
      Left            =   5880
      TabIndex        =   6
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   3
      Left            =   8280
      TabIndex        =   5
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   3495
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   855
      Index           =   1
      Left            =   3840
      TabIndex        =   3
      Top             =   5160
      Width           =   5895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   735
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   5415
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
      Height          =   8055
      Left            =   9960
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   8655
      Left            =   0
      Picture         =   "frmMainHelp.frx":158C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9855
   End
End
Attribute VB_Name = "frmMainHelp"
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

Private Sub Form_Load()
Label1.Caption = "Overview" & vbCrLf & vbCrLf & "The program has been designed specifically for the design and production of ID Cards for a small company It needed to be very simple to use and simple to save the produced cards and also be able to save designs as templates for future use." & vbCrLf & vbCrLf & "The program is mainly designed around the jpg image format as this is a standard digital camera picture format and the was program required to use a standard that was common to almost, if not all, cameras on the market today." & vbCrLf & vbCrLf & "The designed cards are saved in the windows bmp format, this was done for ease of use at the time as this is the native image format and easily done."
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
'*                   reset the cursor                  *
'*******************************************************
Image1.MousePointer = 0
Label1.Caption = ""

Call Form_Load
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
Label1.Caption = "MAKE NEW FOLDER" & vbCrLf & vbCrLf & "The create a new folder button is ONLY enabled when the tick box has been ticked and once the folder has been created the button is disabled and the tick box is unticked - a safeguard against trying to create multiple folder with the same name"
Case 1
Label1.Caption = "TEXT DISPLAY" & vbCrLf & vbCrLf & "Up to 12 lines of text can be displayed on any one design this is partly for ease of use and partly because too much text can lead to a cluttered design." & vbCrLf & "Lines of text are displayed by using round selectors and then typing into the box below them." & vbCrLf & "If you really feel the need for more than 12 lines then you can save the first design as a template then add the other lines of text as needed!" & vbCrLf & "Right click and hold allows the user to move the text around the card area"
Case 2
Label1.Caption = "IMAGE DISPLAY" & vbCrLf & vbCrLf & "As the same format is used for both the ID and background images. To place a picture in the ID image holder select the image that you want from the file selector click on it and there you go. To place a picture in the background you must press the USE IMAGE AS BACKGROUND selector button then select the image you want from the file selector click on it and there you go"
Case 3
Label1.Caption = "CARD ORIENTATION" & vbCrLf & vbCrLf & "Sets the card as either landscape (horizontal) or portrait (vertical) format."
'"IMAGE DISPLAY 2" & vbCrLf & vbCrLf & "The ID image can be selected and sized as the user wishes." & vbCrLf & "A single click displays the resizing handles and allows the user to resize the picture according to their needs." & vbCrLf & "A double click hides the resizing handles." & vbCrLf & "Right click and hold allows the user to move the picture around the card area"
Case 4
Label1.Caption = "THE FILE SELECTOR -" & vbCrLf & "SELECT THE DRIVE - " & vbCrLf & "SELECT THE FOLDER" & vbCrLf & "THEN CLICK ON THE IMAGE THAT YOU TO INSERT EITHER AS THE ID PICTURE OR BACKGROUND IMAGE"
Case 5
Label1.Caption = "BACKGROUND COLOURS" & vbCrLf & vbCrLf & "These are selected from the ID CARD BACKGROUND palette, as the eye dropper cursor moves over the palette the RGB values are displayed below the palette. 1 click sets the card background to the selected colour. Clicking on the default colour restores the default background colour"
Case 6
Label1.Caption = "TEXT OPTIONS" & vbCrLf & vbCrLf & "These are set in FONT OPTIONS in here are:" & vbCrLf & vbCrLf & "The number of fonts(text styles)that you can use" & vbCrLf & vbCrLf & "A drop down selector for you to choose the font you want" & vbCrLf & vbCrLf & "A drop down selector for you to choose the size you want" & vbCrLf & vbCrLf & "A drop down selector for you to choose the colour you want"
Case 7
Label1.Caption = "BACKGROUND IMAGES" & vbCrLf & vbCrLf & "To place a picture in the background you must press the USE IMAGE AS BACKGROUND selector button then select the image you want from the file selector click on it and there you go."
Case 8
Label1.Caption = "EDIT LOCK" & vbCrLf & vbCrLf & "This button either enables or prevents any further editing of the card design"
Case 9
Label1.Caption = "CARD PREVIEW" & vbCrLf & vbCrLf & "Leads From the main screen to the card preview and save screen"
Case 10
Label1.Caption = "CLEAR ALL" & vbCrLf & vbCrLf & "Clears the card design and resetsthe design screen"
Case 11
Label1.Caption = "SHOW HELP MENU" & vbCrLf & vbCrLf & "YOU'RE HERE SO YOU WORKED IT OUT"
Case 12
Label1.Caption = "SHOW ABOUT" & vbCrLf & vbCrLf & "Shows the program about screen and link to website"
Case 13
Label1.Caption = "MINIMISE" & vbCrLf & vbCrLf & "Minimises the main screen to the windows taskbar"
Case 14
Label1.Caption = "INSTALL CONVERTER" & vbCrLf & vbCrLf & "Templates cannot be created without the Intel® JPEG Library dll, to install it and  create the template folder the button with the exclamation mark must be pressed. If you do not require the ability to create a template then you simply do not press the button to install it Installing the dll can be done at anytime as or when you need to create a reuseable template"
Case 15
Label1.Caption = "ALL DOORS LEAD OUT!"
Case 16
Label1.Caption = "TEXT OPTIONS 2" & vbCrLf & vbCrLf & "You can also alter the appearance of your text by using one or more of the six buttons below the drop down selectors and they are:" & vbCrLf & vbCrLf & "BOLD" & vbCrLf & "ITALIC" & vbCrLf & "UNDERLINE" & vbCrLf & " STRIKE THROUGH" & vbCrLf & "CONVERT TO UPPERCASE (CAPITALS)" & vbCrLf & "CONVERT TO LOWERCASE (NON-CAPITALS)"
End Select
End Sub
