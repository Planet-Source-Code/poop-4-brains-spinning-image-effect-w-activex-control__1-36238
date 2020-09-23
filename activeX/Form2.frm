VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "*\AProject1.vbp"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ActiveX Example - SpinImage"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3390
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3390
   StartUpPosition =   3  'Windows Default
   Begin SpinImage.SpinControl Spin 
      Height          =   1455
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2566
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   100
      Left            =   840
      Max             =   499
      SmallChange     =   10
      TabIndex        =   3
      Top             =   1560
      Value           =   100
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "None!"
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   2280
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CMN 
      Left            =   720
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select Image"
      Filter          =   "Image Files |*.gif*;*.jpg*;*.bmp*"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Image:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CMN.ShowOpen
If CMN.FileName = "" Then Exit Sub
Text1.Text = CMN.FileName
On Error Resume Next
Spin.ChangePicture LoadPicture(Text1.Text)
End Sub

Private Sub Command3_Click()
Spin.StopSpin
Unload Me
End Sub

Private Sub Form_Load()
Spin.StartSpin
Spin.ChangePicture Form1.Picture1.Picture
End Sub

Private Sub HScroll1_Change()
Spin.ChangeSpeed HScroll1.Value
End Sub
