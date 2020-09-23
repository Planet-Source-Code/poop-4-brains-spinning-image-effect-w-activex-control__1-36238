VERSION 5.00
Begin VB.UserControl SpinControl 
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1920
   ScaleHeight     =   1800
   ScaleWidth      =   1920
   ToolboxBitmap   =   "UserControl1.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1320
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      Picture         =   "UserControl1.ctx":0312
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox Board 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   233
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "SpinControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Dim W, WS

Function ChangeBackground(Optional Color As Long = -1, Optional Picture As StdPicture = 0)
If Picture <> 0 Then
Board.Picture = Picture
End If

If Color <> -1 Then
Board.Picture = Nothing
Board.BackColor = Color
End If
End Function

Function ChangePicture(Picture As StdPicture)
Picture1.Picture = Picture
WS = -5
W = Picture1.ScaleWidth
End Function

Function StartSpin()
Timer1.Enabled = True
End Function

Function StopSpin()
Timer1.Enabled = False
Clear
End Function

Function Clear()
Board.Cls
End Function

Private Sub Timer1_Timer()
Board.Cls
W = W + WS
If W < -Picture1.ScaleWidth Then WS = 5
If W > Picture1.ScaleWidth Then WS = -5

StretchBlt Board.hdc, Board.ScaleWidth \ 2 - W \ 2, Board.ScaleHeight \ 2 - Picture1.ScaleHeight \ 2, W, Picture1.ScaleHeight, Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, vbSrcCopy
End Sub

Function ChangeSpeed(Speed As Long, Optional Max As Integer = 500)
Timer1.Interval = Max - Speed
End Function

Private Sub UserControl_Resize()
Board.Width = UserControl.ScaleWidth
Board.Height = UserControl.ScaleHeight
End Sub
