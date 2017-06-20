VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3825
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1770
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "GIF"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   13
      Top             =   1440
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   1815
      TabIndex        =   12
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2640
      Top             =   720
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "JPG"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   10
      Top             =   1440
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "BMP"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   9
      Top             =   1440
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "C:\"
      Top             =   1080
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3000
      MaxLength       =   5
      TabIndex        =   4
      Text            =   "120"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   0
      Text            =   "30"
      Top             =   360
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      Top             =   1425
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "开始"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   1440
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      Top             =   705
      Width           =   855
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "保存路径："
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "分"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "运行时间："
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "秒"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "截屏间隔："
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3840
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "更改"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   " 定时截屏 老郭专用版"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "X"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Dim totalTime As Long
Dim durTime As Long
Dim passedTime As Long

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ReleaseCapture
    SendMessage hwnd, &HA1, 2, 0&
End Sub

Private Sub Label6_Click()
    On Error Resume Next
    Form1.Hide
    Form2.Show
End Sub

Private Sub Label7_Click()
    On Error Resume Next
    durTime = Val(Text1.Text)
    totalTime = Val(Text2.Text) * 60
    passedTime = 0
    Form1.Hide
    Timer1.Enabled = True
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ReleaseCapture
    SendMessage hwnd, &HA1, 2, 0&
End Sub

Private Sub Label9_Click()
    On Error Resume Next
    End
End Sub

Private Sub Text1_Change()
    On Error Resume Next
    Dim i As Long
    i = Val(Text1.Text)
    If i < 1 Then i = 1
    If i > 9999 Then i = 9999
    Text1.Text = i
End Sub

Private Sub Text2_Change()
    On Error Resume Next
    Dim i As Long
    i = Val(Text2.Text)
    If i < 1 Then i = 1
    If i > 9999 Then i = 9999
    Text2.Text = i
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    passedTime = passedTime + 1
    If passedTime Mod durTime = 0 Then
        Picture1.Width = Screen.Width
        Picture1.Height = Screen.Height
        BitBlt Picture1.hDC, 0, 0, Picture1.Width, Picture1.Height, GetDC(0), 0, 0, &HCC0020
        Dim path As String
        Dim filePath As String
        path = Format(Now, "yyyy-mm-dd")
        MkDir Text3.Text & path
        filePath = Text3.Text & path & "\" & Format(Now, "hh mm ss")
        If Option1.value = True Then
            SavePicture Picture1.Image, filePath & ".bmp"
        ElseIf Option2.value = True Then
            SaveJPG Picture1.Image, filePath & ".jpg", 2
        Else
            SaveJPG Picture1.Image, filePath & ".gif", 1
        End If
    ElseIf passedTime > totalTime Then
        End
    End If
End Sub
