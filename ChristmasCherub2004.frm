VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2430
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2595
   ScaleWidth      =   2430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   765
      Index           =   4
      Left            =   5520
      Picture         =   "ChristmasCherub2004.frx":0000
      ScaleHeight     =   705
      ScaleWidth      =   945
      TabIndex        =   4
      Top             =   1515
      Width           =   1005
   End
   Begin VB.PictureBox Picture1 
      Height          =   765
      Index           =   3
      Left            =   4125
      Picture         =   "ChristmasCherub2004.frx":1308
      ScaleHeight     =   705
      ScaleWidth      =   945
      TabIndex        =   3
      Top             =   1515
      Width           =   1005
   End
   Begin VB.PictureBox Picture1 
      Height          =   765
      Index           =   2
      Left            =   5505
      Picture         =   "ChristmasCherub2004.frx":2601
      ScaleHeight     =   705
      ScaleWidth      =   945
      TabIndex        =   2
      Top             =   420
      Width           =   1005
   End
   Begin VB.PictureBox Picture1 
      Height          =   765
      Index           =   1
      Left            =   4155
      Picture         =   "ChristmasCherub2004.frx":3909
      ScaleHeight     =   705
      ScaleWidth      =   945
      TabIndex        =   1
      Top             =   420
      Width           =   1005
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2715
      Top             =   3000
   End
   Begin VB.PictureBox picMainSkin 
      AutoSize        =   -1  'True
      Height          =   2505
      Left            =   0
      MouseIcon       =   "ChristmasCherub2004.frx":4BFF
      MousePointer    =   99  'Custom
      Picture         =   "ChristmasCherub2004.frx":4F09
      ScaleHeight     =   2445
      ScaleWidth      =   2250
      TabIndex        =   0
      Top             =   0
      Width           =   2310
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Dim WindowRegion As Long
Dim n
Dim Px
Dim Py
Dim xpos As Long
Dim ypos As Long
Dim Ptimerp

Private Sub Form_Activate()
 Px = Me.Left
 Py = Me.Top
End Sub

Private Sub Form_Load()
    picMainSkin.ScaleMode = vbPixels
    picMainSkin.AutoRedraw = True
    picMainSkin.AutoSize = True
    picMainSkin.BorderStyle = vbBSNone
    'Me.BorderStyle = vbBSNone
      n = 2
    Ptimerp = Timer
   Me.Width = picMainSkin.Width
   Me.Height = picMainSkin.Height
    
    WindowRegion = MakeRegion(picMainSkin)
    SetWindowRgn Me.hwnd, WindowRegion, True
End Sub

Private Sub picMainSkin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xpos = X
ypos = Y
 Timer1.Enabled = False

End Sub


Private Sub picMainSkin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Me.Move X + (Me.Left - xpos), Y + (Me.Top - ypos)
 Px = X + (Me.Left - xpos)
 Py = Y + (Me.Top - ypos)
End If
End Sub

Private Sub picMainSkin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
If Timer < Ptimerp + 0.3 Then Exit Sub
Ptimerp = Timer
'Me.Left = Px
'Me.Top = Py
    'Set picMainSkin.Picture = LoadPicture(App.Path & "\" & n & ".bmp")
    picMainSkin.Picture = Picture1(n).Picture
    Me.Width = picMainSkin.Width
    Me.Height = picMainSkin.Height
    
    WindowRegion = MakeRegion(picMainSkin)
    SetWindowRgn Me.hwnd, WindowRegion, True
n = n + 1
If n > 4 Then n = 1
'Px = Px + 10
'Py = Py + 10
End Sub
