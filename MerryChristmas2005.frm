VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   2115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3120
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   ScaleHeight     =   2115
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   0
      Picture         =   "MerryChristmas2005.frx":0000
      ScaleHeight     =   1950
      ScaleWidth      =   3000
      TabIndex        =   0
      Top             =   0
      Width           =   3000
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Form_Load()
    Picture1.ScaleMode = vbPixels
    Picture1.AutoRedraw = True
    Picture1.AutoSize = True
    Picture1.BorderStyle = vbBSNone
    'Me.BorderStyle = vbBSNone
    
   Me.Width = Picture1.Width
   Me.Height = Picture1.Height
    
    WindowRegion = MakeRegion(Picture1)
    SetWindowRgn Me.hwnd, WindowRegion, True

End Sub
