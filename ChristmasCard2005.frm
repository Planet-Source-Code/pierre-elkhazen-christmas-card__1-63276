VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Merry Christmas"
   ClientHeight    =   4395
   ClientLeft      =   1005
   ClientTop       =   -7005
   ClientWidth     =   3240
   ControlBox      =   0   'False
   Icon            =   "ChristmasCard2005.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   216
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00000040&
      DisabledPicture =   "ChristmasCard2005.frx":0442
      DownPicture     =   "ChristmasCard2005.frx":171A
      Height          =   240
      Left            =   3525
      Picture         =   "ChristmasCard2005.frx":29F2
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Save Rerord"
      Top             =   3855
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   3540
      TabIndex        =   12
      Text            =   "Sender"
      Top             =   3900
      Width           =   2625
   End
   Begin VB.PictureBox Picture10 
      BackColor       =   &H000000FF&
      Height          =   70
      Left            =   3315
      ScaleHeight     =   15
      ScaleWidth      =   3105
      TabIndex        =   10
      Top             =   0
      Width           =   3160
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H000000FF&
      Height          =   4290
      Left            =   3270
      ScaleHeight     =   4230
      ScaleWidth      =   15
      TabIndex        =   9
      Top             =   40
      Width           =   70
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H000000FF&
      Height          =   4290
      Left            =   6435
      ScaleHeight     =   4230
      ScaleWidth      =   15
      TabIndex        =   8
      Top             =   40
      Width           =   70
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H000000FF&
      Height          =   70
      Left            =   3315
      ScaleHeight     =   15
      ScaleWidth      =   3105
      TabIndex        =   7
      Top             =   4290
      Width           =   3160
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3615
      TabIndex        =   6
      Text            =   "Receiver"
      Top             =   3300
      Width           =   2520
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H000000FF&
      Height          =   70
      Left            =   40
      ScaleHeight     =   15
      ScaleWidth      =   3105
      TabIndex        =   3
      Top             =   4290
      Width           =   3160
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H000000FF&
      Height          =   4290
      Left            =   3135
      ScaleHeight     =   4230
      ScaleWidth      =   15
      TabIndex        =   5
      Top             =   40
      Width           =   70
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H000000FF&
      Height          =   4290
      Left            =   0
      ScaleHeight     =   4230
      ScaleWidth      =   15
      TabIndex        =   4
      Top             =   40
      Width           =   70
   End
   Begin VB.PictureBox Picture5 
      AutoSize        =   -1  'True
      Height          =   3570
      Left            =   405
      MouseIcon       =   "ChristmasCard2005.frx":3CCA
      MousePointer    =   99  'Custom
      Picture         =   "ChristmasCard2005.frx":3FD4
      ScaleHeight     =   3510
      ScaleWidth      =   2340
      TabIndex        =   0
      Top             =   420
      Width           =   2400
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00C0E0FF&
         FillStyle       =   7  'Diagonal Cross
         Height          =   3690
         Left            =   0
         Top             =   0
         Width           =   2580
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Born to you this day a Saviour Who is Christ the Lord"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   645
         Left            =   -1800
         TabIndex        =   1
         Top             =   2805
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      Height          =   70
      Left            =   15
      ScaleHeight     =   15
      ScaleWidth      =   3105
      TabIndex        =   2
      Top             =   0
      Width           =   3160
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer2 
      Height          =   1755
      Left            =   3750
      TabIndex        =   11
      Top             =   405
      Width           =   2235
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   -1  'True
      PlayCount       =   25
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Shape Shape14 
      FillColor       =   &H00FFFFC0&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   2415
      Width           =   75
   End
   Begin VB.Shape Shape13 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   5565
      Shape           =   3  'Circle
      Top             =   2685
      Width           =   75
   End
   Begin VB.Shape Shape12 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   3585
      Shape           =   3  'Circle
      Top             =   3090
      Width           =   75
   End
   Begin VB.Shape Shape11 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   5940
      Shape           =   3  'Circle
      Top             =   3660
      Width           =   75
   End
   Begin VB.Shape Shape10 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   3750
      Width           =   75
   End
   Begin VB.Shape Shape9 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   5985
      Shape           =   3  'Circle
      Top             =   3015
      Width           =   75
   End
   Begin VB.Shape Shape8 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   75
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   3780
      Shape           =   3  'Circle
      Top             =   3615
      Width           =   75
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   4500
      Shape           =   3  'Circle
      Top             =   2250
      Width           =   75
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   5940
      Shape           =   3  'Circle
      Top             =   2460
      Width           =   75
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   3765
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   75
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pierre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   3555
      TabIndex        =   15
      Top             =   3855
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "All Children"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   3645
      TabIndex        =   14
      Top             =   3330
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H00C0E0FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   3900
      Left            =   230
      Top             =   255
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   3735
      Picture         =   "ChristmasCard2005.frx":66AB
      Stretch         =   -1  'True
      Top             =   2310
      Width           =   2235
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H0080C0FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   3900
      Left            =   3495
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    TransparentForm Me
End Sub

Private Sub Form_Paint()
Me.Left = Form1.Left - Me.Width
Shape1.Move 15, 20
Shape2.Move 3390, 20
End Sub

Private Sub Form_Resize()
    TransparentForm Me

End Sub
Public Sub TransparentForm(frm As Form)
    frm.ScaleMode = vbPixels
    Const RGN_DIFF = 4
    Const RGN_OR = 2

    Dim outer_rgn As Long
    Dim inner_rgn As Long
    Dim wid As Single
    Dim hgt As Single
    Dim border_width As Single
    Dim title_height As Single
    Dim ctl_left As Single
    Dim ctl_top As Single
    Dim ctl_right As Single
    Dim ctl_bottom As Single
    Dim control_rgn As Long
    Dim combined_rgn As Long
    Dim ctl As Control
If frm.WindowState = vbMinimized Then Exit Sub

    ' Create the main form region.
    wid = frm.ScaleX(frm.Width, vbTwips, vbPixels)
    hgt = frm.ScaleY(frm.Height, vbTwips, vbPixels)
    outer_rgn = CreateRectRgn(0, 0, wid, hgt)

    border_width = (wid - frm.ScaleWidth) / 2
    title_height = hgt - border_width - frm.ScaleHeight
    inner_rgn = CreateRectRgn(border_width, title_height, wid - border_width, _
        hgt - border_width)

    ' Subtract the inner region from the outer.
    combined_rgn = CreateRectRgn(0, 0, 0, 0)
    CombineRgn combined_rgn, outer_rgn, inner_rgn, RGN_DIFF

    ' Create the control regions.
    For Each ctl In frm.Controls
        If ctl.Container Is frm Then
            ctl_left = frm.ScaleX(ctl.Left, frm.ScaleMode, vbPixels) _
                + border_width
            ctl_top = frm.ScaleX(ctl.Top, frm.ScaleMode, vbPixels) + title_height
            ctl_right = frm.ScaleX(ctl.Width, frm.ScaleMode, vbPixels) + ctl_left
            ctl_bottom = frm.ScaleX(ctl.Height, frm.ScaleMode, vbPixels) + ctl_top
            control_rgn = CreateRectRgn(ctl_left, ctl_top, ctl_right, ctl_bottom)
            CombineRgn combined_rgn, combined_rgn, control_rgn, RGN_OR
        End If
    Next ctl

    'Restrict the window to the region.
    SetWindowRgn frm.hwnd, combined_rgn, True
End Sub



Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xpos = X
ypos = Y

End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Me.Move X + (Me.Left - xpos), Y + (Me.Top - ypos)
End If
End Sub

Private Sub Text1_Click()
CmdSave.Visible = True
End Sub

Private Sub Text2_Click()
CmdSave.Visible = True
End Sub
