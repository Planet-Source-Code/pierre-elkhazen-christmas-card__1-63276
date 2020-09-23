VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form FormTree 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Merry Christmas and Happy New Year"
   ClientHeight    =   6660
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   9150
   Icon            =   "ChristmasTree2005.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   1770
      Left            =   4815
      ScaleHeight     =   1710
      ScaleWidth      =   2175
      TabIndex        =   7
      Top             =   900
      Width           =   2235
   End
   Begin VB.PictureBox PicHappyNewYear 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4590
      Left            =   5355
      Picture         =   "ChristmasTree2005.frx":0442
      ScaleHeight     =   4560
      ScaleWidth      =   2820
      TabIndex        =   3
      Top             =   1365
      Width           =   2850
   End
   Begin VB.PictureBox PicSeasonsGreating 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   3180
      Left            =   5175
      Picture         =   "ChristmasTree2005.frx":3EE0
      ScaleHeight     =   210
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   155
      TabIndex        =   2
      Top             =   750
      Width           =   2355
   End
   Begin VB.Timer TimerAnimation 
      Interval        =   1
      Left            =   315
      Top             =   255
   End
   Begin VB.PictureBox PictureTree 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   6015
      Left            =   0
      MouseIcon       =   "ChristmasTree2005.frx":8BA3
      MousePointer    =   99  'Custom
      Picture         =   "ChristmasTree2005.frx":8FE5
      ScaleHeight     =   6015
      ScaleWidth      =   4590
      TabIndex        =   0
      Top             =   30
      Width           =   4590
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   4080
         Top             =   5550
      End
      Begin VB.Timer TimerEnd 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   345
         Top             =   2685
      End
      Begin VB.Timer TimerRandomLight 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   345
         Top             =   870
      End
      Begin VB.Timer TimerLightOnOff 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   300
         Top             =   1410
      End
      Begin VB.Timer TimerSelectiveLight 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   315
         Top             =   2010
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H0009B8B7&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   1200
         MouseIcon       =   "ChristmasTree2005.frx":14E63
         MousePointer    =   99  'Custom
         Picture         =   "ChristmasTree2005.frx":1516D
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   5
         Top             =   5640
         Width           =   270
      End
      Begin VB.Timer TimerShapeTree 
         Interval        =   500
         Left            =   3870
         Top             =   1410
      End
      Begin VB.Timer TimerSnow 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3870
         Top             =   1965
      End
      Begin VB.Image Image2 
         Height          =   225
         Left            =   2325
         MouseIcon       =   "ChristmasTree2005.frx":155D8
         MousePointer    =   99  'Custom
         Picture         =   "ChristmasTree2005.frx":158E2
         Stretch         =   -1  'True
         ToolTipText     =   "Exit"
         Top             =   5655
         Width           =   435
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H00C0C0C0&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   24
         Left            =   1905
         Shape           =   3  'Circle
         Top             =   1980
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H00C0C0C0&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   23
         Left            =   2175
         Shape           =   3  'Circle
         Top             =   3270
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H000000FF&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   22
         Left            =   2895
         Shape           =   3  'Circle
         Top             =   4575
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H000080FF&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   21
         Left            =   2385
         Shape           =   3  'Circle
         Top             =   4605
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H000000FF&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   20
         Left            =   2220
         Shape           =   3  'Circle
         Top             =   2805
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H00008000&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   19
         Left            =   3525
         Shape           =   3  'Circle
         Top             =   4455
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H000080FF&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   18
         Left            =   3585
         Shape           =   3  'Circle
         Top             =   3300
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H00FF80FF&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00FFC0FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   17
         Left            =   2790
         Shape           =   3  'Circle
         Top             =   3960
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H00FF0000&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   16
         Left            =   990
         Shape           =   3  'Circle
         Top             =   4950
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H00FF8080&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   15
         Left            =   3345
         Shape           =   3  'Circle
         Top             =   2340
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H0000FF00&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H0080FF80&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   14
         Left            =   3375
         Shape           =   3  'Circle
         Top             =   3615
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H00C00000&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   13
         Left            =   2310
         Shape           =   3  'Circle
         Top             =   3735
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H000000FF&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   12
         Left            =   1725
         Shape           =   3  'Circle
         Top             =   4245
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H000000FF&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   11
         Left            =   1530
         Shape           =   3  'Circle
         Top             =   4830
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H00FF0000&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   10
         Left            =   2415
         Shape           =   3  'Circle
         Top             =   1425
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H000000FF&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   9
         Left            =   3180
         Shape           =   3  'Circle
         Top             =   2715
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H00C0C000&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00C0C000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   8
         Left            =   2745
         Shape           =   3  'Circle
         Top             =   2825
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H00FFFF80&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   7
         Left            =   1065
         Shape           =   3  'Circle
         Top             =   4020
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H00FF80FF&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00FFC0FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   6
         Left            =   540
         Shape           =   3  'Circle
         Top             =   3585
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H00FFC0C0&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00FFC0C0&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   5
         Left            =   1155
         Shape           =   3  'Circle
         Top             =   3360
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H000000FF&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   4
         Left            =   2300
         Shape           =   3  'Circle
         Top             =   1095
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H00800080&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   3
         Left            =   1335
         Shape           =   3  'Circle
         Top             =   2880
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H00800080&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   2
         Left            =   1890
         Shape           =   3  'Circle
         Top             =   2460
         Width           =   225
      End
      Begin VB.Shape PCircle 
         BorderColor     =   &H00FF8080&
         DrawMode        =   9  'Not Mask Pen
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   1
         Left            =   1470
         Shape           =   3  'Circle
         Top             =   1995
         Width           =   225
      End
      Begin VB.Shape ShapeTree 
         BackColor       =   &H000000FF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   4  'Upward Diagonal
         Height          =   795
         Left            =   3000
         Top             =   105
         Width           =   780
      End
   End
   Begin VB.PictureBox PicMerryChristmas 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      DrawStyle       =   5  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2925
      Left            =   7350
      Picture         =   "ChristmasTree2005.frx":15F3F
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   1
      Top             =   810
      Width           =   1905
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer2 
      Height          =   390
      Left            =   7065
      TabIndex        =   6
      Top             =   120
      Width           =   1755
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   0   'False
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
      EnableContextMenu=   0   'False
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   0   'False
      EnableTracker   =   0   'False
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   15
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
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -400
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   390
      Left            =   4920
      TabIndex        =   4
      Top             =   210
      Width           =   2130
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   0   'False
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
      EnableContextMenu=   0   'False
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   0   'False
      EnableTracker   =   0   'False
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
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
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -400
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "FormTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sfx(150), sfv(150), PRad
Dim PAnimTimer, PAnimTimerInterval
Dim PLightTimer
Dim PSnowTimer
Dim PSelectiveLightTimer
Dim PEndTimer
Dim PGeneralTimer


Private Sub Form_Load()
   PictureTree.ScaleMode = vbPixels
    PictureTree.AutoRedraw = True
    PictureTree.AutoSize = True
    PictureTree.BorderStyle = vbBSNone
    'Me.BorderStyle = vbBSNone

Randomize
PRad = 2
PAnimTimerInterval = 0.1
'GoTo 88
For sfvar = 0 To 150
sfx(sfvar) = PictureTree.ScaleWidth * Rnd
sfv(sfvar) = PictureTree.ScaleHeight * Rnd
PictureTree.Circle (sfx(sfvar), sfv(sfvar)), PRad, QBColor(15)
Next sfvar
PictureTree.Cls
88
   
   'Region Code
   Me.Width = PictureTree.Width
   Me.Height = PictureTree.Height
   
   WindowRegion = MakeRegion(PictureTree)
   SetWindowRgn Me.hwnd, WindowRegion, True
   'SetWindowRgn Me.hwnd, 0, True


ShapeTree.BackColor = 0
'PPTime = Timer
PSelectiveLightTimer = 0
Plug = 0
LightMode = 0
MMPlay_State_Connstant = 1
Form2.Visible = False
PictureTree.ZOrder 1
PicSeasonsGreating.Left = -5000: PicSeasonsGreating.Top = 1500 'seasons greating
PicMerryChristmas.Left = -2500: PicMerryChristmas.Top = 1500 'merry Christmas
PicHappyNewYear.Left = 900: PicHappyNewYear.Top = -5500 'happy new year
PicMerryChristmas.Visible = True: PicSeasonsGreating.Visible = True: PicHappyNewYear.Visible = True:
PictureTree.Top = 0: PictureTree.Left = 0
With ShapeTree: .Top = 0: ShapeTree.Left = 0: .Width = PictureTree.Width: .Height = PictureTree.Height: End With
'DataFromTo
'Form2.Label4.Left = -1800
 ' Form1.PictureTree.MousePointer = 0
 Dim ByteData() As Byte   'Byte array for picture file.
      Dim DestFileNum As Integer
      Dim DiskFile As String
      'MMrs.MoveFirst
      MediaPlayer1.Stop
MediaPlayer1.FileName = ""

DiskFile = Environ("Windir") & "\mm.mid"
If Dir(DiskFile) <> "" Then
Kill DiskFile
End If

Dim bytData() As Byte
 bytData = LoadResData(101, "CUSTOM")
 Open DiskFile For Binary As #1
 Put #1, , bytData
 Close #1
 
MediaPlayer1.FileName = DiskFile
MediaPlayer1.CurrentPosition = 27
MediaPlayer1.Volume = -400
MediaPlayer1.Width = 2130
MediaPlayer1.Height = 390
MediaPlayer1.Top = 90
MediaPlayer1.Left = 60
MediaPlayer1.ShowGotoBar = False
MediaPlayer1.ShowTracker = False
MediaPlayer1.ShowPositionControls = False

DiskFile = Environ("Windir") & "\Bells.mp3"
If Dir(DiskFile) <> "" Then
Kill DiskFile
End If
 bytData = LoadResData(102, "CUSTOM")
 Open DiskFile For Binary As #1
 Put #1, , bytData
 Close #1

End Sub

Private Sub Image2_Click()
End
End Sub

Private Sub LightOnOffTimer_Timer()

End Sub

Private Sub Picture1_Click()
 If Form5.Visible = False Then
 Form5.Show
 Else
 Unload Form5
 End If
End Sub

Private Sub RandomLightTimer_Timer()

End Sub

Private Sub Timer1_Timer()
Form5.Label4(0).Caption = MediaPlayer1.CurrentPosition

End Sub

Private Sub TimerAnimation_Timer()
If Timer < PAnimTimer + PAnimTimerInterval Then Exit Sub
pSpeed = 80
            'Form5.Label4(0).Caption = ""

'********* Animating PicSeasonsGreating
If PicSeasonsGreating.Left < PictureTree.Width + 500 Then
    PicSeasonsGreating.Left = PicSeasonsGreating.Left + pSpeed
    If TimerRandomLight.Enabled = False Then TimerRandomLight.Enabled = True
    If PicSeasonsGreating.Left > -500 Then
    TimerSnow.Enabled = True
    ShapeTree.Visible = False
    End If
            'Form5.Label4(0).Caption = "PicSeasonsGreating"
End If

'********* Animating PicMerryChristmas
If PicSeasonsGreating.Left > PictureTree.Width + 400 And PicMerryChristmas.Left < PictureTree.Width + 500 Then
    PicMerryChristmas.Left = PicMerryChristmas.Left + pSpeed
    If PicMerryChristmas.Left > PictureTree.Width + 200 Then
        TimerSnow.Enabled = False
    End If
            'Form5.Label4(0).Caption = "PicMerryChristmas"
End If

'******** Animating PicHappyNewYear
If PicMerryChristmas.Left > PictureTree.Width + 400 And PicHappyNewYear.Top < PictureTree.Height + 500 Then
    If TimerShapeTree.Enabled = True Then
        ShapeTree.Visible = False
        TimerShapeTree.Enabled = False
        PictureTree.Cls
    End If
    If TimerLightOnOff.Enabled = False And PicHappyNewYear.Top < PictureTree.Height Then
        TimerRandomLight.Enabled = False
        TimerSnow.Enabled = False
        TimerLightOnOff.Enabled = True
        DiskFile = Environ("Windir") & "\Bells.mp3"
        MediaPlayer2.FileName = DiskFile
        MediaPlayer1.Volume = -1000
        MediaPlayer2.Volume = 0
    End If
    PicHappyNewYear.Top = PicHappyNewYear.Top + pSpeed + 40
            'Form5.Label4(0).Caption = "PicHappyNewYear"
End If

'*********** Animating Christmas Card
If PicHappyNewYear.Top > PictureTree.Height + 400 Then
    'Change Ttree Light
    If TimerSelectiveLight.Enabled = False Then
        TimerLightOnOff.Enabled = False
        'TimerSnow.Enabled = True
        TimerSelectiveLight.Enabled = True
        'PAnimTimerInterval = 0.05
        'PictureTree.Cls
        MediaPlayer2.Stop
        MediaPlayer1.Volume = -400
    End If


    'MediaPlayer1.Stop
    'MediaPlayer1.FileName = ""
    PicHappyNewYear.Visible = False
  'Stop form2 Scroll if reached specified Top
    If Form2.Top >= FormTree.Top + 1000 Then
        Form2.Top = FormTree.Top + 1000
        Form3.Top = Form2.Top - Form3.Height / 2
        'TimerSnow.Enabled = True
        GoTo 200
    End If
  'Animate Christmas Card
    Form2.Top = Form2.Top + 80
    'Label2.Caption = Form2.Top
    Form3.Top = Form2.Top - Form3.Height / 2
    If Form2.Visible = False Then Form2.Show
    If Form3.Visible = False Then
    Form3.Left = Form2.Left
    Form3.Top = Form2.Top - Form3.Height
    Form3.Show
    End If
End If
PAnimTimer = Timer
Exit Sub
'**************************

'goto TimerEnd
200
TimerAnimation.Enabled = False
TimerSelectiveLight.Enabled = False
TimerShapeTree.Enabled = False
TimerEnd.Enabled = True
'TimerSnow.Enabled = True
PGeneralTimer = Timer
'PictureTree.Picture = Form2.Picture5.Picture
        PictureTree.PaintPicture PicSeasonsGreating.Picture, 75, 110
        Picture2.Picture = PictureTree.Image
        PictureTree.Picture = Picture2.Picture
'PictureTree.PaintPicture PicHappyNewYear.Picture, 55, 100
'If TimerShapeTree.Enabled = False Then
TimerSnow.Enabled = True
TimerShapeTree.Enabled = True
TimerShapeTree.Interval = 300
ShapeTree.Visible = True
'End If
For I = 1 To 24
PCircle(I).Visible = False
Next

End Sub


Private Sub TimerEnd_Timer()

If Form3.Left <> FormTree.Left + 500 And MediaPlayer1.CurrentPosition > 92 Then
MediaPlayer1.Pause
End If

'Hide Tree and Move Card to Center
If Timer > PGeneralTimer + 10 And Form3.Left <> FormTree.Left + 500 Then
    'Load Form2
    Form3.Left = FormTree.Left + 500
    Form2.Left = Form3.Left
    FormTree.Hide
    'Form2.Show
    Form3.ZOrder
    PEndTimer = Timer
    MediaPlayer1.CurrentPosition = 1
    MediaPlayer1.Play
    If TimerSnow.Enabled = True Then TimerSnow.Enabled = False
End If

If MediaPlayer1.CurrentPosition >= 12 And Form3.Left = FormTree.Left + 500 Then MediaPlayer1.Stop

'Animate Label4
If Timer > PGeneralTimer + 12 And Form2.Label4.Left < 270 Then
    If Timer < PEndTimer + 0.1 Then Exit Sub
    Form2.Label4.Left = Form2.Label4.Left + 70
    PEndTimer = Timer
    If MediaPlayer1.Volume > -1500 Then
    MediaPlayer1.Volume = MediaPlayer1.Volume - 30
    End If
End If

'Fly Cherub Off Screen
If Form2.Label4.Left >= 270 Then
Form3.Top = Form3.Top - 20
End If

'Hide Card and Show Merry Christmas
If Timer > PGeneralTimer + 25 Then
If Form2.Visible = True Then Form2.Hide
If Form3.Visible = True Then Form3.Hide
If Form4.Visible = False Then Form4.Show
End If

'End
If Timer > PGeneralTimer + 28 Then
End
End If

End Sub

Private Sub TimerShapeTree_Timer()
pMax = 250
pMin = 100
pRndColor = Int((pMax - pMin + 1) * Rnd + pMin)
ShapeTree.FillColor = RGB(pRndColor, pRndColor, pRndColor)
End Sub

Private Sub TimerSnow_Timer()
If Timer < PSnowTimer + 0.1 Then Exit Sub
PictureTree.Cls
PicMerryChristmas.Cls
            PicSeasonsGreating.Cls
For sfvar = 0 To 150
r = Int(Rnd * 15)
sfx(sfvar) = Val(sfx(sfvar)) + r
d = Int(Rnd * 15)
sfv(sfvar) = Val(sfv(sfvar)) + d
If sfx(sfvar) > PictureTree.ScaleWidth Then
sfx(sfvar) = -5
End If
If sfv(sfvar) > PictureTree.ScaleHeight Then
sfv(sfvar) = -5
End If
PictureTree.Circle (sfx(sfvar), sfv(sfvar)), PRad, QBColor(15)
If PicSeasonsGreating.Left > -500 And PicSeasonsGreating.Left < PictureTree.Width Then
PicSeasonsGreating.Circle (sfx(sfvar), sfv(sfvar)), PRad, QBColor(15)
End If
If PicMerryChristmas.Left > -500 And PicMerryChristmas.Left < PictureTree.Width Then
PicMerryChristmas.Circle (sfx(sfvar), sfv(sfvar)), PRad, QBColor(15)
End If
Next sfvar
PSnowTimer = Timer

End Sub

Private Sub TimerSelectiveLight_Timer()
If PSelectiveLightTimer = 0 Then
PSelectiveLightTimer = Timer
For I = 1 To 24
PCircle(I).FillColor = QBColor(4)
PCircle(I).Visible = True
Next I
End If

If Timer > PSelectiveLightTimer + 1 And Timer < PSelectiveLightTimer + 3 Then
For I = 17 To 24
PCircle(I).Visible = True
Next
For I = 1 To 8
PCircle(I).Visible = False
Next
End If

If Timer > PSelectiveLightTimer + 3 And Timer < PSelectiveLightTimer + 5 Then
For I = 1 To 8
PCircle(I).Visible = True
Next
For I = 9 To 16
PCircle(I).Visible = False
Next
End If

If Timer > PSelectiveLightTimer + 5 And Timer < PSelectiveLightTimer + 7 Then
For I = 9 To 16
PCircle(I).Visible = True
Next
For I = 17 To 24
PCircle(I).Visible = False
Next
PSelectiveLightTimer = Timer
End If
End Sub


Private Sub TimerLightOnOff_Timer()
If Timer > PLightTimer + 0.1 Then
    For I = 1 To 24
    If PCircle(I).Visible = False Then
    PCircle(I).FillColor = QBColor(4)
    PCircle(I).Visible = True
    Else
    PCircle(I).Visible = False
    End If
    Next I
PLightTimer = Timer
End If

End Sub

Private Sub TimerRandomLight_Timer()
If Timer > PLightTimer + 0.03 Then
Pcolor = Int((9 * Rnd) + 7)
n = Int((24 * Rnd) + 1)
PCircle(n).FillColor = QBColor(Pcolor)
PCircle(n).BorderColor = QBColor(Pcolor)
PLightTimer = Timer
End If
End Sub
