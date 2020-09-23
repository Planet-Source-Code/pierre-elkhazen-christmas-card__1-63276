VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   2580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4275
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   ScaleHeight     =   2580
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picMessage 
      AutoSize        =   -1  'True
      Height          =   2415
      Left            =   0
      MouseIcon       =   "FormMessage.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "FormMessage.frx":030A
      ScaleHeight     =   2355
      ScaleWidth      =   4080
      TabIndex        =   0
      Top             =   0
      Width           =   4140
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Visit my Medjugorje Site"
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   3
         Left            =   540
         TabIndex        =   4
         Top             =   285
         Width           =   2580
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(Self Running EXE. No Setup Required. Works on all Windows Versions)"
         ForeColor       =   &H00000080&
         Height          =   405
         Index           =   1
         Left            =   315
         TabIndex        =   3
         Top             =   1170
         Width           =   3000
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Compiled EXE File Ready to be sent to your Friends available at that Site."
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   0
         Left            =   330
         TabIndex        =   2
         Top             =   735
         Width           =   3135
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "http://geocities.com/medjugorjesite"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   285
         MouseIcon       =   "FormMessage.frx":1F7BC
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   495
         Width           =   3090
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WindowRegion As Long
Private Sub Form_Load()
    picMessage.Left = 0
    picMessage.Top = 0
    picMessage.ScaleMode = vbPixels
    picMessage.AutoRedraw = True
    picMessage.AutoSize = True
    picMessage.BorderStyle = vbBSNone
   Me.Width = picMessage.Width
   Me.Height = picMessage.Height
    
    WindowRegion = MakeRegion(picMessage)
    SetWindowRgn Me.hwnd, WindowRegion, True

End Sub

Private Sub Label4_Click(Index As Integer)
Unload Me
End Sub

Private Sub Label5_Click()
ShellExecute Me.hwnd, "open", "http://geocities.com/medjugorjesite", ByVal 0&, "", 3

End Sub


Private Sub picMessage_Click()
Unload Me

End Sub
