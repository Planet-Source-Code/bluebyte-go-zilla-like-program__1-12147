VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrPause 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2625
      Top             =   2835
   End
   Begin VB.PictureBox picMainSkin 
      Height          =   3060
      Left            =   0
      MouseIcon       =   "frmSplash.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmSplash.frx":030A
      ScaleHeight     =   3000
      ScaleWidth      =   6015
      TabIndex        =   0
      Top             =   0
      Width           =   6075
      Begin VB.Label lblOs 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "For Windows"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   3045
         TabIndex        =   2
         Top             =   1995
         Width           =   1275
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Loading ..."
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   2310
         TabIndex        =   1
         Top             =   2310
         Width           =   2010
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()

    lblVersion = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    Dim WindowRegion As Long
    ' I set all these settings here so you won't forget
    ' them and have a non-working demo... Set them in
    ' design time
    picMainSkin.ScaleMode = vbPixels
    picMainSkin.AutoRedraw = True
    picMainSkin.AutoSize = True
    picMainSkin.BorderStyle = vbBSNone
    Me.BorderStyle = vbBSNone
        
    'Set picMainSkin.Picture = LoadPicture(App.Path & "\bigsqueel.bmp")
    
    Me.Width = picMainSkin.Width
    Me.Height = picMainSkin.Height
    
    WindowRegion = MakeRegion(picMainSkin)
    SetWindowRgn Me.hWnd, WindowRegion, True
    Me.Top = Screen.Height / 2 - Me.Height / 2
    Me.Left = Screen.Width / 2 - Me.Width / 2
    DoEvents
    'Me.Show
    tmrPause.Enabled = True
    End Sub

Private Sub picMainSkin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      
      ' Pass the handling of the mouse down message to
      ' the (non-existing really) form caption, so that
      ' the form itself will be dragged when the picture is dragged.
      '
      ' If you have Win 98, Make sure that the "Show window
      ' contents while dragging" display setting is on for nice results.
      
      'ReleaseCapture
      'SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub



Private Sub tmrPause_Timer()
frmMain.Show
DoEvents
Unload Me
End Sub


