VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmGetConnected 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1605
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4785
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmGetConnected.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2415
      Top             =   1785
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Attempting Connection ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Index           =   0
      Left            =   420
      TabIndex        =   1
      Top             =   45
      Width           =   3795
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Waiting ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   1
      Left            =   460
      TabIndex        =   2
      Top             =   65
      Width           =   3795
   End
   Begin VB.Image imgMin 
      Height          =   225
      Left            =   4305
      MouseIcon       =   "frmGetConnected.frx":058A
      MousePointer    =   99  'Custom
      Top             =   50
      Width           =   225
   End
   Begin VB.Image imgClose 
      Height          =   225
      Left            =   4515
      MouseIcon       =   "frmGetConnected.frx":0894
      MousePointer    =   99  'Custom
      Top             =   50
      Width           =   225
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Initialising proxy connection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   210
      TabIndex        =   0
      Top             =   840
      Width           =   4380
   End
   Begin VB.Image imgBackDrop 
      Height          =   1605
      Left            =   0
      Picture         =   "frmGetConnected.frx":0B9E
      Top             =   0
      Width           =   4800
   End
   Begin VB.Menu mnuTray 
      Caption         =   "mnuTray"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show Download"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop Download"
      End
   End
End
Attribute VB_Name = "frmGetConnected"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Systray Data
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function dada Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim t As NOTIFYICONDATA
' End of Systray

Dim txtTargetFile As String
Dim vFullFile As String
Dim Bytes() As Byte
Dim bStarted As Boolean
Dim lTotalSize As Long
Dim gbCheck As Boolean
Dim GridLine As Integer
Dim bCompleted As Boolean
Dim sError As String
Dim PreviousBytes As Long

'Time Calc
Dim lLastTime As Double
Dim tLastTime As Double
Dim lTime As Double
Dim lTimeDiff As Double
Dim lTimeLeft As Double
Dim lTotalTime As Double
Dim tTimeStarted As Double
Dim bTimeStarted As Boolean
'Dim sProtocol As String
Const BytesSize = 256
Public Sub StartConnection()

On Error GoTo error_h

  
    'Set up connection details
 
    'downctl.Protocol = GetProtocol(Url)
    Inet1.Url = "http://www.yahoo.com"
    Inet1.Proxy = GetSetting("BlueByte", "Connection", "ProxyAddr", "")

    'inet.AccessType = icNamedProxy
    
    Inet1.UserName = GetSetting("BlueByte", "Connection", "ProxyUsr", "")
    Inet1.Password = GetSetting("BlueByte", "Connection", "ProxyPass", "")
    Me.Show
    
    DoEvents
    
   Dim download() As Byte
    

    lblHeader = "Refreshing Proxy Connection"
    download() = Inet1.OpenURL(, icByteArray)
    lblHeader = "Connected"
    frmMain.Visible = True
    Unload Me
    DoEvents
    Exit Sub
    
error_h:
    
    Unload Me

End Sub



Private Sub Form_Load()

Me.Width = imgBackDrop.Width
Me.Height = imgBackDrop.Height

'Position the Download Form
Me.Left = frmMain.Width / 2
Me.Top = frmMain.Height / 2
DoEvents
Me.Show
DoEvents

End Sub

Private Sub Form_Unload(Cancel As Integer)
Inet1.Cancel
DoEvents
End Sub

Private Sub imgBackDrop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Call DragForm(Me)
End If
End Sub

Private Sub imgClose_Click()
Unload frmMain
Inet1.Cancel
DoEvents
Unload Me
End Sub

Private Sub imgMin_Click()
Me.Hide
End Sub

Private Sub lblCaption_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    Call DragForm(Me)
End If

End Sub
