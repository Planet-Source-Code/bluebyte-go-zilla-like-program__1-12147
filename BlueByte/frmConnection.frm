VERSION 5.00
Begin VB.Form frmConnection 
   BorderStyle     =   0  'None
   ClientHeight    =   6015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   ControlBox      =   0   'False
   Icon            =   "frmConnection.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkHideWindows 
      BackColor       =   &H00800000&
      Caption         =   "Download Windows : On"
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   3465
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4410
      Width           =   2220
   End
   Begin VB.OptionButton optConnection 
      BackColor       =   &H00800000&
      Caption         =   "Connect Via Modem"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   0
      Left            =   1050
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   630
      Width           =   2010
   End
   Begin VB.OptionButton optConnection 
      BackColor       =   &H00800000&
      Caption         =   "Connect Via Proxy"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   1
      Left            =   3465
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   630
      Value           =   -1  'True
      Width           =   2010
   End
   Begin VB.TextBox txtProxy 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   1890
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2520
      Width           =   3060
   End
   Begin VB.TextBox txtProxy 
      Height          =   330
      Index           =   2
      Left            =   1890
      TabIndex        =   2
      Top             =   2100
      Width           =   3060
   End
   Begin VB.TextBox txtProxy 
      Height          =   330
      Index           =   1
      Left            =   1890
      TabIndex        =   1
      Top             =   1680
      Width           =   3060
   End
   Begin VB.TextBox txtProxy 
      Height          =   330
      Index           =   0
      Left            =   1890
      TabIndex        =   0
      Top             =   1260
      Width           =   3060
   End
   Begin VB.Label lblHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "Blue Byte"
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
      Index           =   0
      Left            =   525
      TabIndex        =   12
      Top             =   105
      Width           =   4950
   End
   Begin VB.Label lblHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "Blue Byte"
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
      Left            =   565
      TabIndex        =   13
      Top             =   125
      Width           =   4950
   End
   Begin VB.Image imgClose 
      Height          =   225
      Left            =   5670
      MouseIcon       =   "frmConnection.frx":27A2
      MousePointer    =   99  'Custom
      ToolTipText     =   "Click here to Exit this Window"
      Top             =   50
      Width           =   225
   End
   Begin VB.Image imgMin 
      Height          =   225
      Left            =   5460
      MouseIcon       =   "frmConnection.frx":2AAC
      MousePointer    =   99  'Custom
      ToolTipText     =   "Minimize Window"
      Top             =   50
      Width           =   225
   End
   Begin VB.Label cmdCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4410
      MouseIcon       =   "frmConnection.frx":2DB6
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   5430
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label cmdOk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2940
      MouseIcon       =   "frmConnection.frx":30C0
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   5430
      Width           =   1275
   End
   Begin VB.Image lblButton 
      Height          =   360
      Index           =   1
      Left            =   4410
      Picture         =   "frmConnection.frx":33CA
      Top             =   5355
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image lblButton 
      Height          =   360
      Index           =   0
      Left            =   2940
      Picture         =   "frmConnection.frx":4D2E
      Top             =   5355
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Proxy Address"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   0
      Left            =   210
      TabIndex        =   7
      Top             =   1260
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Proxy Port"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   1
      Left            =   210
      TabIndex        =   6
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User name"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   2
      Left            =   210
      TabIndex        =   5
      Top             =   2100
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Password"
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   3
      Left            =   210
      TabIndex        =   4
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Image imgBackDrop 
      Height          =   6000
      Left            =   0
      Picture         =   "frmConnection.frx":6692
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "frmConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkHideWindows_Click()

If chkHideWindows.Value = 0 Then
    chkHideWindows.Caption = "Download Windows : On"
Else
    chkHideWindows.Caption = "Download Windows : Off"
End If
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Call DragForm(Me)
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()

'Save the settings to the registry

If optConnection(1).Value = True Then
    SaveSetting "BlueByte", "Connection", "Type", "Proxy"
Else
    SaveSetting "BlueByte", "Connection", "Type", "Modem"
End If

SaveSetting "BlueByte", "Connection", "ProxyAddr", txtProxy(0).Text
SaveSetting "BlueByte", "Connection", "ProxyPort", txtProxy(1).Text
SaveSetting "BlueByte", "Connection", "ProxyUsr", txtProxy(2).Text
SaveSetting "BlueByte", "Connection", "ProxyPass", txtProxy(3).Text
SaveSetting "Bluebyte", "Connection", "dWindow", Me.chkHideWindows.Value

If GetSetting("BlueByte", "Connection", "Type", "Proxy") = "Proxy" Then
    frmMain.lblConnect.Caption = "Refresh Proxy"
Else
    frmMain.lblConnect.Caption = ""
End If

Unload Me

End Sub

Private Sub Form_Load()

'Position the Download Form
Me.Top = frmMain.Top
Me.Left = frmMain.Left

'Get the registry settings and fill the fields with them.

If GetSetting("BlueByte", "Connection", "Type", "Proxy") = "Proxy" Then
    optConnection(1).Value = True
Else
    optConnection(0).Value = True
End If

txtProxy(0).Text = GetSetting("BlueByte", "Connection", "ProxyAddr", txtProxy(0).Text)
txtProxy(1).Text = GetSetting("BlueByte", "Connection", "ProxyPort", txtProxy(1).Text)
txtProxy(2).Text = GetSetting("BlueByte", "Connection", "ProxyUsr", txtProxy(2).Text)
txtProxy(3).Text = GetSetting("BlueByte", "Connection", "ProxyPass", txtProxy(3).Text)
chkHideWindows.Value = GetSetting("BlueByte", "Connection", "dWindow", 0)

End Sub

Private Sub imgBackDrop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Call DragForm(Me)
End If
End Sub

Private Sub imgClose_Click()
Unload Me
End Sub

Private Sub imgForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Call DragForm(Me)
End If
End Sub

Private Sub imgMax_Click()
Me.WindowState = 2
End Sub

Private Sub imgMin_Click()
Me.WindowState = 1
End Sub

Private Sub lblHeader_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Call DragForm(Me)
End If
End Sub

