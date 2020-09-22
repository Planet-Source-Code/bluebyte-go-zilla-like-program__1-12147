VERSION 5.00
Object = "*\ADownloadcontrol.vbp"
Begin VB.Form frmDownLoad 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1650
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4875
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmDownLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin DownloadControl.Downloader DownCTL 
      Height          =   750
      Left            =   1155
      TabIndex        =   9
      Top             =   1995
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   1323
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4215
      Top             =   1845
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   2
      Left            =   2955
      MouseIcon       =   "frmDownLoad.frx":058A
      MousePointer    =   99  'Custom
      Picture         =   "frmDownLoad.frx":0894
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      ToolTipText     =   "icon2"
      Top             =   1785
      Width           =   510
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   0
      Left            =   2415
      MouseIcon       =   "frmDownLoad.frx":0B9E
      MousePointer    =   99  'Custom
      Picture         =   "frmDownLoad.frx":0EA8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      ToolTipText     =   "icon1"
      Top             =   1785
      Width           =   510
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   360
      Index           =   1
      Left            =   210
      ScaleHeight     =   300
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   735
      Width           =   4425
      Begin VB.CheckBox chkPrg 
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
   End
   Begin VB.Label lblProtocol 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   210
      TabIndex        =   10
      Top             =   1155
      Width           =   540
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
      ForeColor       =   &H8000000E&
      Height          =   225
      Index           =   0
      Left            =   420
      TabIndex        =   5
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
      TabIndex        =   6
      Top             =   65
      Width           =   3795
   End
   Begin VB.Image imgMin 
      Height          =   225
      Left            =   4305
      MouseIcon       =   "frmDownLoad.frx":11B2
      MousePointer    =   99  'Custom
      Top             =   50
      Width           =   225
   End
   Begin VB.Image imgClose 
      Height          =   225
      Left            =   4515
      MouseIcon       =   "frmDownLoad.frx":14BC
      MousePointer    =   99  'Custom
      Top             =   50
      Width           =   225
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   210
      TabIndex        =   4
      Top             =   420
      Width           =   4380
   End
   Begin VB.Label lblPerc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0% Completed"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   105
      TabIndex        =   3
      Top             =   1260
      Visible         =   0   'False
      Width           =   4590
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Timer"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   210
      TabIndex        =   2
      Top             =   1155
      Width           =   4425
   End
   Begin VB.Image imgBackDrop 
      Height          =   1605
      Left            =   0
      Picture         =   "frmDownLoad.frx":17C6
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
Attribute VB_Name = "frmDownLoad"
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
Dim iFTPfool As Integer
Dim bFTP As Boolean
Public Sub StartDownload(FileName As String, Url As String, iGridLine As Integer)

On Error GoTo Error_H

    'Initialize the Interface
    GridLine = iGridLine
    
    txtTargetFile = GetFileSave(FileName, GetExtensionType(FileName))
    lblCaption(0) = txtTargetFile
    lblCaption(1) = txtTargetFile
    
    If DownCTL.GetProtocol(Url) = icFTP Then
        bFTP = True ' We need this too fool around with the status bar
        lblProtocol = "FTP"
    Else
        lblProtocol = "HTTP"
    End If
    
    'Set up connection details
    DownCTL.Url = Url
    
    If GetSetting("BlueByte", "Connection", "Type", "Proxy") = "Proxy" Then 'Proxy
        DownCTL.Proxy = GetSetting("BlueByte", "Connection", "ProxyAddr", "")
        DownCTL.UserName = GetSetting("BlueByte", "Connection", "ProxyUsr", "")
        DownCTL.Password = GetSetting("BlueByte", "Connection", "ProxyPass", "")
    Else 'Modem
        DownCTL.Proxy = ""
        DownCTL.UserName = ""
        DownCTL.Password = ""
    End If
    
    DownCTL.TargetFile = txtTargetFile
    DownCTL.DownloadFile
    
    Me.Show
        DoEvents
    
    If GetSetting("BlueByte", "Connection", "dWindow", 0) <> 0 Then
        Me.Visible = False
    End If
    
    
    DoEvents
    Exit Sub
    
Error_H:
    frmMain.grdFiles.Row = GridLine
    frmMain.grdFiles.TextMatrix(GridLine, 5) = "Error : " & Err.Description
    sError = Err.Description
    frmMain.grdFiles.Row = GridLine
    frmMain.grdFiles.Col = 0
    Set frmMain.grdFiles.CellPicture = frmMain.imgError
    frmMain.grdFiles.CellPictureAlignment = flexAlignCenterCenter
    Unload Me

End Sub

Private Sub DownCtl_AsyncReadComplete(AsyncProp As AsyncProperty)
MsgBox "async read complete"
End Sub

Private Sub DownCtl_DownloadComplete()
'The deed is done !
bCompleted = True
Unload Me
End Sub

Private Sub DownCTL_DownloadError(sErrorD As String)
    frmMain.grdFiles.Row = GridLine
    frmMain.grdFiles.TextMatrix(GridLine, 5) = "Error : " & sErrorD
    sError = sErrorD
    frmMain.grdFiles.Row = GridLine
    frmMain.grdFiles.Col = 0
    Set frmMain.grdFiles.CellPicture = frmMain.imgError
    frmMain.grdFiles.CellPictureAlignment = flexAlignCenterCenter
    Unload Me
End Sub

Private Sub Form_Load()

Me.Width = imgBackDrop.Width
Me.Height = imgBackDrop.Height

'Systray
    t.cbSize = Len(t)
    t.hWnd = Picture1(0).hWnd
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = Picture1(0).Picture
    t.szTip = txtTargetFile & Chr$(0)
    dada NIM_ADD, t
    App.TaskVisible = False
'End of Systray

frmMain.lblAvg = "Activated"
iActiveForms = iActiveForms + 1

'Position the Download Form
If frmMain.Left >= Screen.Width / 2 Then
    Me.Left = frmMain.Left - Me.Width
Else
    Me.Left = frmMain.Left + frmMain.Width
End If

'If frmMain.Top >= Screen.Height / 2 Then
    Me.Top = frmMain.Top + (Me.Height * iActiveForms) - Me.Height
'Else
    'Me.Left = frmMain.Left + frmMain.Width
'End If

'Me.Top = frmMain.Top + frmMain.Height
Me.Show
DoEvents

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Timer1.Enabled = False
    t.cbSize = Len(t)
    t.hWnd = Picture1(0).hWnd
    t.uId = 1&
    dada NIM_DELETE, t
End Sub

Private Sub Form_Unload(Cancel As Integer)

'Do some stuff on the exiting of the form to be displayed on the Main form.
frmMain.grdFiles.Row = GridLine
If bCompleted Then
    frmMain.grdFiles.TextMatrix(GridLine, 5) = "Completed"
    frmMain.grdFiles.Col = 0
    Set frmMain.grdFiles.CellPicture = frmMain.imgDone
    frmMain.grdFiles.TextMatrix(GridLine, 7) = Date & " - " & Time
    DoEvents
Else
    If sError = "" Then
        frmMain.grdFiles.TextMatrix(GridLine, 5) = "Stopped"
        frmMain.grdFiles.Col = 0
        Set frmMain.grdFiles.CellPicture = frmMain.imgStop
        DoEvents
    Else
        frmMain.grdFiles.TextMatrix(GridLine, 5) = "Error : " & sError
        frmMain.grdFiles.Col = 0
        Set frmMain.grdFiles.CellPicture = frmMain.imgError
        DoEvents
    End If
End If

iActiveForms = iActiveForms - 1

'If iActiveForms < 1 Then
'    frmMain.tmrSpeed.Enabled = False
'    frmMain.optSpeed.Visible = False
'    frmMain.lblAvg = "Not Active"
'    TotalBytes = 0
'    TotalSeconds = 0
'End If

    Timer1.Enabled = False
    t.cbSize = Len(t)
    t.hWnd = Picture1(0).hWnd
    t.uId = 1&
    dada NIM_DELETE, t

DownCTL.Cancel
DoEvents

'Call frmMain.PositionForms
DoEvents
End Sub

Private Sub imgBackDrop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Call DragForm(Me)
End If
End Sub

Private Sub imgClose_Click()
    frmMain.grdFiles.Row = GridLine
    frmMain.grdFiles.TextMatrix(GridLine, 5) = "Cancelled"
    sError = "Cancelled"
    frmMain.grdFiles.Row = GridLine
    frmMain.grdFiles.Col = 0
    Set frmMain.grdFiles.CellPicture = frmMain.imgError
    frmMain.grdFiles.CellPictureAlignment = flexAlignCenterCenter
    Unload Me
    DownCTL.Cancel
    DoEvents
    Unload Me
End Sub

Private Sub imgMin_Click()
Me.Hide
End Sub
Private Sub DownCtl_AsyncReadProgress(AsyncProp As AsyncProperty)
On Error GoTo Error_H:
Dim status                As String
  
    Select Case AsyncProp.StatusCode
      Case vbAsyncStatusCodeConnecting
        status = "connecting"
      Case vbAsyncStatusCodeEndDownloadData
        status = "download complete"
      Case vbAsyncStatusCodeBeginDownloadData
        status = "Begin download"
      Case vbAsyncStatusCodeDownloadingData
        status = "Downloading..."
      Case vbAsyncStatusCodeFindingResource
        status = "Finding resource"
      Case vbAsyncStatusCodeMIMETypeAvailable
        status = "MIME type"
      Case vbAsyncStatusCodeSendingRequest
        status = "Sending request"
    End Select
  
    Progress AsyncProp.BytesRead, AsyncProp.BytesMax, status, , (AsyncProp.BytesRead - PreviousBytes)
    PreviousBytes = AsyncProp.BytesRead
    DoEvents
    DoEvents
Exit Sub
    
Error_H:
    
    frmMain.grdFiles.Row = GridLine
    frmMain.grdFiles.TextMatrix(GridLine, 5) = "Error : " & Err.Description
    sError = Err.Description
    frmMain.grdFiles.Row = GridLine
    frmMain.grdFiles.Col = 0
    Set frmMain.grdFiles.CellPicture = frmMain.imgError
    frmMain.grdFiles.CellPictureAlignment = flexAlignCenterCenter
    Unload Me
    DownCTL.Cancel
    DoEvents
    Unload Me
    
End Sub
Private Sub DownCtl_StateChanged(State As Integer)
 
 On Error GoTo Error_H
 
Select Case State

    Case 0
        lblHeader = "No state information available"
        DoEvents
        
    Case 1
        lblHeader = "Looking up IP address for remote server"
        DoEvents
        
    Case 2
        lblHeader = "Found IP Address for remote server"
        DoEvents
        
    Case 3
        lblHeader = "Connecting to remote server"
        DoEvents
        
    Case 4
        lblHeader = "Connected to remote server"
        DoEvents
        
    Case 5
        lblHeader = "Requesting information from remote server"
        DoEvents
        
    Case 6
        lblHeader = "Request sent successfully to remote server"
        DoEvents
        
    Case 6
        lblHeader = "Request sent successfully to remote server"
        DoEvents
        
    Case 7
        lblHeader = "Receiving response from remote server"
        DoEvents
        
    Case 9
        lblHeader = "Disconnecting from remote server"
        DoEvents
        
    Case 10
        lblHeader = "Disconnected from remote server"
        DoEvents
        
    Case 11
        lblHeader = "Error communicating with remote server"
        DoEvents
    

   Case icResponseReceived ' 12
        lblHeader = "Response received from remote server"
        DoEvents
End Select
   
DoEvents

Exit Sub
    
Error_H:
    frmMain.grdFiles.Row = GridLine
    frmMain.grdFiles.TextMatrix(GridLine, 5) = "Error : " & Err.Description
    sError = Err.Description
    frmMain.grdFiles.Row = GridLine
     frmMain.grdFiles.Col = 0
    Set frmMain.grdFiles.CellPicture = frmMain.imgError
    frmMain.grdFiles.CellPictureAlignment = flexAlignCenterCenter
    Unload Me
End Sub


Private Sub lblCaption_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    Call DragForm(Me)
End If

End Sub



Public Function Progress(Value, MaxValue, Optional HeaderX As String, Optional color As ColorConstants, Optional lSize As Long)
'' This is the actual progress bar function.

'On Error GoTo error_h

DoEvents
Dim Perc
Dim bb As Integer
chkPrg.Visible = True

'Get a color to do it in
If color = 0 Then color = vbWhite

'Display the header , if any was returned
If HeaderX <> "" Then
    lblHeader = HeaderX
Else
    lblHeader = "Busy Processing...Please wait"
End If

'Show the amount of bytes that is done !
If Value > 1000 Then
        lblTime = "Downloaded : " & Format((Value / 1000), "0.00") & " kb"
        frmMain.grdFiles.TextMatrix(GridLine, 3) = Format((Value / 1000), "0.00") & " kb"
        If Value > 1000000 Then
            lblTime = "Downloaded : " & Format((Value / 1000000), "0.00") & "mb"
            frmMain.grdFiles.TextMatrix(GridLine, 3) = Format((Value / 1000000), "0.00") & " mb"
        End If
    Else
        lblTime = "Downloaded : " & Value & " bytes"
        frmMain.grdFiles.TextMatrix(GridLine, 3) = Format((Value / 1000), "0.00") & " b"
    End If

If MaxValue = 0 Then
    

Else
    'Now work out the percentage (0-100) of where we currently are
    Perc = (Value / MaxValue) * 100
    If Perc < 0 Then Perc = 0
    If Perc > 100 Then Perc = 100
    Perc = Int(Perc)
    If Perc > 30 Then ' Display sux if the bar is too short---hack it a bit
        chkPrg.Caption = Int(Perc) & "% Completed" 'Just the Label Display
    Else
        chkPrg.Caption = Int(Perc) & "%" 'Just the Label Display
    End If
    frmMain.grdFiles.TextMatrix(GridLine, 2) = Perc & "%"
    
    'Calculate the Time Remaining
    If bTimeStarted = False Then
        bTimeStarted = True
        tTimeStarted = Time
    Else
        If Perc > 0 Then
            lTime = Time - tTimeStarted
            lTotalTime = (100 / Perc) * lTime
            lTimeLeft = lTotalTime - lTime
            lblTime = "Time Remaining : " & Format((lTimeLeft), "hh:mm:ss")
        End If
    End If
End If

'Fool the ftp transfer
If bFTP Then
    iFTPfool = iFTPfool + 1
    If iFTPfool > 1000 Then
        iFTPfool = 0
    End If
    Perc = iFTPfool / 10
End If

DoEvents
Dim dSpeed As Long
dSpeed = frmMain.SendSpeed(lSize)
DoEvents
DoEvents
DoEvents

chkPrg.Width = Int(Perc)
DoEvents
Exit Function
Error_H:
    MsgBox "Oops"
End Function


Private Sub mnuShow_Click()
Me.Visible = True
Me.Show
End Sub

Private Sub mnuStop_Click()
Unload Me
End Sub

Private Sub picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Hex(X) = "1E3C" Then
        Me.PopupMenu mnuTray
    End If
End Sub
