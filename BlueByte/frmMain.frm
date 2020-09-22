VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrGraph 
      Interval        =   1000
      Left            =   4200
      Top             =   5565
   End
   Begin VB.ListBox lstSpeed 
      Height          =   255
      Left            =   3885
      TabIndex        =   13
      Top             =   4935
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox picGraph 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      Height          =   645
      Left            =   210
      ScaleHeight     =   10
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   12
      Top             =   945
      Width           =   5580
      Begin VB.Label lblSpeed 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H8000000E&
         Height          =   225
         Left            =   4200
         TabIndex        =   15
         Top             =   0
         Width           =   1275
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   0
      Left            =   3675
      MouseIcon       =   "frmMain.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      ToolTipText     =   "icon1"
      Top             =   6090
      Width           =   510
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   1
      Left            =   4215
      MouseIcon       =   "frmMain.frx":0614
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":091E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      ToolTipText     =   "icon2"
      Top             =   6090
      Width           =   510
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5475
      Top             =   6150
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   5250
      Top             =   5355
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid grdFiles 
      Height          =   3750
      Left            =   210
      TabIndex        =   0
      Top             =   1575
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   6615
      _Version        =   393216
      Cols            =   4
      BackColorFixed  =   8388608
      ForeColorFixed  =   65535
      BackColorBkg    =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label lblAvg 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   210
      TabIndex        =   14
      Top             =   7560
      Width           =   5430
   End
   Begin VB.Label lblTray 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "System Tray"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4410
      MouseIcon       =   "frmMain.frx":0C28
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   5565
      Width           =   1275
   End
   Begin VB.Image imgButton 
      Height          =   360
      Index           =   4
      Left            =   4305
      Picture         =   "frmMain.frx":0F32
      Top             =   5460
      Width           =   1335
   End
   Begin VB.Label lblWhat 
      BackStyle       =   0  'Transparent
      Caption         =   " * Right click on the Grid to avtivate the Menu"
      Height          =   225
      Index           =   1
      Left            =   105
      TabIndex        =   8
      Top             =   5565
      Width           =   4005
   End
   Begin VB.Label lblWhat 
      BackStyle       =   0  'Transparent
      Caption         =   " * Drag a link from Internet Explorer into the Grid above"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   7
      Top             =   5355
      Width           =   4005
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3045
      MouseIcon       =   "frmMain.frx":2896
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   600
      Width           =   1275
   End
   Begin VB.Image imgDown 
      Height          =   180
      Left            =   1470
      Picture         =   "frmMain.frx":2BA0
      Top             =   6195
      Width           =   345
   End
   Begin VB.Image imgDone 
      Height          =   180
      Left            =   945
      Picture         =   "frmMain.frx":2F44
      Top             =   6195
      Width           =   345
   End
   Begin VB.Image imgError 
      Height          =   180
      Left            =   525
      Picture         =   "frmMain.frx":32E8
      Top             =   6195
      Width           =   345
   End
   Begin VB.Image imgStop 
      Height          =   180
      Left            =   0
      Picture         =   "frmMain.frx":368C
      Top             =   6195
      Width           =   345
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
      Height          =   330
      Index           =   0
      Left            =   525
      TabIndex        =   4
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
      Height          =   240
      Index           =   1
      Left            =   570
      TabIndex        =   5
      Top             =   140
      Width           =   4950
   End
   Begin VB.Label cmdExit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4410
      MouseIcon       =   "frmMain.frx":3A30
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   600
      Width           =   1275
   End
   Begin VB.Label lblSettings 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1575
      MouseIcon       =   "frmMain.frx":3D3A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   600
      Width           =   1275
   End
   Begin VB.Label lblConnect 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Connect"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   210
      MouseIcon       =   "frmMain.frx":4044
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   600
      Width           =   1275
   End
   Begin VB.Image imgButton 
      Height          =   360
      Index           =   3
      Left            =   4410
      Picture         =   "frmMain.frx":434E
      Top             =   525
      Width           =   1335
   End
   Begin VB.Image imgButton 
      Height          =   360
      Index           =   2
      Left            =   3045
      Picture         =   "frmMain.frx":5CB2
      Top             =   525
      Width           =   1335
   End
   Begin VB.Image imgButton 
      Height          =   360
      Index           =   1
      Left            =   1575
      Picture         =   "frmMain.frx":7616
      Top             =   525
      Width           =   1335
   End
   Begin VB.Image imgButton 
      Height          =   360
      Index           =   0
      Left            =   210
      Picture         =   "frmMain.frx":8F7A
      Top             =   525
      Width           =   1335
   End
   Begin VB.Image imgMin 
      Height          =   225
      Left            =   5480
      MouseIcon       =   "frmMain.frx":A8DE
      MousePointer    =   99  'Custom
      Top             =   50
      Width           =   225
   End
   Begin VB.Image imgClose 
      Height          =   225
      Left            =   5710
      MouseIcon       =   "frmMain.frx":ABE8
      MousePointer    =   99  'Custom
      Top             =   50
      Width           =   225
   End
   Begin VB.Image imgBackDrop 
      Height          =   6000
      Left            =   0
      Picture         =   "frmMain.frx":AEF2
      Top             =   0
      Width           =   6000
   End
   Begin VB.Menu mnuGrid 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuDownLoad 
         Caption         =   "Download"
      End
      Begin VB.Menu mnuTrash 
         Caption         =   "Trash"
      End
      Begin VB.Menu xx 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowColumn 
         Caption         =   "Show Column Value"
      End
      Begin VB.Menu xxx 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "TrayMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMain 
         Caption         =   "Main Window"
      End
      Begin VB.Menu mnuExit2 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
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

Dim iLastRow As Integer
Dim iLastCol As Integer


Dim dAvg As Double
Dim colForms As New Collection

Public Function SendSpeed(lSize As Long) As Long
TotalBytes = TotalBytes + lSize
SpeedBytes = SpeedBytes + lSize
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Clear Systray
    Timer1.Enabled = False
    t.cbSize = Len(t)
    t.hWnd = Picture1(0).hWnd
    t.uId = 1&
    dada NIM_DELETE, t
End Sub

Private Sub lblHelp_Click()
frmHelp.Show
End Sub

Private Sub lblTray_Click()
Me.Hide
End Sub

Private Sub mnuExit2_Click()
Unload Me
End Sub

Private Sub mnuMain_Click()
Me.Show
End Sub

Private Sub tmrGraph_Timer()

On Error GoTo Error_H

SpeedBytes = SpeedBytes / 2
lstSpeed.AddItem Format((SpeedBytes / 1000), "00.00")
lblSpeed = (Format((SpeedBytes / 1000), "00.00") / 1) & " kb/sec"

If lstSpeed.ListCount > 21 Then
    lstSpeed.RemoveItem 0
End If

picGraph.Line (0, 0)-(100, 10), &H800000, BF
picGraph.Refresh

Dim ii As Integer
For ii = 0 To lstSpeed.ListCount - 1
    If Not IsNumeric(lstSpeed.List(ii)) Then
        lstSpeed.List(ii) = 0
    End If
    
    If CLng(lstSpeed.List(ii)) > 0 Then
        picGraph.Line (100 - (ii * 5), 9)-(100 - (ii * 5) + 4, 10), &HFFFFC0, BF
    End If
    If CLng(lstSpeed.List(ii)) > 1 Then
        picGraph.Line (100 - (ii * 5), 8)-(100 - (ii * 5) + 4, 9), &HFFFF80, BF
    End If
    If CLng(lstSpeed.List(ii)) > 2 Then
        picGraph.Line (100 - (ii * 5), 7)-(100 - (ii * 5) + 4, 8), &HFFFF00, BF
    End If
    If CLng(lstSpeed.List(ii)) > 3 Then
        picGraph.Line (100 - (ii * 5), 6)-(100 - (ii * 5) + 4, 7), &HC0C0FF, BF
    End If
    If CLng(lstSpeed.List(ii)) > 4 Then
        picGraph.Line (100 - (ii * 5), 5)-(100 - (ii * 5) + 4, 6), &H8080FF, BF
    End If
    If CLng(lstSpeed.List(ii)) > 5 Then
        picGraph.Line (100 - (ii * 5), 4)-(100 - (ii * 5) + 4, 5), &HFF&, BF
    End If
    If CLng(lstSpeed.List(ii)) > 6 Then
        picGraph.Line (100 - (ii * 5), 3)-(100 - (ii * 5) + 4, 4), &HFFC0FF, BF
    End If
    If CLng(lstSpeed.List(ii)) > 7 Then
        picGraph.Line (100 - (ii * 5), 2)-(100 - (ii * 5) + 4, 3), &HFF80FF, BF
    End If
    If CLng(lstSpeed.List(ii)) > 8 Then
        picGraph.Line (100 - (ii * 5), 1)-(100 - (ii * 5) + 4, 2), &HFF00FF, BF
    End If
    If CLng(lstSpeed.List(ii)) > 9 Then
        picGraph.Line (100 - (ii * 5), 0)-(100 - (ii * 5) + 4, 1), &HFF00&, BF
    End If
    
    'picGraph.Line (100 - (ii * 5), (20 - CLng(lstSpeed.List(ii) * 2)))-(100 - (ii * 5) + 4, 20), vbYellow, BF
Next ii

SpeedBytes = 0

Exit Sub

Error_H:
    MsgBox "Oops"
End Sub

Private Sub tmrSpeed_Timer()

On Error GoTo Error_H

TotalSeconds = TotalSeconds + 1
dAvg = TotalBytes / TotalSeconds
dAvg = dAvg / 1000
dAvg = Format(dAvg, "00.00")
optSpeed.Width = dAvg
'optSpeed.BackColor = RGB(255, 255, 255 - (dAvg * 10))
lblAvg = dAvg & " KB / sec"



End Sub

Private Sub cmdExit_Click()
'End the application
    Unload Me
    DoEvents
    End
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'A file was dropped into the form
If Data.GetFormat(vbCFText) Then  'if text
        gbCheck = AddLink(Data.GetData(vbCFText), grdFiles)
End If
End Sub

Private Sub mnuDownload_Click()

'optSpeed.Visible = True
'optSpeed.Width = 0
'tmrSpeed.Enabled = True


'Download requested
grdFiles.Row = iLastRow
grdFiles.Col = 0
Set grdFiles.CellPicture = frmMain.imgDown
grdFiles.CellPictureAlignment = flexAlignCenterTop

Dim frmX As New frmDownLoad
colForms.Add frmX
Call colForms(colForms.Count).StartDownload(grdFiles.TextMatrix(iLastRow, 1), grdFiles.TextMatrix(iLastRow, 4), iLastRow)
'Call frmX.StartDownload(grdFiles.TextMatrix(iLastRow, 1), grdFiles.TextMatrix(iLastRow, 4), iLastRow)
Set frmX = Nothing
End Sub

Private Sub Form_Load()

lstSpeed.Clear
Dim ii As Integer
For ii = 1 To 20
    lstSpeed.AddItem "0"
Next ii

tmrGraph.Enabled = True

Me.Visible = False
Unload frmSplash
DoEvents

If GetSetting("BlueByte", "Connection", "Type", "Proxy") = "Proxy" Then
    frmGetConnected.StartConnection
    lblConnect.Caption = "Refresh Proxy"
Else
    If GetSetting("BlueByte", "Connection", "Type", "Proxy") = "Modem" Then
    Else
        frmConnection.Show vbModal
    End If
    lblConnect.Caption = ""
End If


'Position the form on the screen

Me.Top = GetSetting("BlueByte", "Forms", "Top", Me.Top = Screen.Height / 2 - Me.Height / 2)
Me.Left = GetSetting("BlueByte", "Forms", "Left", Me.Left = Screen.Width / 2 - Me.Width / 2)
If Me.Top < 0 Or Me.Top > Screen.Height Then Me.Top = 0
If Me.Left < 0 Or Me.Left > Screen.Width Then Me.Left = 0

'Systray
    t.cbSize = Len(t)
    t.hWnd = Picture1(0).hWnd
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = Picture1(0).Picture
    t.szTip = "BlueByte" & Chr$(0)
    dada NIM_ADD, t
    Timer1.Enabled = True
    App.TaskVisible = False
'End of Systray
    DoEvents
    InitGrid
    'Me.Hide
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Clear Systray
Timer1.Enabled = False
t.cbSize = Len(t)
t.hWnd = Picture1(0).hWnd
t.uId = 1&
dada NIM_DELETE, t
    
'Save the Form position to the registry
SaveSetting "BlueByte", "Forms", "Top", Me.Top
SaveSetting "BlueByte", "Forms", "Left", Me.Left

'Save the data in the grid to a text file
gbCheck = SaveGrid(grdFiles)

End Sub
Private Sub Timer1_Timer()
    Static i As Long, img As Long
    t.cbSize = Len(t)
    t.hWnd = Picture1(0).hWnd
    t.uId = 1&
    t.uFlags = NIF_ICON
    t.hIcon = Picture1(i).Picture
    dada NIM_MODIFY, t
    Timer1.Enabled = True
    i = i + 1
    If i = 2 Then i = 0
End Sub
Private Sub picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Hex(X) = "1E3C" Then
        Me.PopupMenu mnuTray
    End If
    'Me.PopupMenu mnuTray
End Sub
Private Sub grdFiles_DblClick()
'Start a download
Call mnuDownload_Click

End Sub

Private Sub grdFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Save the row and column positions for referencing
iLastRow = grdFiles.MouseRow
iLastCol = grdFiles.MouseCol

If Button = 2 Then
    PopupMenu mnuGrid
End If

End Sub

Private Sub grdFiles_OLEDragDrop(Data As MSFlexGridLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

'A file was dropped into the grid
If Data.GetFormat(vbCFText) Then  'if text
        gbCheck = AddLink(Data.GetData(vbCFText), grdFiles)
End If

End Sub

Private Sub imgBackDrop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Call DragForm(Me)
End If
End Sub

Private Sub imgClose_Click()
Unload Me
End Sub

Private Sub imgMin_Click()
'Save the Form position to the registry
SaveSetting "BlueByte", "Forms", "Top", Me.Top
SaveSetting "BlueByte", "Forms", "Left", Me.Left
Me.Hide
End Sub
Public Function InitGrid()

'Set up the initial Flexgrid Layout.
With grdFiles
    .Rows = 1
    .Cols = 8
    .Col = 0
    .Row = 0
    .Text = ""
    .Col = 1
    .Text = "Name"
    .Col = 2
    .Text = " % Done"
    .Col = 3
    .Text = "Size Done"
    .Col = 4
    .Text = "Url"
    .Col = 5
    .Text = "Status"
    .Col = 6
    .Text = "Date Added"
    .Col = 7
    .Text = "Date Completed"
    
    .ColWidth(0) = 500
    .ColWidth(1) = 3000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1000
    .ColWidth(4) = 4000
    .ColWidth(5) = 2000
    .ColWidth(6) = 2000
    .ColWidth(7) = 2000
End With

'Load the grid from the saved text file
gbCheck = LoadGrid(grdFiles)

End Function


Private Sub lblConnect_Click()
If lblConnect.Caption = "Refresh Proxy" Then
    Me.Visible = False
    frmGetConnected.StartConnection
End If
End Sub

Private Sub lblHeader_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Call DragForm(Me)
End If
End Sub

Private Sub lblSettings_Click()

'Show the settings form
frmConnection.Show vbModal

End Sub
Public Function SaveGrid(GridX As MSFlexGrid, Optional sFileName As String) As Boolean

'Saves the lines in the FlexGrid to a Text File
Dim fno As Integer
Dim fname As String
Dim aa As Integer
Dim bb As Integer
Dim sLine As String

If GridX.Rows < 2 Then Exit Function
If GridX.Cols < 2 Then Exit Function

'obtain the next free file handle from the system
fno = FreeFile
If sFileName = "" Then
     sFileName = App.Path & "/Grid.txt"
End If

'open and save the textbox to a file
Open sFileName For Output As #fno 'Open the Text File

    For aa = 1 To GridX.Rows - 1 'Loop the Lines
        For bb = 0 To GridX.Cols - 1 'Loop the Columns
            sLine = sLine & GridX.TextMatrix(aa, bb) & Chr(9) ' Build the Text File
        Next bb
        sLine = sLine & vbCrLf
    
    Next aa
    Print #fno, (sLine) ' Save the Text File
    
Close #fno 'Close the Text File

End Function

Public Function LoadGrid(GridX As MSFlexGrid, Optional sFileName As String) As Boolean

'Load the Grid from the Text File
Dim fno As Integer
Dim fname As String
Dim TempLine As String
Dim Mycheck

    'obtain the next free file handle from the system
    fno = FreeFile
    If sFileName = "" Then
         sFileName = App.Path & "/Grid.txt"
    End If
    
    Mycheck = Dir(sFileName) ' Check if Text File Exists
    
    If Mycheck = "" Then Exit Function

    'load the file into the textbox
    Open sFileName For Input As #fno 'Open the Text File
    Do While Not EOF(fno)
        Input #fno, TempLine ' Red the Lines
        GridX.AddItem Chr(9) & TempLine ' Update the Grid
        'Set the picture one the grid
        GridX.Col = 0
        GridX.Row = GridX.Rows - 1
        GridX.CellPictureAlignment = flexAlignCenterCenter
        
        If Trim(GridX.TextMatrix(GridX.Row, 4)) <> "" Then
            If GridX.TextMatrix(GridX.Row, 5) = "Completed" Then
                Set GridX.CellPicture = frmMain.imgDone
            Else
                If Mid(GridX.TextMatrix(GridX.Row, 5), 1, 5) = "Error" Then
                    Set GridX.CellPicture = frmMain.imgError
                Else
                    Set GridX.CellPicture = frmMain.imgStop
                End If
            End If
        Else
            If GridX.Row > 1 Then ' Cannot remove the fixed row....of course !
                grdFiles.RemoveItem GridX.Row ' Remove the blank lines
            End If
        End If
    Loop
    
    Close #fno 'Close the Text File
    
End Function

Private Sub mnuShowColumn_Click()

Dim Result
Result = grdFiles.TextMatrix(iLastRow, iLastCol)
Result = InputBox("Change the Column Value", "BlueByte", Result)

If Result = "" Then
Else
    grdFiles.TextMatrix(iLastRow, iLastCol) = Result
End If

End Sub

Private Sub mnuTrash_Click()
'Delete a line on the Grid
Dim ii As Integer

If iLastRow > 1 Then
    grdFiles.RemoveItem iLastRow
Else
    For ii = 0 To grdFiles.Cols - 1
        grdFiles.TextMatrix(iLastRow, ii) = ""
    Next ii
End If

End Sub
Public Sub PositionForms()
Dim ii As Integer
For ii = 1 To colForms.Count

    If frmMain.Left >= Screen.Width / 2 Then
    colForms(ii).Left = frmMain.Left - colForms(ii).Width
Else
    colForms(ii).Left = frmMain.Left + frmMain.Width
End If

'If frmMain.Top >= Screen.Height / 2 Then
    colForms(ii).Top = frmMain.Top + (Me.Height * ii) - colForms(ii).Height

Next ii

DoEvents
End Sub
