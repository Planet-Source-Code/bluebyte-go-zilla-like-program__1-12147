Attribute VB_Name = "modMain"
Option Explicit

'For Dragging Borderless Forms...
Public Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReleaseCapture Lib "USER32" () As Long

Global iActiveForms As Integer
Global SpeedBytes As Long
Global TotalBytes As Long
Global TotalSeconds As Long

Global Const WM_NCLBUTTONDOWN = &HA1
Global Const HTCAPTION = 2

Public Sub DragForm(frm As Form)

On Local Error Resume Next

'Move the borderless form...
Call ReleaseCapture
Call SendMessage(frm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)

End Sub
Public Function GetFileSave(sDefault As String, sExtension As String) As String

Dim sTargetFile As String
Dim sSourceFile As String
Dim Mycheck


frmMain.dlg.Filter = sExtension & " Files(*." & sExtension & ")|*." & sExtension & "|All Files (*.*)|*.*"
frmMain.dlg.DefaultExt = sExtension
frmMain.dlg.FileName = sDefault
frmMain.dlg.ShowSave
sTargetFile = frmMain.dlg.FileName

If sTargetFile = "" Then Exit Function

Mycheck = Dir(sTargetFile)

If Mycheck = frmMain.dlg.FileTitle Then
    Dim Result
    Result = MsgBox("This file already exists. Overrite ?", vbYesNo, "File Exists")
    If Result = vbYes Then
        Kill sTargetFile
        DoEvents
    Else
        Exit Function
    End If
End If

GetFileSave = frmMain.dlg.FileTitle

End Function
