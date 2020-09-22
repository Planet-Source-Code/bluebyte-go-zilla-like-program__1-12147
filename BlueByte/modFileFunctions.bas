Attribute VB_Name = "modFileFunctions"
Option Explicit
Global gbCheck As Boolean

Public Function AddLink(txtUrl As String, GridX As MSFlexGrid) As Boolean
'Add the dragged link to the grid
Dim FileName As String
txtUrl = CleanURL(txtUrl)

GridX.AddItem Chr(9) & GetFileName(txtUrl) & Chr(9) & "" & Chr(9) & "" & Chr(9) & CleanURL(txtUrl) & Chr(9) & "File Added" & Chr(9) & Date & " - " & Time
GridX.Col = 0
GridX.Row = GridX.Rows - 1
Set GridX.CellPicture = frmMain.imgStop
GridX.CellPictureAlignment = flexAlignCenterCenter

End Function
Public Function CleanURL(StringX As String) As String

'Cleans the URL if it came form a stript
CleanURL = StringX
Dim pos As Integer
Dim ii As Integer
Dim bStart As Boolean
Dim iStart As Integer
Dim CleanString As String
pos = InStr(1, LCase(StringX), "javascript", vbTextCompare)

If pos > 0 Then
    For ii = 1 To Len(StringX)
    
        If bStart Then
            CleanString = CleanString & Mid(StringX, ii, 1)
        End If
    
        If Mid(StringX, ii, 1) = "'" Then
            If bStart = True Then
                bStart = False
                CleanString = Left(CleanString, Len(CleanString) - 1)
            Else
                bStart = True
            End If
        End If
        
    Next ii
CleanURL = CleanString
End If



End Function
Public Function GetFileName(StringX As String) As String

Dim ii As Integer
Dim pos As Integer

'Get the normal file name
GetFileName = "Undetected.html"

For ii = Len(StringX) To 1 Step -1
    If Mid(StringX, ii, 1) = "/" Or Mid(StringX, ii, 1) = "?" Or Mid(StringX, ii, 1) = "\" Then
        GetFileName = Right(StringX, Len(StringX) - ii)
        'Exit Function
        Exit For
    End If
Next ii

End Function

Public Function GetExtensionType(StringX As String) As String
Dim ii As Integer

For ii = Len(StringX) To 1 Step -1
    If Mid(StringX, ii, 1) = "." Then
        GetExtensionType = Right(StringX, Len(StringX) - ii)
    End If
Next ii


End Function
