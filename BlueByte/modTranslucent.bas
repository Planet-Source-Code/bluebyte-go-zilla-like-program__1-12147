Attribute VB_Name = "modTranslucent"
Option Explicit




Public Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowRgn Lib "USER32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "GDI32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "USER32" () As Long
Public Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Public Const RGN_OR = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

'Use the to set ontop
Public Declare Function SetWindowPos Lib "USER32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
' SetWindowPos Flags
Global Const SWP_NOSIZE = &H1
Global Const SWP_NOMOVE = &H2
Global Const SWP_NOZORDER = &H4
Global Const SWP_NOREDRAW = &H8
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Global Const SWP_SHOWWINDOW = &H40
Global Const SWP_HIDEWINDOW = &H80
Global Const SWP_NOCOPYBITS = &H100
Global Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering

Global Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Global Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
' SetWindowPos() hwndInsertAfter values
Global Const HWND_TOP = 0
Global Const HWND_BOTTOM = 1
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

Public Function MakeRegion(picSkin As PictureBox) As Long
    
    ' Make a windows "region" based on a given picture box'
    ' picture. This done by passing on the picture line-
    ' by-line and for each sequence of non-transparent
    ' pixels a region is created that is added to the
    ' complete region. I tried to optimize it so it's
    ' fairly fast, but some more optimizations can
    ' always be done - mainly storing the transparency
    ' data in advance, since what takes the most time is
    ' the GetPixel calls, not Create/CombineRgn
    
    Dim X As Long, Y As Long, StartLineX As Long
    Dim FullRegion As Long, LineRegion As Long
    Dim TransparentColor As Long
    Dim InFirstRegion As Boolean
    Dim InLine As Boolean  ' Flags whether we are in a non-tranparent pixel sequence
    Dim hDC As Long
    Dim PicWidth As Long
    Dim PicHeight As Long
    
    hDC = picSkin.hDC
    PicWidth = picSkin.ScaleWidth
    PicHeight = picSkin.ScaleHeight
    
    InFirstRegion = True: InLine = False
    X = Y = StartLineX = 0
    
    ' The transparent color is always the color of the
    ' top-left pixel in the picture. If you wish to
    ' bypass this constraint, you can set the tansparent
    ' color to be a fixed color (such as pink), or
    ' user-configurable
    TransparentColor = GetPixel(hDC, 0, 0)
    
    For Y = 0 To PicHeight - 1
        For X = 0 To PicWidth - 1
            
            If GetPixel(hDC, X, Y) = TransparentColor Or X = PicWidth Then
                ' We reached a transparent pixel
                If InLine Then
                    InLine = False
                    LineRegion = CreateRectRgn(StartLineX, Y, X, Y + 1)
                    
                    If InFirstRegion Then
                        FullRegion = LineRegion
                        InFirstRegion = False
                    Else
                        CombineRgn FullRegion, FullRegion, LineRegion, RGN_OR
                        ' Always clean up your mess
                        DeleteObject LineRegion
                    End If
                End If
            Else
                ' We reached a non-transparent pixel
                If Not InLine Then
                    InLine = True
                    StartLineX = X
                End If
            End If
        Next
    Next
    
    MakeRegion = FullRegion
End Function

