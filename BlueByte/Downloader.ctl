VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.UserControl Downloader 
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   ScaleHeight     =   510
   ScaleWidth      =   495
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4095
      Top             =   1155
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "Downloader.ctx":0000
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "Downloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:
Const m_def_TargetFile = 0
Const m_def_SourceURL = 0
Const m_def_FileSize = 0
'Property Variables:
Dim m_TargetFile As Variant
Dim m_SourceURL As Variant
Dim m_FileSize As Variant
'Event Declarations:
Event AsyncReadComplete(AsyncProp As AsyncProperty) 'MappingInfo=UserControl,UserControl,-1,AsyncReadComplete
Event AsyncReadProgress(AsyncProp As AsyncProperty) 'MappingInfo=UserControl,UserControl,-1,AsyncReadProgress
Event DownloadError(sError As String)
Event DownloadComplete()
Event StateChanged(State As Integer) 'MappingInfo=Inet1,Inet1,-1,StateChanged
Attribute StateChanged.VB_Description = "StateChanged event"

Dim gbcheck As Boolean
Dim RealURL As String


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Inet1,Inet1,-1,Cancel
Public Sub Cancel()
Attribute Cancel.VB_Description = "Method used to cancel the request currently being executed"
    Inet1.Cancel
End Sub
'
'Public Sub CancelAsyncRead(Optional Property As Variant)
'
'End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Inet1,Inet1,-1,Execute
Public Sub Execute(Optional URL As Variant, Optional Operation As Variant, Optional InputData As Variant, Optional InputHdrs As Variant)
Attribute Execute.VB_Description = "Issue a request to the remote computer"
    Inet1.Execute URL, Operation, InputData, InputHdrs
End Sub

'Private Sub Inet1_StateChanged(State As Integer)
'    RaiseEvent StateChanged(State)
'End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Inet1,Inet1,-1,StillExecuting
Public Property Get StillExecuting() As Boolean
Attribute StillExecuting.VB_Description = "Returns whether this control is currently busy"
    StillExecuting = Inet1.StillExecuting
End Property

Public Function DownloadFile() As Variant

On Error GoTo Error_H:

Dim download() As Byte

 Inet1.Protocol = GetProtocol(RealURL)
' If Inet1.Protocol = icFTP Then
'    gbcheck = StartConnection
'    Inet1.URL = RealURL
'End If
DoEvents
'Start the download
'If Inet1.Protocol = icHTTP Then
'Me.AsyncRead Inet1.URL, vbAsyncTypeByteArray, "TESTER", vbAsyncReadForceUpdate
'Else


Me.AsyncRead RealURL, vbAsyncTypeByteArray, "TESTER", vbAsyncReadForceUpdate
'End If

DoEvents
Exit Function
Error_H:
RaiseEvent DownloadError("Error : In DownloadFile")

End Function

Public Property Get TargetFile() As Variant
    TargetFile = m_TargetFile
End Property

Public Property Let TargetFile(ByVal New_TargetFile As Variant)
    m_TargetFile = New_TargetFile
    PropertyChanged "TargetFile"
End Property

Public Property Get SourceURL() As Variant
    SourceURL = m_SourceURL
End Property

Public Property Let SourceURL(ByVal New_SourceURL As Variant)
    m_SourceURL = New_SourceURL
    PropertyChanged "SourceURL"
End Property

Public Property Get FileSize() As Variant
    FileSize = m_FileSize
End Property

Public Property Let FileSize(ByVal New_FileSize As Variant)
    m_FileSize = New_FileSize
    PropertyChanged "FileSize"
End Property

Private Sub Inet1_StateChanged(ByVal State As Integer)
RaiseEvent StateChanged(State)

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_TargetFile = m_def_TargetFile
    m_SourceURL = m_def_SourceURL
    m_FileSize = m_def_FileSize
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_TargetFile = PropBag.ReadProperty("TargetFile", m_def_TargetFile)
    m_SourceURL = PropBag.ReadProperty("SourceURL", m_def_SourceURL)
    m_FileSize = PropBag.ReadProperty("FileSize", m_def_FileSize)
    Inet1.Password = PropBag.ReadProperty("Password", "")
    Inet1.Protocol = PropBag.ReadProperty("Protocol", 1)
    Inet1.Proxy = PropBag.ReadProperty("Proxy", "")
    Inet1.RemoteHost = PropBag.ReadProperty("RemoteHost", "")
    Inet1.RemotePort = PropBag.ReadProperty("RemotePort", 80)
    Inet1.RequestTimeout = PropBag.ReadProperty("RequestTimeout", 60)
    Inet1.UserName = PropBag.ReadProperty("UserName", "")
    Inet1.URL = PropBag.ReadProperty("URL", "")
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("TargetFile", m_TargetFile, m_def_TargetFile)
    Call PropBag.WriteProperty("SourceURL", m_SourceURL, m_def_SourceURL)
    Call PropBag.WriteProperty("FileSize", m_FileSize, m_def_FileSize)
    Call PropBag.WriteProperty("Password", Inet1.Password, "")
    Call PropBag.WriteProperty("Protocol", Inet1.Protocol, 1)
    Call PropBag.WriteProperty("Proxy", Inet1.Proxy, "")
    Call PropBag.WriteProperty("RemoteHost", Inet1.RemoteHost, "")
    Call PropBag.WriteProperty("RemotePort", Inet1.RemotePort, 80)
    Call PropBag.WriteProperty("RequestTimeout", Inet1.RequestTimeout, 60)
    Call PropBag.WriteProperty("UserName", Inet1.UserName, "")
    Call PropBag.WriteProperty("URL", Inet1.URL, "")
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,AsyncRead
Public Sub AsyncRead(Target As String, AsyncType As Long, Optional PropertyName As Variant, Optional AsyncReadOptions As Variant)
    
On Error GoTo Error_H

UserControl.AsyncRead Target, AsyncType, PropertyName, AsyncReadOptions

Exit Sub

Error_H:
RaiseEvent DownloadError("Error : In AsyncRead")

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,CancelAsyncRead
Public Sub CancelAsyncRead(Optional Property As Variant)
Attribute CancelAsyncRead.VB_Description = "Cancel an asynchronous data request."
    UserControl.CancelAsyncRead Property
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
On Error GoTo Error_H
    'Save the file once the download is completed
    Dim sfilename As String
    Dim b() As Byte
    
    sfilename = Me.TargetFile
    DoEvents
    If AsyncProp.AsyncType = vbAsyncTypeByteArray Then
        b = AsyncProp.Value
        sfilename = Me.TargetFile
        Open sfilename For Binary Access Write As #1
        Put #1, , b()
        Close #1
    End If
    DoEvents
    RaiseEvent DownloadComplete

Exit Sub

Error_H:
RaiseEvent DownloadError("Not Connected")
    
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)

On Error GoTo Error_H

    RaiseEvent AsyncReadProgress(AsyncProp)
    
Exit Sub

Error_H:
RaiseEvent DownloadError("Error : In AsyncReadProgress")

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Inet1,Inet1,-1,Password
Public Property Get Password() As String
Attribute Password.VB_Description = "Password to use for authentication"
    Password = Inet1.Password
End Property

Public Property Let Password(ByVal New_Password As String)
    Inet1.Password() = New_Password
    PropertyChanged "Password"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Inet1,Inet1,-1,Protocol
Public Property Get Protocol() As ProtocolConstants
Attribute Protocol.VB_Description = "Protocol to use for this URL"
    Protocol = Inet1.Protocol
End Property

Public Property Let Protocol(ByVal New_Protocol As ProtocolConstants)
    Inet1.Protocol() = New_Protocol
    PropertyChanged "Protocol"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Inet1,Inet1,-1,Proxy
Public Property Get Proxy() As String
Attribute Proxy.VB_Description = "Proxy server to use when accessing the net"
    Proxy = Inet1.Proxy
End Property

Public Property Let Proxy(ByVal New_Proxy As String)
    Inet1.Proxy() = New_Proxy
    PropertyChanged "Proxy"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Inet1,Inet1,-1,RemoteHost
Public Property Get RemoteHost() As String
Attribute RemoteHost.VB_Description = "Returns/Sets the remote computer"
    RemoteHost = Inet1.RemoteHost
End Property

Public Property Let RemoteHost(ByVal New_RemoteHost As String)
    Inet1.RemoteHost() = New_RemoteHost
    PropertyChanged "RemoteHost"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Inet1,Inet1,-1,RemotePort
Public Property Get RemotePort() As Integer
Attribute RemotePort.VB_Description = "Returns/Sets the internet port to be used on the remote computer"
    RemotePort = Inet1.RemotePort
End Property

Public Property Let RemotePort(ByVal New_RemotePort As Integer)
    Inet1.RemotePort() = New_RemotePort
    PropertyChanged "RemotePort"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Inet1,Inet1,-1,RequestTimeout
Public Property Get RequestTimeout() As Long
Attribute RequestTimeout.VB_Description = "Gets/Sets number of seconds to wait for request to complete"
    RequestTimeout = Inet1.RequestTimeout
End Property

Public Property Let RequestTimeout(ByVal New_RequestTimeout As Long)
    Inet1.RequestTimeout() = New_RequestTimeout
    PropertyChanged "RequestTimeout"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Inet1,Inet1,-1,UserName
Public Property Get UserName() As String
Attribute UserName.VB_Description = "User name to use for authentication"
    UserName = Inet1.UserName
End Property

Public Property Let UserName(ByVal New_UserName As String)
    Inet1.UserName() = New_UserName
    PropertyChanged "UserName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Inet1,Inet1,-1,URL
Public Property Get URL() As String
Attribute URL.VB_Description = "Returns/Sets the URL used by this control"
    URL = Inet1.URL
End Property

Public Property Let URL(ByVal New_URL As String)
    RealURL = New_URL
    Inet1.URL() = New_URL
    PropertyChanged "URL"
End Property
Public Function GetProtocol(sUrl As String) As String
Dim TempUrl As String
Dim bFtp As Boolean

TempUrl = Trim(UCase(sUrl))

If InStr(1, TempUrl, "FTP://", vbTextCompare) Then
    bFtp = True
End If

If InStr(1, TempUrl, "FTP.", vbTextCompare) Then
    bFtp = True
End If

If bFtp = True Then
    GetProtocol = icFTP ' 2
Else
    GetProtocol = icHTTP ' 4
End If

DoEvents

End Function
