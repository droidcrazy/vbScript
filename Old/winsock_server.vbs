'
' SERVER WINSOCK VBSCRIPT
'
' NOTES: (FEBRUARY 19, 2007)
'
' Delays are required where they are located,
' or it sends data too quick, and errors.
' 
' Listens on Port 80 by default.
' Port is user setting.
'
' Creates a Log file.
' c:\WSServer.log
'
' this only recieves basic text, no long essays or files.
' for that, it would require some minor but required changes.
'

Option Explicit
Dim winsock

'****** CHANGE THESE SETTINGS *********

Const LocalPort            = 86

'***************************************

Const sckClosed            = 0 '// Default. Closed 
Const sckOpen              = 1 '// Open 
Const sckListening         = 2 '// Listening 
Const sckConnectionPending = 3 '// Connection pending 
Const sckResolvingHost     = 4 '// Resolving host 
Const sckHostResolved      = 5 '// Host resolved 
Const sckConnecting        = 6 '// Connecting 
Const sckConnected         = 7 '// Connected 
Const sckClosing           = 8 '// Peer is closing the connection 
Const sckError             = 9 '// Error 

MsgBox "Start Server"
WriteData Now & " - Server Started"

'// CREATE WINSOCK
On Error Resume Next
Set winsock = Wscript.CreateObject("MSWINSOCK.Winsock", "winsock_")
If Err.Number <> 0 Then
    MsgBox "Winsock Object Error!" & vbCrLf & "Script will exit now."
    WriteData Now & " - Winsock Object Error."
    WScript.Quit
End If
On Error Goto 0

'// LISTEN NOW
winsock.LocalPort = LocalPort
ServerListen

'// MAIN DELAY - INFINITE LOOP
'// SOCKET ERROR RAISES WINSOCK ERROR SUB
while winsock.State <> sckError
    WScript.Sleep 200
Wend

'// JUST INCASE
ServerClose()

'// WINSOCK CONNECT REQUEST
Sub winsock_ConnectionRequest(requestID)
    If winsock.State <> sckClosed Then
        winsock.Close
    End If
    winsock.Accept requestID
    WriteData Now & " - Server Requested ID: " & requestID
    winsock.SendData "Server Received okay"
    WScript.Sleep 1000  '// REQUIRED OR ERRORS
End Sub

'// WINSOCK DATA ARRIVES
Sub winsock_dataArrival(bytesTotal)
    Dim strData
    WriteData Now & " - Server Data Arrives"
    winsock.GetData strData, vbString
    WriteData Now & " - Server Recieved: " & strData 
    WScript.Sleep 2000  '// REQUIRED OR ERRORS
    ServerListen()
End Sub

'// WINSOCK ERROR
Sub winsock_Error(Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay)
    MsgBox "Server Error " & Number & vbCrLf & Description
    WriteData Now & " - Server Error: " & Number & ". " & Description
    ServerClose()
End Sub

'// LISTEN FOR REQUEST
Sub ServerListen()
    If winsock.State <> sckClosed Then
        WriteData Now & " - Server Closed (Listen)"
        winsock.Close
    End If
    WriteData Now & " - Server Listen"
    winsock.Listen
End SUb

'// EXIT SCRIPT
Sub ServerClose()
    If winsock.state <> sckClosed Then winsock.Close
    Set winsock = Nothing
    WriteData Now & " - Server Closed."
    Wscript.Quit
End SUb

'// CREATE LOG ENTRY
Function WriteData(Data)
    Dim fso, file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile("C:\WSServer.log", 8, True)
    file.write Data & vbCrLf
    file.Close
    Set file = Nothing
    Set fso = Nothing
End Function