'
' CLIENT WINSOCK VBSCRIPT
'
' NOTES: (FEBRUARY 19, 2007)
'
' Delays are required where they are located,
' or it sends data too quick, and errors.
' 
' Uses Port 80 by default.
' IP and Port are user settings.
'
' Creates a Log file.
' c:\WSClient.log
'
' this only sends basic text, no long essays or files.
' for that, it would require some minor but required changes.
'

Option Explicit
Dim winsock, SendData, RecieveData, secs

'****** CHANGE THESE SETTINGS *********

Const RemoteHost           = "10.125.17.138"
Const RemotePort           = 86

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

MsgBox "Client Started."
WriteData Now & " - Client Started"

'// CREATE WINSOCK
On Error Resume Next
Set winsock = Wscript.CreateObject("MSWINSOCK.Winsock", "winsock_")
If Err.Number <> 0 Then
    MsgBox "Winsock Object Error!" & vbCrLf & Hex(Err.Number) & ":" & Err.Description & vbCrLf & "Script will exit now."
    WriteData Now & " - Winsock Object Error." 
    WScript.Quit
End If
On Error Goto 0

'// CONNECT NOW
winsock.RemoteHost = RemoteHost
winsock.RemotePort = RemotePort 
winsock.Connect
        
'// MAIN DELAY - WAITS FOR CONNECTED STATE
'// SOCKET ERROR RAISES WINSOCK ERROR SUB
while winsock.State <> sckError And winsock.state <> sckConnected And winsock.state <> sckClosing And secs <> 25
    WScript.Sleep 1000  '// 1 sec delay in loop
    secs = secs + 1     '// wait 25 secs max
Wend

'// CONNECTION TIMED OUT
If secs > 24 Then
    MsgBox "Timed Out"
    WriteData Now & " - Client Timed Out"
    ClientClose()
End If

'// SEND DATA NOW
winsock.SendData "Test" 

'// WAIT INCASE NO RESPONSE
Wscript.Sleep 25000
WriteData Now & " - Client No Response. Winsock State: " & winsock.state
ClientClose()

'// WINSOCK DATA ARRIVES
Sub winsock_dataArrival(bytesTotal)
    Dim strData
    winsock.GetData strData, vbString
    RecieveData = strData 
    WriteData Now & " - Client Recieved: " & RecieveData
    winsock.SendData "Test"  
    WScript.Sleep 1000
    WriteData Now & " - Client Sent Data"
    ClientClose()
End Sub

'// WINSOCK ERROR
Sub winsock_Error(Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay)
    MsgBox "Cient Error: " & Number & vbCrLf & Description
    WriteData Now & " - Cient Error: " & Number & ". " & Description
    ClientClose()
End Sub

'// EXIT SCRIPT
Sub ClientClose()
    If winsock.state <> sckClosed Then winsock.Close
    Set winsock = Nothing
    WriteData Now & " - Client Closed."
    Wscript.Quit
End SUb

'// CREATE LOG ENTRY
Function WriteData(Data)
    Dim fso, file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile("C:\WSClient.log", 8, True)
    file.write Data & vbCrLf
    file.Close
    Set file = Nothing
    Set fso = Nothing
End Function