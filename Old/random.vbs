Option Explicit
Dim Title, NumChar, Count, strRdm, intRdm
Title = "Random Password Generator"
On Error Resume Next
Do While Not IsNumeric(NumChar) or NumChar < 1 Or NumChar > 20
NumChar = InputBox("Enter a number between 1 and 20 to generate a " & _
                   "case sensitive string with that number of characters:", _
                   Title, 8)
If NumChar = "" Then WScript.Quit
If NOT IsNumeric(NumChar) Or NumChar < 1 Or NumChar > 20 Then MsgBox Chr(34) & NumChar & Chr(34) & " is invalid input." & vbcrlf & vbcrlf & "Input must be a number between 1 and 20",, Title Else NumChar = CInt(NumChar)
Loop
On Error GoTo 0

Randomize Timer

Do Until Count = NumChar
  Count = Count + 1
  GetRdm
  strRdm = strRdm & Chr(intRdm)
Loop

InputBox NumChar & " character case sensitive string:" & vbcrlf & vbcrlf & _
         "(Press Ctrl + C to copy results to Clipboard)", Title, strRdm

Sub GetRdm
  intRdm = Int((122 - 49) * Rnd + 48)
  If intRdm > 57 And intRdm < 65 Or intRdm > 90 And intRdm < 97 Then GetRdm
End Sub
