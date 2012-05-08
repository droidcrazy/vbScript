Dim ToAddress
Dim MessageSubject
Dim MessageBody
Dim MessageAttachment

Dim ol, ns, newMail

ToAddress = "becker, tommy"
MessageSubject = "test 1"
MessageBody = "testing"
MessageAttachment = "h:\changed.txt"

' connect to Outlook
Set ol = WScript.CreateObject("Outlook.Application")
Set ns = ol.getNamespace("MAPI")

Set newMail = ol.CreateItem(olMailItem)
newMail.Subject = MessageSubject
newMail.Body = MessageBody & vbCrLf

' validate the recipient, just in case...
Set myRecipient = ns.CreateRecipient(ToAddress)
myRecipient.Resolve
If Not myRecipient.Resolved Then
  MsgBox "Unknown recipient"
Else
  newMail.Recipients.Add(ToAddress)
  newMail.Attachments.Add(MessageAttachment).Displayname = "Check this out"
  newMail.Send
End If

Set ol = Nothing
