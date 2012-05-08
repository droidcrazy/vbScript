'**************************************************
'* CertExpiryCheck.vbs
'* Enumerate certificates with day left for expiry 
'**************************************************

Option Explicit

Const CAPICOM_ACTIVE_DIRECTORY_USER_STORE = 3
Const CAPICOM_CURRENT_USER_STORE = 2
Const CAPICOM_LOCAL_MACHINE_STORE = 1
Const CAPICOM_MEMORY_STORE = 0
Const CAPICOM_SMART_CARD_USER_STORE = 4
Const CAPICOM_CERTIFICATE_FIND_SUBJECT_NAME = 1        
Const CAPICOM_STORE_OPEN_READ_ONLY = 0

Dim Store, Certificates, Certificate


Set Store = CreateObject("CAPICOM.Store")
Store.Open CAPICOM_LOCAL_MACHINE_STORE, "MY" ,CAPICOM_STORE_OPEN_READ_ONLY
Set Certificates = Store.Certificates

If Certificates.Count >0 Then
   For Each Certificate in Certificates
    WScript.Echo "*** Subject " & Certificate.SubjectName & " ***"
    WScript.Echo "Issued by " & Certificate.IssuerName 
    WScript.Echo "Valid from " & Certificate.ValidFromDate & " to " & Certificate.ValidToDate 
    WScript.Echo "Days to expiry " & DateDiff("d",now(),Certificate.ValidToDate)
    WScript.Echo 
   Next
 Else
End If

Set Certificates = Nothing
Set Store = Nothing

Sub CommandUsage
  MsgBox "Usage: CertExpiryCheck.vbs  [SubjectName] ", vbInformation,"CertExpiryCheck"
  WScript.Quit(1)
End Sub