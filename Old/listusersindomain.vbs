' Use ADO to search the domain for all users.
Set adoConnection = CreateObject("ADODB.Connection")
Set adoCommand = CreateObject("ADODB.Command")
adoConnection.Provider = "ADsDSOOBject"
adoConnection.Open "Active Directory Provider"
Set adoCommand.ActiveConnection = adoConnection

' Determine the DNS domain from the RootDSE object.
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")

' Filter on all users.
strFilter = "(&(objectCategory=person)(objectClass=user))"

strQuery = "<LDAP://" & strDNSDomain & ">;" & strFilter _
    & ";distinguishedName;subtree"

adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 100
adoCommand.Properties("Timeout") = 30
adoCommand.Properties("Cache Results") = False

' Enumerate all users. Write each user's Distinguished Name to the file.
Set adoRecordset = adoCommand.Execute
Do Until adoRecordset.EOF
    strDN = adoRecordset.Fields("distinguishedName").Value
    ' Escape any forward slash characters, "/", with the backslash
    ' escape character. All other characters that should be escaped are.
    strDN = Replace(strDN, "/", "\/")
    WScript.Echo strDN
    adoRecordset.MoveNext
Loop
adoRecordset.Close
