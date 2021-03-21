'SearchAD.vbs
'On Error Resume Next
' Connect to the LDAP server's root object
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")
strTarget = "LDAP://" & strDNSDomain
wscript.Echo "Starting search from " & strTarget


'Crea fichero
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("grupos.txt", True)


' Connect to Ad Provider
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Proveedor de Active Directory"

Set objCmd =   CreateObject("ADODB.Command")
Set objCmd.ActiveConnection = objConnection

' Show only computers
'objCmd.CommandText = "SELECT Name, ADsPath FROM '" & strTarget & "' WHERE objectCategory = 'computer'"

' or show only Users
'objCmd.CommandText = "SELECT Name, ADsPath FROM '" & strTarget & "' WHERE objectCategory = 'user'"

' or show only Workgroups
objCmd.CommandText = "SELECT Name, ADsPath FROM '" & strTarget & "' WHERE objectCategory = 'group'"

Const ADS_SCOPE_SUBTREE = 2
objCmd.Properties("Page Size") = 100
objCmd.Properties("Timeout") = 30
objCmd.Properties("Searchscope") = ADS_SCOPE_SUBTREE
objCmd.Properties("Cache Results") = False

Set objRecordSet = objCmd.Execute

' Iterate through the results
objRecordSet.MoveFirst
objFile.WriteLine "  Nombre Grupos en AD:"
Do Until objRecordSet.EOF
   sComputerName = objRecordSet.Fields("Name")
   sADsPath = objRecordSet.Fields("ADsPath")

                objFile.WriteLine  sADsPath & "-" & sComputerName
   'wscript.Echo sComputerName
   objRecordSet.MoveNext
Loop