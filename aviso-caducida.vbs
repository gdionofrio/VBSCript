'strOU = "C=bar,DC=sa,DC=dir,DC=bunge,DC=com"
strOU = "bue-212-bar2/DC=bar,DC=sa,DC=dir,DC=bunge,DC=com"

Set fso = CreateObject ("Scripting.FileSystemObject")
Set stdout = fso.GetStandardStream (1)

set fso = CreateObject("Scripting.FileSystemObject")
curDir = fso.GetAbsolutePathName(".")

Set FileSystem = WScript.CreateObject("Scripting.FileSystemObject")

Set fso = CreateObject ("Scripting.FileSystemObject")
Set stdout = fso.GetStandardStream (1)


strDate = CStr(Int(Now()))
strDate = Replace(strDate, "/", "-")
UserCSV = "ver-usuarios-" & strDate & ".csv"

CuentasVencidas = "Cuentas-Vencidas-" & strDate & ".csv"

Set OutPutFile = FileSystem.CreateTextFile(UserCSV, True) 
Set OutPutCuentasV = FileSystem.CreateTextFile(CuentasVencidas, True) 

OutPutFile.writeline "cn;id;dias_sin_logueo;codigo"
OutPutCuentasV.writeline "cn;id;dias_sin_logueo;codigo;fecha_creación;Cant_dias_creado"

'On Error Resume Next


Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open "Provider=ADsDSOObject;"


Set objCommand = CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 1000

objCommand.CommandText = _
  "<LDAP://" & strOU & ">;" & _
  "(&(objectclass=user)(objectcategory=person));" & _
  "cn,description,adspath,accountExpires,pwdlastset,distinguishedname,sAMAccountName;lastlogontimestamp;subtree"
Set objRecordSet = objCommand.Execute

objRecordSet.MoveFirst


	Do Until objRecordSet.EOF
        Set objUser = GetObject(objRecordSet.Fields("AdsPath").Value)
		
	    
		msgbox cdbl(objRecordSet.fields("pwdlastset"))
		OutPutFile.writeline objUser.cn & ";" & objUser.sAMAccountName & ";" & objRecordSet.fields("accountExpires") 
		'& ";" & objUser.pwdlastset
		objRecordSet.MoveNext 
    loop

	OutPutFile.close
	OutPutCuentasV.close
	

	a=SendEmail(paraDeshabilitar,"Usuarios para ser deshabilitados por inactividad de mas de 90 días del dominio BAR",strDate,curDir)
	a=SendEmail(paraBorrar,"Usuarios para ser borrados por inactividad de mas de 180 días dominio BAR",strDate,curDir)
	a=SendEmail(paraDeshabilitar_no_deshabilitada,"Usuarios para ser borrados por inactividad de más de 180 días del dominio BAR (activos)",strDate,curDir)
	a=SendEmail(UsuariosSinActividad,"Usuarios sin actividad dominio BAR",strDate,curDir)




'---------------------------------------------------------------------------------------
'Function: GetLastLogonStamp
'Last Modified: 10/13/05 .csm
'This function uses the Active Directory 2003 schema attribute 'lastlogontimestamp' to
'pull the information on the last logon date.  It returns the number of days that the
'user last logged into the system.  Since this attribute is only replicated every 14
'days, users are warned to expect a 14 day variance in the output.
'---------------------------------------------------------------------------------------

Function GetLastLogonStamp(strUserAccountName)
Dim objRecordSetLast
Set objRootDSE = GetObject("LDAP://RootDSE")
strConfig = objRootDSE.Get("configurationNamingContext")
strDNSDomain = objRootDSE.Get("defaultNamingContext")
Set objcmd = CreateObject("ADODB.Command")
Set objConn = CreateObject("ADODB.Connection")
objConn.Provider = "ADsDSOObject"
objConn.Open "Active Directory Provider"
objcmd.ActiveConnection = objConn

objcmd.Properties("Page Size") = 100
objcmd.Properties("Timeout") = 60
objcmd.Properties("Cache Results") = False

objCmd.CommandText = strQuery

strBase = "<LDAP://" & strDNSDomain & ">"
strFilter = "(&(objectCategory=person)(objectClass=user) (sAMAccountName=" & strUserAccountName & _
  		"))"
strAttributes = "lastlogontimestamp"
strQuery = strBase & ";" & strFilter & ";" & strAttributes _
    & ";subtree"
 
objCmd.CommandText = strQuery
On Error Resume Next
Set objRecordSetLast = objCmd.Execute
If Err.Number <> 0 Then
    On Error GoTo 0
    GetLastLogonStamp = "Error"
Else
   On Error Resume Next
    Do Until objRecordSetLast.EOF
      lngDate = objRecordSetLast.Fields("lastLogontimestamp")
      objRecordSetLast.MoveNext
    Loop
End If
If IsNull(lngdate)  Then
Else
   If CStr(Integer8Date(lngdate)) = "1/1/1601" Then
	GetLastLogonStamp = "Unknown"
   Else
	GetLastLogonStamp = Int(Now - Integer8Date(lngdate))
   End If
End If		
End Function

'---------------------------------------------------------------------------------------
'Function: Integer8Date
'Last Modified: 9/28/05 .csm
'This function accepts a date (in Integer8 format) and outputs it to date format
'adjusted to consider local time zone bias
'Source = rmeuller.net
'---------------------------------------------------------------------------------------
Function Integer8Date(objDate)
' Obtain local time zone bias from machine registry.
Set objShell = CreateObject("Wscript.Shell")
lngBiasKey = objShell.RegRead("HKLM\System\CurrentControlSet\Control\" _
  & "TimeZoneInformation\ActiveTimeBias")
If UCase(TypeName(lngBiasKey)) = "LONG" Then
  lngBias = lngBiasKey
ElseIf UCase(TypeName(lngBiasKey)) = "VARIANT()" Then
  lngBias = 0
  For k = 0 To UBound(lngBiasKey)
    lngBias = lngBias + (lngBiasKey(k) * 256^k)
  Next
End If
'Do the conversion
Dim lngAdjust, lngDate, lngHigh, lngLow
lngAdjust = lngBias
lngHigh = objDate.HighPart
lngLow = objdate.LowPart
' Account for bug in IADslargeInteger property methods.
If lngLow < 0 Then
	lngHigh = lngHigh + 1
End If
If (lngHigh = 0) And (lngLow = 0) Then
	lngAdjust = 0
End If
lngDate = #1/1/1601# + (((lngHigh * (2 ^ 32)) _
	+ lngLow) / 600000000 - lngAdjust) / 1440
Integer8Date = CDate(lngDate)
End Function

'---------------------------------------------------------------------------------------
'Function: SendEmail
'---------------------------------------------------------------------------------------
Function SendEmail(file1,msg,fecha,dir)


Set objMessage = CreateObject("CDO.Message") 
objMessage.Subject = msg & "   " & fecha
objMessage.From = "seguridadIT@bunge.com" 
'objMessage.To = "gerardo.dionofrio.ext@bunge.com" 
objMessage.To = "security.bsc.it@bunge.com" 

 
objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "10.1.6.24"
objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25  


Set FileSystem = WScript.CreateObject("Scripting.FileSystemObject")
archivo =dir & "\" & file1
Set InputArchivo = FileSystem.OpenTextFile(archivo,1)
objMessage.TextBody=InputArchivo.ReadAll 
InputArchivo.close

objMessage.Configuration.Fields.Update
objMessage.Send

SendEmail="a"
end function

'---------------------------------------------------------------------------------------
' Listado de los posibles valores que puede tomar el campo de >UserAccountControl<
' Los valores se pueden sumar.
'
'SCRIPT			0x0001						1
'ACCOUNTDISABLE	0x0002						2
'HOMEDIR_REQUIRED	0x0008					8	
'LOCKOUT	0x0010							16
'PASSWD_NOTREQD	0x0020						32
'PASSWD_CANT_CHANGE 0x0040					64
'ENCRYPTED_TEXT_PWD_ALLOWED	0x0080			128
'TEMP_DUPLICATE_ACCOUNT	0x0100				256
'NORMAL_ACCOUNT	0x0200						512
'INTERDOMAIN_TRUST_ACCOUNT	0x0800			2048
'WORKSTATION_TRUST_ACCOUNT	0x1000			4096
'SERVER_TRUST_ACCOUNT	0x2000				8192
'DONT_EXPIRE_PASSWORD	0x10000				65536
'MNS_LOGON_ACCOUNT	0x20000					131072
'SMARTCARD_REQUIRED	0x40000					262144
'TRUSTED_FOR_DELEGATION	0x80000				524288
'NOT_DELEGATED	0x100000					1048576
'USE_DES_KEY_ONLY	0x200000				2097152
'DONT_REQ_PREAUTH	0x400000				4194304
'PASSWORD_EXPIRED	0x800000				8388608
'TRUSTED_TO_AUTH_FOR_DELEGATION	0x1000000	16777216
'PARTIAL_SECRETS_ACCOUNT	0x04000000	 	67108864