'vcfdfd
















dire_1= "C:\Backups & Logs\Bkp-F5\"
soft_1= " F5 "

dire_2="C:\Backups & Logs\Bkp-Checkpoint\"
soft_2=" CheckPoint "



a= SendEmail(listar2(dire_1),soft_1)
a= SendEmail(listar2(dire_2),soft_2)


'========================================================================


Function SendEmail(msg,prd)

Set objMessage = CreateObject("CDO.Message") 
objMessage.Subject = "control de backup" & " de " & prd &"   -- " + now()
objMessage.From = "seguridadIT@bunge.com" 
bjMessage.To = "security.bsc.it@bunge.com" 
'objMessage.To = "gerardo.dionofrio.ext@bunge.com"


'strHTML = "<HTML>"
'strHTML = strHTML & "<HEAD><head><meta http-equiv=Content-Type content=text/html;charset=utf-8></head>"
'strHTML = strHTML & "<BODY>"
'strHTML = strHTML & "<b> Le informamos que la contraseña de su cuenta " & ID & " expira en " & dias & " día, </br> la cual podrá cambiar Ud. mismo desde una computadora de Bunge antes de dicha fecha o solicitar </br> a nuestro Help Desk que lo haga por Ud. enviando  un  correo a SupportCenter.BSC.IT@bunge.com" _
'                    & " </br>  Sepa que de no realizar el cambio de contraseña, no podrá ingresar a la red de Bunge luego de la fecha mencionada. </br> </br> </br> NO RESPONDA ESTE MAIL, PUES SE HA GENERADO AUTOMÁTICAMENTE. </b></br>"
'strHTML = strHTML & "</BODY>"
'strHTML = strHTML & "</HTML>"


objMessage.HTMLBody = msg

 
objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "10.1.6.24"
objMessage.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25  

'Set FileSystem = WScript.CreateObject("Scripting.FileSystemObject")
'archivo =dire & "\" & file1

'Set InputArchivo = FileSystem.OpenTextFile(archivo,1)

'objMessage.TextBody=InputArchivo.ReadAll 
'iputArchivo.close

objMessage.Configuration.Fields.Update
objMessage.Send

SendEmail="a"
end function


'========================================================================

FUNCTION LISTAR2(PATH)
'Wscript.Echo path

Set objFSO = CreateObject("Scripting.FileSystemObject")


Set objFolder = objFSO.GetFolder(path)

Set colFiles = objFolder.Files
For Each objFile in colFiles
    if Weekday(objFile.DateCreated)= 2 then
	     ' Wscript.Echo 
		 'a= a & objFile.DateCreated &  "   "  & objFile.Name   & vbCrLf
		 a= a & objFile.DateCreated &  "   "  & objFile.Name   & "<br>"
	end if
Next
'Wscript.Echo  a
LISTAR2 = a

end function

' Wscript.Echo 