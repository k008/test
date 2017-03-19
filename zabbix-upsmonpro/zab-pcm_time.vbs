set fso=createobject("scripting.filesystemobject") 
set f=fso.GetFile("C:\Users\Asus\Desktop\Программы\VNC\Dialog-Aqua-Frunze\DIALOG-AQUA-user-pk-vpn.vnc") 
'msgbox day(f.DateLastModified)& month(f.DateLastModified) & year(f.DateLastModified)
'msgbox DateDiff("s", "01/01/1970 00:00:00", (fConverTimetoSec(Now)))
'msgbox fSectoGMT(DateDiff("s", "01/01/1970 00:00:00", Now()))

msgbox Now() & vbCRLF & f.DateLastModified

'call fSectoGMT(6)

function fSectoGMT(fsecond)
	Dim objWMI, objCollection, objItem, Daylight
	Daylight=3
	Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
	Set objCollection = objWMI.ExecQuery("SELECT * FROM Win32_ComputerSystem")
	For Each objItem In objCollection
		If objItem.DaylightInEffect Then
			WScript.Echo "Смещение относительно Гринвича (час.): " & objItem.CurrentTimeZone*60 & _
				vbNewLine & "Режим ""Автоматический переход на летнее время и обратно"" включен."
		Else
			WScript.Echo "Смещение относительно Гринвича (час.): " & objItem.CurrentTimeZone*60 & _
				vbNewLine & "Режим ""Автоматический переход на летнее время и обратно"" отключен."
		End If
		fSectoGMT=fsecond-objItem.CurrentTimeZone*60
	Next
	Set objCollection = Nothing
	Set objWMI = Nothing
end function

function fConverTimetoSec(fTime)
	'msgbox Hour(fTime)
	'msgbox Minute(dTime)/60
	'msgbox Second(fTime)
	fTime=(Hour(fTime)*3600)+(Minute(dTime)*60) + Second(fTime)
	fConverTimetoSec=fTime
	msgbox ftime	
end function

Set objArgs = WScript.Arguments
'Вывод всех аргументов (как пример работы с ними)
For I = 0 to objArgs.Count - 1
   WScript.Echo objArgs(I)
Next 