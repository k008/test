Fdir="D:\Data-Storage\vbs\zabbix-upsmonpro\"
report="D:\Data-Storage\vbs\zabbix-upsmonpro\report.txt"

' проверил - лог, записал (время, дата, путь)'
' размер файла'
' дата файла'
' выбирать папки от 1 до 12'
Sub WriteReport(sDataReport)
  Dim FSOL, EReport, FileReport
  Set FSOL = CreateObject("Scripting.FileSystemObject")
  EReport = 0
  If FSOL.FileExists(report) Then
    Set FileReport=FSOL.OpenTextFile(report, 8)
	EReport=1
  End If

  If Not FSOL.FileExists(report) Then
    SET FileReport=FSOL.CreateTextFile(report, True)
	EReport=0
  End If

	FileReport.WriteLine(sData)
    FileReport.Close
End Sub

Sub Scan()
End Sub

Function CheckSize(csFile)
  Dim FSOL1, getcsFile, csFileSize
  Set FSOL1 = CreateObject("Scripting.FileSystemObject")
  If FSOL1.FileExists(fDir & csFile) Then
    Set getcsFile = FSOL1.GetFile(fDir & csFile)
    csFileSize = getcsFile.Size
    MsgBox "Размер файла " & csFile & "=" & csFileSize
  End If
End Function

Function fSectoGMT(fsecond)
	Dim objWMI, objCollection, objItem
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
End Function

Function ConverNowToUNIXTime()
	ConverNowToUNIXTime=DateDiff("s", "01/01/1970 00:00:00", Now())
End Function