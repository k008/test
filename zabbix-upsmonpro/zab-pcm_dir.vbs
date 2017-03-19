AllDir = "D:\Data-Storage\vbs\zabbix-upsmonpro"                     ' ---------  Полное имя рабочего каталога (без слэжа \ на конце)

OutStr = AllDir + vbCrLf
OutStr = OutStr + AllFolders(AllDir)

MsgBox OutStr

' ---------------------------------------------------------------------------
Function AllFolders(WDir)
'   MsgBox WDir
    Rezult = ""
    Set F = CreateObject("Scripting.FileSystemObject").GetFolder(WDir)
    Set SubF = F.SubFolders

    For Each Folder In SubF
		Rezult = Rezult + AllFiles(WDir + "\" + Folder.Name) + vbCrLf
        'Rezult = Rezult + AllFolders(WDir + "\" + Folder.Name)
    Next

    AllFolders = Rezult
End Function

Function AllFiles(W1dir)
	Set FL = CreateObject("Scripting.FileSystemObject").GetFolder(W1Dir).Files
	For Each FF in FL
		AllFiles = FF
		'msgbox FF + "4444"

	Next

End Function