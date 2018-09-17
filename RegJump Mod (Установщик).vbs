' Установщик RegJump Mod в контекстное меню "Отправить"

Option Explicit
Dim oShell, oFSO, curPath, AppData, SendTo, InstFolder, AppLink, AppName, AppPath, AppArgs, WinDir, JobCommand, oShellApp, StartPath, StartLink, StartLinkUn, ContextDesktop, ContextFolder, LnkObjPath

Set oShell     = CreateObject("WScript.Shell")
Set oFSO       = CreateObject("Scripting.FileSystemObject")
Set oShellApp  = CreateObject("Shell.Application")

AppName   = "RegJump MOD by Dragokas ver 2.14"

SendTo      = oShell.SpecialFolders("SendTo")
AppData     = oShell.SpecialFolders("AppData")
WinDir      = oShell.ExpandEnvironmentStrings("%SystemRoot%")
AppLink     = SendTo & "\" & "Реестр - прыжок из буфера.lnk"
StartPath   = oShell.SpecialFolders("AllUsersPrograms") & "\RegJump MOD"
StartLink   = StartPath & "\" & "Реестр - прыжок из буфера.lnk"
StartLinkUn = StartPath & "\" & "Удалить RegJump Mod.lnk"

' Контекстное меню для рабочего стола и папки (по кнопке Shift)
ContextDesktop = "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Directory\Background\shell"
ContextFolder  = "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Directory\shell"

'Папка установки
InstFolder = AppData & "\RegJump Mod"
AppPath = InstFolder & "\" & "Jump Registry из буфера.vbs"

curPath = oFSO.GetParentFolderName(WScript.ScriptFullname)

if GetWindowsVersion() = "Vista" then call Elevate()

'Деинсталляция
if strcomp(curPath, StartPath, 1) = 0 then
	if msgbox("Вы уверены, что хотите удалить RegJump MOD ?", vbYesNo, AppName) = vbYes then 
		call Uninstall()
	else
		WScript.Quit
	end if
end if
if oFSO.FolderExists(InstFolder) then
	if msgbox("RegJump Mod уже установлена. Хотите удалить ее?", vbYesNo, AppName) = vbYes then
		call Uninstall()
	else
		WScript.Quit
	end if
end if

'Проверка, что запущен не из архива
if not oFSO.FolderExists(curPath & "\bin") then
	WScript.Echo "Сначала нужно распаковать все файлы из архива."
	WScript.Quit
end if

'Создаю папку для установки приложения
if not oFSO.FolderExists(InstFolder) then oFSO.CreateFolder InstFolder

'Копирую файлы приложения и аддона
oFSO.CopyFile curPath & "\bin\*", InstFolder & "\", true

'Снимаю поток безопасности
oShell.Run "cmd.exe /c ""<NUL set /p=>""" & InstFolder & "\Jump Registry из буфера.vbs" & """:Zone.Identifier:$DATA""", 0, false

' Предложение запускать без UAC
if GetWindowsVersion() = "Vista" then
	JobCommand = "/create /tn ""RegJump Mod SkipUAC"" /SC ONCE /ST 00:00 /F /RL HIGHEST " & _
                 "/tr ""wscript.exe \""" & AppData & "\RegJump Mod\Jump Registry из буфера.vbs\"" NoUAC"""
	oShellApp.ShellExecute "schtasks.exe", JobCommand, WinDir & "\System32", "runas", 1
	'AppArgs = "JobLauncher"
	LnkObjPath = "schtasks.exe"
	AppArgs = "/i /run /tn ""RegJump Mod SkipUAC"""
else
	LnkObjPath = AppPath
end if

'Создание папки в меню ПУСК
if not oFSO.FolderExists(StartPath) then oFSO.CreateFolder StartPath

'Создание ярлыка в папке SendTo (контекстное меню)
with oShell.CreateShortcut(StartLink)
	.Description        = AppName
	.IconLocation       = InstFolder & "\" & "Icon_RED.ico"
	.TargetPath         = LnkObjPath 'AppPath
	.Arguments          = AppArgs
	.WorkingDirectory   = InstFolder
	.WindowStyle        = 7 'minimized
	.Save
end with

'Создание ярлыка для удаления программы
oFSO.CopyFile WScript.ScriptFullname, InstFolder & "\Удалить RegJump MOD.vbs", true
oShell.Run "cmd.exe /c ""<NUL set /p=>""" & InstFolder & "\Удалить RegJump MOD.vbs" & """:Zone.Identifier:$DATA""", 0, false
oShell.Run "cmd.exe /c ""<NUL set /p=>""" & WScript.ScriptFullname & """:Zone.Identifier:$DATA""", 0, false

with oShell.CreateShortcut(StartLinkUn)
	.Description        = "Удаление RegJump Mod"
	.IconLocation       = InstFolder & "\" & "Icon_Green.ico"
	.TargetPath         = InstFolder & "\Удалить RegJump MOD.vbs"
	.Arguments          = ""
	.WorkingDirectory   = InstFolder
	.WindowStyle        = 7 'minimized
	.Save
end with

if Msgbox("Желаете создать пункт в контекстном меню ""Отправить"" ?", vbYesNo, AppName) = vbYes then
	'Копирование ярлыка в SendTo
	'oFSO.CopyFile StartLink, AppLink, true

  with oShell.CreateShortcut(AppLink)
	.Description        = AppName
	.IconLocation       = InstFolder & "\" & "Icon_RED.ico"
	.TargetPath         = AppPath
	.Arguments          = ""
	.WorkingDirectory   = InstFolder
	.WindowStyle        = 7 'minimized
	.Save
  end with

end if

if GetWindowsVersion() = "Vista" then
	Dim ContextStart: ContextStart = "wscript.exe" & " " & """" & AppPath & """" & " " & AppArgs

	'Создание контекстного меню по клавише Shift
	oShell.RegWrite ContextDesktop & "\RegJump MOD\",         "Реестр - прыжок из буфера",       "REG_SZ"
	oShell.RegWrite ContextFolder  & "\RegJump MOD\",         "Реестр - прыжок из буфера",       "REG_SZ"
	oShell.RegWrite ContextDesktop & "\RegJump MOD\Extended", "",                                "REG_SZ"
	oShell.RegWrite ContextFolder  & "\RegJump MOD\Extended", "",                                "REG_SZ"
	oShell.RegWrite ContextDesktop & "\RegJump MOD\Icon",     InstFolder & "\" & "Icon_RED.ico", "REG_SZ"
	oShell.RegWrite ContextFolder  & "\RegJump MOD\Icon",     InstFolder & "\" & "Icon_RED.ico", "REG_SZ"
	oShell.RegWrite ContextDesktop & "\RegJump MOD\Position", "Bottom",                          "REG_SZ"
	oShell.RegWrite ContextFolder  & "\RegJump MOD\Position", "Bottom",                          "REG_SZ"
	oShell.RegWrite ContextDesktop & "\RegJump MOD\command\", ContextStart,                      "REG_SZ"
	oShell.RegWrite ContextFolder  & "\RegJump MOD\command\", ContextStart,                      "REG_SZ"
end if

if Msgbox ("СОВЕТ: давайте назначим сочетание горячих клавиш для вызова этой программы ? ", vbYesNo, AppName) = vbYes then
	Dim oFolder, oFolderItem, objFIV
	Set oFolder = oShellApp.Namespace(oFSO.GetDriveName(StartLink))
	Set oFolderItem = oFolder.ParseName(mid(StartLink, 4))
    oFolderItem.InvokeVerb "Properties"
	WScript.Sleep 1000	
	msgbox "Нажмите желаемую комбинацию клавиш в поле ""Быстрый вызов""," & vbCrLf & "например Ctrl + Shift + Q",, AppName
	WScript.Sleep 60000
	'For Each objFIV In oFolderItem.Verbs
	'    If objFIV.Name = "Сво&йства" Then
	'        objFIV.DoIt
	'        Exit For
	'    End If
	'Next
	Set oFolder = Nothing: Set oFolderItem = Nothing ': Set objFIV = Nothing
end if

Set oShell = Nothing: Set oFSO = Nothing: Set oShellApp = Nothing


Function GetWindowsVersion() '"NT" или "Vista" core
	dim ver
	ver = CreateObject("WScript.Shell").RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentVersion")
	if left(ver, 2) = "5." then GetWindowsVersion = "NT" else GetWindowsVersion = "Vista"
End Function

Sub Uninstall()
	on error resume next
	' проверяю, устанавливался ли вместе с задачей
	if oShell.CreateShortcut(AppLink).TargetPath = WinDir & "\System32\cmd.exe" then
		oShellApp.ShellExecute "schtasks.exe", "/delete /TN ""RegJump Mod SkipUAC"" /F", WinDir & "\System32", "runas", 0
	end if
	oShell.RegDelete ContextDesktop & "\RegJump MOD\command\"
	oShell.RegDelete ContextFolder  & "\RegJump MOD\command\"
	oShell.RegDelete ContextDesktop & "\RegJump MOD\"
	oShell.RegDelete ContextFolder  & "\RegJump MOD\"
	oFSO.DeleteFile AppLink, true
	oFSO.DeleteFile SendTo & "\" & "Jump Registry из буфера.vbs", true
	Err.Clear
	oFSO.DeleteFolder StartPath, true
	if oFSO.FolderExists(InstFolder) then oFSO.DeleteFolder InstFolder, true
	if err.Number <> 0 then
		msgbox "Удалите самостоятельно папку: " & InstFolder
		oShell.Run "explorer.exe " & """" & InstFolder & """"
	else
		msgbox "RegJump MOD успешно удален."
	end if
	WScript.Quit
End Sub

Sub Elevate()
    Const DQ = """"
	if WScript.Arguments.Count = 0 then
	    oShellApp.ShellExecute WScript.FullName, DQ & WScript.ScriptFullName & DQ & " " & DQ & "Admin" & DQ, "", "runas", 1
		WScript.Quit
	end if
End Sub