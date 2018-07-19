Option Explicit: Dim SeparateRegedit

' $$$ RegJump MOD by Alex Dragokas            [  SafeZone.cc  ]
'
' $$$ Прыжок в раздел реестра, имя которого скопировано в буфер
' $$$ ver. 2.12.

' >>>>>>>>>   Настройки   <<<<<<<<

'SeparateRegedit = false ' не реализовано


' Поддерживаются такие форматы данных в буфере:

' 1) Сокращения ульев (HKLM, HKCU, HKCR, HKU)
' 2) [раздел реестра] (формат REG-файлов / логов RSIT, SITLog)
' 3) "раздел реестра"
' [не реализовано] -------- ' 4) Многострочные данные (для нескольких разделов открываются разные окна редактора реестра)
' 5) формат INI-файлов (в т.ч. формат AVZ html лога)
' 6) Игнорирование лишних слов перед именем раздела.
' 7) Логи HijackThis, MBAM, ComboFix

Dim oShell, oShellApp, InstFolder, rData, AppName, oFSO, oTS
Dim pos, SubRoot, RootKey
Dim sLines, sLine
Dim CompPrefix, lcode
Dim oRegEx, oMatches
Dim RunJob, AppPath
Dim sWinVer, sLastKey

Set oRegEx = CreateObject("VBScript.RegExp")
oRegEx.IgnoreCase = True

Set oShellApp  = CreateObject("Shell.Application")
set oShell     = CreateObject("WScript.Shell")
InstFolder = oShell.SpecialFolders("AppData") & "\RegJump Mod"
AppName = "RegJump MOD by Dragokas"

sWinVer = GetWindowsVersion()

' Проверка необходимости перезапуска от имени Задачи
if sWinVer = "Vista" then 
	if WScript.Arguments.Count = 0 then 
			oShell.Run "schtasks.exe /i /run /tn ""RegJump Mod SkipUAC""", 0, false
			WScript.Quit
	else
		if WScript.Arguments(0) <> "NoUAC" then
			oShell.Run "schtasks.exe /i /run /tn ""RegJump Mod SkipUAC""", 0, false
			WScript.Quit
		end if
	end if
end if

if sWinVer = "Vista" then call Elevate()

'AppPath = oFSO.GetParentFolderName(WScript.ScriptFullName)
AppPath = left(WScript.ScriptFullname, instrrev(WScript.ScriptFullname, "\") - 1)

rData = GetFromClipBoard()
if typename(rData) = "Null" then msgbox "Неверный тип данных в буфере!",,AppName: WScript.Quit

if Len(rData) = 0 then
	'попытка получить буфер через ConClip
	set oFSO       = CreateObject("Scripting.FileSystemObject")
	if not oFSO.FileExists(AppPath & "\GetClip.exe") then
		msgbox "Программа GetClip не найдена!" & vbcrlf & "Буфер обмена пуст!"
		WScript.Quit
	end if
	oShell.Run "cmd.exe /c """"" & AppPath & "\GetClip.exe"" /text > """ & AppPath & "\Clip.txt" & """""", 6, true
	if oFSO.FileExists(AppPath & "\Clip.txt") then
	    set oTS = oFSO.OpenTextFile(AppPath & "\Clip.txt", 1, false)
		rData = oTS.ReadAll
		oTS.Close
	end if
	if Len(rData) = 0 then
		msgbox "Буфер обмена пуст!"
		WScript.Quit
	end if
end if

' Мультистрочная ?
if instr(rData, vbLf) <> 0 then
	'Dim sLines, sLine
	sLines = Split(Replace(rData, vbCr, ""), vbLf) ' unix-style support
	For each sLine in sLines
		rData = ParseLine(sLine)				   ' парсинг строки. Выделить только улей\раздел\параметр
		if CheckValidKey(rData) then Exit For	   ' формат верный ?
	Next
else
	rData = ParseLine(rData)
end if

if not CheckValidKey(rData) then 
    ' попытка найти соответствие улья по части имени подраздела
    'Dim pos, SubRoot, RootKey

    pos = instr(rData, "\")
    if pos <> 0 then
    	SubRoot = left(rData, pos - 1)
    else
    	SubRoot = rData
    end if

	Select case UCase(SubRoot)
	case ".DEFAULT"
		RootKey = "HKEY_USERS"
	case "BCD00000000", "COMPONENTS", "HARDWARE", "SAM", "SCHEMA", "SECURITY", "SOFTWARE", "SYSTEM"
		RootKey = "HKEY_LOCAL_MACHINE"
	case "APPEVENTS", "CONSOLE", "CONTROL PANEL", "ENVIRONMENT", "EUDC", "IDENTITIES", "KEYBOARD LAYOUT", "NETWORK", "PRINTERS", "VOLATILE ENVIRONMENT"
		RootKey = "HKEY_CURRENT_USER"
	case else
		if left(UCase(SubRoot), 4) = "S-1-" then
			RootKey = "HKEY_USERS"
		end if
	end select
	if 0 <> len(RootKey) then rData = RootKey & "\" & rData

end if

if not CheckValidKey(rData) then
	'пытаемся найти имя корня вразнобой
	oRegEx.Pattern = ".*(HKEY_CLASSES_ROOT|HKEY_CURRENT_USER|HKEY_LOCAL_MACHINE|HKEY_USERS|HKEY_CURRENT_CONFIG|HKLM|HKCU|HKCR|HKU).*"
	Set oMatches = oRegEx.Execute(rData)
	if oMatches.Count <> 0 then
	    RootKey = oMatches.Item(0).SubMatches(0)
	end if

	oRegEx.Pattern = "(BCD00000000|COMPONENTS|HARDWARE|SAM|SCHEMA|SECURITY|SOFTWARE|SYSTEM)\\.*"
	Set oMatches = oRegEx.Execute(rData)
	if oMatches.Count <> 0 then
		if 0 = len(RootKey) then
			rData = "HKEY_LOCAL_MACHINE" & "\" & oMatches.Item(0).Value
		else
			rData = RootKey & "\" & oMatches.Item(0).Value
		end if
	else
		oRegEx.Pattern = "(APPEVENTS|CONSOLE|CONTROL PANEL|ENVIRONMENT|EUDC|IDENTITIES|KEYBOARD LAYOUT|NETWORK|PRINTERS|VOLATILE ENVIRONMENT)\\.*"			
		Set oMatches = oRegEx.Execute(rData)
		if oMatches.Count <> 0 then
			if 0 = len(RootKey) then
				rData = "HKEY_CURRENT_USER" & "\" & oMatches.Item(0).Value
			else
				rData = RootKey & "\" & oMatches.Item(0).Value
			end if
		else
			oRegEx.Pattern = "(\.DEFAULT\\|S-1-\d.*).*"			
			Set oMatches = oRegEx.Execute(rData)
			if oMatches.Count <> 0 then
				if 0 = len(RootKey) then
					rData = "HKEY_USERS" & "\" & oMatches.Item(0).Value
				else
					rData = RootKey & "\" & oMatches.Item(0).Value
				end if
			end if
		end if
	End if
end if

if not CheckValidKey(rData) then Msgbox "Не могу разобраться в таком буфере ^_^" & _
										vbCrLf & vbCrLf & rData,,AppName: WScript.Quit

'Преобразование сокращения улья в полное имя
rData = ExpandHiveName(rData)

call CloseRegedit()

'поиск префикса для Lastkey
On Error Resume Next
sLastKey = oShell.RegRead ("HKCU\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit\Lastkey")
On Error Goto 0

if len(sLastKey) <> 0 then
	pos = instr(sLastKey, "\")
	if pos <> 0 then
		CompPrefix = Left(sLastKey, pos)
	else
		CompPrefix = sLastKey & "\"
	end if
else
	lcode = GetInterfaceLangCode()
	if lcode <> 0 then
		if sWinVer = "NT" then
			if lcode = 1033 then CompPrefix = "My Computer\" else CompPrefix = "Мой компьютер\"
		else
			if lcode = 1033 then CompPrefix = "Computer\" else CompPrefix = "Компьютер\"
		end if
	else
		if sWinVer = "NT" then
			if GetOSInstallLangCode() = "0409" then CompPrefix = "My Computer\" else CompPrefix = "Мой компьютер\"
		else
			if GetOSInstallLangCode() = "0409" then CompPrefix = "Computer\" else CompPrefix = "Компьютер\"
		end if
	end if
end if

if right(rData, 1) = "\" then rData = Left(rData, len(rData) - 1)

oShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit\Lastkey",CompPrefix & rData,"REG_SZ"

oShell.Run "regedit.exe", 1, false

' зачистка
set oFSO       = CreateObject("Scripting.FileSystemObject")
if oFSO.FileExists(AppPath & "\Clip.txt") then oFSO.DeleteFile AppPath & "\Clip.txt", true
WScript.Quit

Function ExpandHiveName(byval sData)
	Dim arr, n, ret
	arr = Split(sData, "\")
	Select case Ucase(arr(0))
	case "HKCR"
		ret = "HKEY_CLASSES_ROOT"
	case "HKCU"
		ret = "HKEY_CURRENT_USER"
	case "HKLM"
		ret = "HKEY_LOCAL_MACHINE"
	case "HKU"
		ret = "HKEY_USERS"
	case "HKCC"
		ret = "HKEY_CURRENT_CONFIG"
	case else
		ret = arr(0)
	End Select
	For n = 1 to UBound(arr)
		ret = ret & "\" & arr(n)
	Next
	ExpandHiveName = ret
End Function

Function GetOSInstallLangCode()
	On Error Resume Next
	Dim lcode
	lcode = oShell.RegRead ("HKLM\SYSTEM\CurrentControlSet\Control\Nls\Language\InstallLanguage")
	if lcode = 0 then GetOSInstallLangCode = 0 else GetOSInstallLangCode = lcode
End Function

Function GetInterfaceLangCode()
	On Error Resume Next
	Dim lcode
	lcode = oShell.RegRead ("HKCU\Software\Microsoft\Windows\CurrentVersion\Controls Folder\Presentation LCID")
	if lcode = 0 then GetInterfaceLangCode = 0 else GetInterfaceLangCode = lcode
End Function

Function CloseRegedit()
	'Dim objWMIService, colProcess, objProcess
	'Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2") 
	'Set colProcess = objWMIService.ExecQuery ("Select * from Win32_Process Where (Name='Regedit.exe' or Name = 'Regedt32.exe')")
	'For Each objProcess in colProcess
	'	objProcess.Terminate: Exit For 'just 1 process of regedit.exe to close
	'Next
	'Set objWMIService = Nothing: set colProcess = Nothing: set objProcess = Nothing
	oShell.Run "taskkill.exe /f /im regedit.exe", 6, true
End Function

Function ParseLine(sLine)
	Dim newData, Part
	newData = sLine

	newData = Replace(newData, "/", "\")
	if Right(newData, 1) = "\" then newData = Left(newData, Len(newData) - 1)

	' если в средине текста
	oRegEx.Pattern = "(HKLM|HKCU|HKCR|HKU|HKEY_).*"
	Set oMatches = oRegEx.Execute(newData)
	if oMatches.Count <> 0 then
	    newData = oMatches.Item(0).Value
	else
		ParseLine = sLine
		Exit Function	' не найдено ничего подходящего
	End if

	' убираю обрамление пробелами
	newData = trim(newData)

	' поддержка прямого копирования текста из .reg-файлов (формат: [улей/ключ])
	if left (newData, 1) = "[" then newData = mid (newData, 2)
	if right(newData, 1) = "]" then newData = left(newData, len(newData) - 1)

	' заменяю двойные \\ на одиночные \
	newData = Replace(newData, "\\", "\")

	if Left(newData, 1) = "\" then newData = mid(newData, 2)

	'Замена HKUS -> HKU (формат HijackThis)
	if strcomp(left(newData,5),"HKUS\", vbTextCompare) = 0 then newData = "HKU\" & mid(newData,6)
	if strcomp(left(newData,5),"HKUS,", vbTextCompare) = 0 then newData = "HKU," & mid(newData,6)

	' если формат файла INF (или AVZ)
	oRegEx.Pattern = "^""?(HKLM|HKCU|HKCR|HKU|HKEY_CLASSES_ROOT|HKEY_CURRENT_USER|HKEY_LOCAL_MACHINE|HKEY_USERS|HKEY_CURRENT_CONFIG)""?,.*"
	if oRegEx.test(newData) then 							' Формат INF-файла ?
		'Dim Part
		Part = Split(newData,",")
		if Ubound(Part) = 0 then
			newData = trim(UnQuote(trim(Part(0))))
		elseif Ubound(Part) = 1 then
			newData = trim(UnQuote(trim(Part(0)))) & "\" & trim(UnQuote(trim(Part(1))))
		else
			newData = trim(UnQuote(trim(Part(0)))) & "\" & trim(UnQuote(trim(Part(1)))) & "\" & trim(UnQuote(trim(Part(2))))
		end if
	else
		oRegEx.Pattern = "^'?(HKLM|HKCU|HKCR|HKU|HKEY_CLASSES_ROOT|HKEY_CURRENT_USER|HKEY_LOCAL_MACHINE|HKEY_USERS|HKEY_CURRENT_CONFIG)'?,.*"		
		if oRegEx.test(newData) then 							' Формат INF-файла ?
			Part = Split(newData,",")
			if Ubound(Part) = 0 then
				newData = trim(UnQuote(trim(Part(0))))
			elseif Ubound(Part) = 1 then
				newData = trim(UnQuote(trim(Part(0)))) & "\" & trim(UnQuote(trim(Part(1))))
			else
				newData = trim(UnQuote(trim(Part(0)))) & "\" & trim(UnQuote(trim(Part(1)))) & "\" & trim(UnQuote(trim(Part(2))))
			end if
		end if
	end if
	
	' убираю кавычки
	newData = UnQuote(newData)

	ParseLine = newData
	
End Function

Function CheckValidKey(rData)
	Dim oRegEx
	Set oRegEx = CreateObject("VBScript.RegExp")
	oRegEx.IgnoreCase = True
	oRegEx.Pattern = "^(HKLM|HKCU|HKCR|HKU|HKEY_).*"
	CheckValidKey = false
	if oRegEx.test(rData) then CheckValidKey = true
	set oRegEx = Nothing
End Function

Function GetFromClipBoard()
	On Error Resume Next
	GetFromClipBoard = CreateObject("htmlfile").ParentWindow.ClipboardData.GetData("text")
End Function

Function GetWindowsVersion() '"NT" или "Vista" core
	On Error Resume Next
	dim ver
	ver = CreateObject("WScript.Shell").RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentVersion")
	if left(ver, 2) = "5." then GetWindowsVersion = "NT" else GetWindowsVersion = "Vista"
End Function

Function UnQuote(Str) ' убирает обрамляющие кавычки в т.ч. идущие подряд
    Dim s: s = Str
    Do While Left(s, 1) = """"
        s = Mid(s, 2)
    Loop
    Do While Right(s, 1) = """"
        s = Left(s, Len(s) - 1)
    Loop
    Do While Left(s, 1) = "'"
        s = Mid(s, 2)
    Loop
    Do While Right(s, 1) = "'"
        s = Left(s, Len(s) - 1)
    Loop
    UnQuote = s
End Function

Sub Elevate()
    Const DQ = """"
	if WScript.Arguments.Count = 0 then
	    oShellApp.ShellExecute WScript.FullName, DQ & WScript.ScriptFullName & DQ & " " & DQ & "Admin" & DQ, "", "runas", 1
		WScript.Quit
	end if
End Sub