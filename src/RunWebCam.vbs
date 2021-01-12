'**************************************************************************************************
'	RunWebCam.vbs
'
'	Description: Script run MPC-HC in fullscreen live mode
'	Parameters:  None
'
'	Author: Dmitrii Glukhov
'**************************************************************************************************

Dim WshShell
Dim FSObject
Dim Program
Dim Timer
Dim Counter
Dim Result

Program = "C:\Program Files\K-Lite Codec Pack\Media Player Classic\mpc-hc.exe"
Timer = 100
Counter = 0
	
Set FSObject = CreateObject("Scripting.FileSystemObject")
If FSObject.FileExists(Program) Then
	Set WshShell = WScript.CreateObject("WScript.Shell")
	WshShell.Run """" & Program & """"
	Do
		WScript.Sleep Timer
		Result = WshShell.AppActivate("Media Player Classic Home Cinema")
		Counter = Counter + 1	
	Loop Until (Result Or Counter = 100)
	If Result Then
		WScript.Sleep Timer
        	WshShell.SendKeys "%{ENTER}"
		WScript.Sleep Timer
		WshShell.SendKeys "+W"
	Else
		MsgBox "Ошибка переключения на главное окно программы", 16, "Ошибка"
	End If
	Set WshShell = Nothing
Else
	MsgBox	"Media Player Classic Home Cinema не установлен." & vbNewLine &_
		"Получить программу можно по адресу http://mpc-hc.org/", 16, "Ошибка"
End If
Set FSObject = Nothing

'**************************************************************************************************
