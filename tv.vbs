' Troca a fonte de som no Skype
' Entre Microfone e a Mixagem est�reo


' Define variables
Dim objShell			' Command Line shell
Dim objSkypeAPI 		' Skype API
Dim outFone
Dim outTv

' Set variables
Set objShell =  CreateObject("WScript.Shell")
Set objSkypeAPI = Wscript.CreateObject("Skype4COM.Skype", "Skype_")
outFone = "Fones de ouvido / Alto falantes (Realtek High Definition Audio)"
outTv = "2D HD LG TV (�udio Intel(R) para telas)"


' If ( objSkypeAPI.Settings.AudioOut = outFone ) Then
' 	objSkypeAPI.Settings.AudioOut = outTv
' Else
' 	objSkypeAPI.Settings.AudioOut = outFone
' End If
objSkypeAPI.Settings.AudioOut = outTv

' Free memory
Set objShell = Nothing
Set objSkypeAPI = Nothing