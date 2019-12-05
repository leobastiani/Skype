' Troca a fonte de som no Skype
' Entre Microfone e a Mixagem estéreo


' Define variables
Dim objShell			' Command Line shell
Dim objSkypeAPI 		' Skype API
Dim inMic
Dim inMix

' Set variables
Set objShell =  CreateObject("WScript.Shell")
Set objSkypeAPI = Wscript.CreateObject("Skype4COM.Skype", "Skype_")
inMic = "Microfone (Realtek High Definition Audio)"
inMix = "Mixagem estéreo (Realtek High Definition Audio)"


If ( objSkypeAPI.Settings.AudioIn = inMic ) Then
	objSkypeAPI.Settings.AudioIn = inMix
Else
	objSkypeAPI.Settings.AudioIn = inMic
End If

' Free memory
Set objShell = Nothing
Set objSkypeAPI = Nothing