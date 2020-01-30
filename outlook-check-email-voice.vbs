Set Sapi = Wscript.CreateObject("SAPI.SpVoice")
dim str

Sapi.speak "Checking for new messages. Please standby."
WScript.Sleep 2000

Set otl = createobject("outlook.application")
Set session = otl.getnamespace("mapi")

session.logon "Outlook", , True, True 
	Set inbox = session.getdefaultfolder(6)
	c = 0

	For Each m In inbox.items
		If m.unread Then c = c + 1
	Next
session.logoff

s = "s"
If c = 1 Then s = ""

Sapi.speak "You have" & c & "unread message" & s
