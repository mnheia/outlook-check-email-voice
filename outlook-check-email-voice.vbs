' outlook-check-email-voice checks the number of unread messages in Outlook.
'
' Copyright (c) 2018, Mnheia <mnheia@gmail.com>
'
' This module is free software; you can redistribute it and/or modify it
' under the terms of GNU general public license (gpl) version 3.
' See the LICENSE file for details.

' Configuration
Language = "en"       ' en, de, ru, es, fr, zh

Set Sapi = WScript.CreateObject("SAPI.SpVoice")
SetVoice Sapi, Language

Select Case LCase(Language)
    Case "de"
        checkingText = U("Ich pr\u00FCfe auf neue Nachrichten. Bitte warten.")
        unreadText = "Ungelesene Nachrichten"
    Case "ru"
        checkingText = U("\u041F\u0440\u043E\u0432\u0435\u0440\u044F\u044E \u043D\u043E\u0432\u044B\u0435 \u0441\u043E\u043E\u0431\u0449\u0435\u043D\u0438\u044F. \u041F\u043E\u0436\u0430\u043B\u0443\u0439\u0441\u0442\u0430, \u043F\u043E\u0434\u043E\u0436\u0434\u0438\u0442\u0435.")
        unreadText = U("\u041D\u0435\u043F\u0440\u043E\u0447\u0438\u0442\u0430\u043D\u043D\u044B\u0435 \u0441\u043E\u043E\u0431\u0449\u0435\u043D\u0438\u044F")
    Case "es"
        checkingText = "Comprobando mensajes nuevos. Por favor espere."
        unreadText = U("Mensajes no le\u00EDdos")
    Case "fr"
        checkingText = "Recherche de nouveaux messages. Veuillez patienter."
        unreadText = "Messages non lus"
    Case "zh"
        checkingText = U("\u6B63\u5728\u68C0\u67E5\u65B0\u90AE\u4EF6\uFF0C\u8BF7\u7A0D\u5019\u3002")
        unreadText = U("\u672A\u8BFB\u90AE\u4EF6")
    Case Else
        checkingText = "Checking for new messages. Please standby."
        unreadText = "Unread messages"
End Select

Sapi.Speak checkingText
WScript.Sleep 2000

Set otl = CreateObject("Outlook.Application")
Set session = otl.GetNamespace("MAPI")

session.Logon
Set inbox = session.GetDefaultFolder(6)

c = 0
For Each m In inbox.Items
    If m.Unread Then c = c + 1
Next

session.Logoff

Sapi.Speak unreadText
Sapi.Speak CStr(c)

Sub SetVoice(Sapi, ByRef Language)
    Select Case LCase(Language)
        Case "de"
            languageId = "407"
        Case "ru"
            languageId = "419"
        Case "es"
            languageId = "40A"
        Case "fr"
            languageId = "40C"
        Case "zh"
            languageId = "804"
        Case Else
            Language = "en"
            languageId = "409"
    End Select

    Set voices = Sapi.GetVoices("Language=" & languageId)

    If voices.Count > 0 Then
        Set Sapi.Voice = voices.Item(0)
    Else
        Language = "en"
        Set voices = Sapi.GetVoices("Language=409")
        If voices.Count > 0 Then Set Sapi.Voice = voices.Item(0)
    End If
End Sub

Function U(text)
    result = ""

    Do While Len(text) > 0
        pos = InStr(text, "\u")

        If pos = 0 Then
            result = result & text
            Exit Do
        End If

        result = result & Left(text, pos - 1)
        result = result & ChrW(CLng("&H" & Mid(text, pos + 2, 4)))
        text = Mid(text, pos + 6)
    Loop

    U = result
End Function
