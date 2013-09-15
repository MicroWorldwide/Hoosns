Attribute VB_Name = "modINI"
Option Explicit

Sub Parse_INI()

    Dim str As String
    Dim i As Long
    i = 0
    
    On Error Resume Next
    Open App.Path & "\" & iniFile For Input As #1
    If Err.Number <> 0 Then
        Exit Sub
    End If
    
    Do Until EOF(1)
        i = i + 1
        Line Input #1, str
        str = LCase(TrimComplete(str))
        If str <> "" And Not Left(1, ";") Then
            If InStr(str, ";") Then 'take off the comment
                str = TrimComplete(Left(str, InStr(str, ";") - 1))
            End If
            If Left(str, 4) = "port" Then
                str = TrimComplete(Mid(str, InStr(str, "=") + 1, Len(str)))
                If IsNumeric(str) Then 'this line and next if can't be together (?)
                    If str < 65535 And str > 0 Then
                        PortNumber = Int(str)
                    Else
                        Call Log_Update(logFile, bolWriteLog, "Invalid port number on line " & i & " in '" & iniFile & "', using default: " & PortNumber, True)
                    End If
                Else
                    Call Log_Update(logFile, bolWriteLog, "Non numeric data on line " & i & " in '" & iniFile & "'")
                End If
            ElseIf Left(str, 7) = "guihide" Then
                str = TrimComplete(Mid(str, InStr(str, "=") + 1, Len(str)))
                If str = "1" Then
                    bolGUIHide = True
                ElseIf str = "0" Then
                    bolGUIHide = False
                Else
                    Call Log_Update(logFile, bolWriteLog, "Non boolean on line 'guihide' " & i & " in '" & iniFile & "'")
                End If
            ElseIf Left(str, 11) = "listeningip" Then
                Dim objRegExp As RegExp
                str = TrimComplete(Mid(str, InStr(str, "=") + 1, Len(str)))
                Set objRegExp = New RegExp
                objRegExp.IgnoreCase = True
                objRegExp.Global = True
                'IP4 address.
                objRegExp.Pattern = "^(25[0-5]|2[0-4]\d|[0-1]?\d?\d)(\.(25[0-5]|2[0-4]\d|[0-1]?\d?\d)){3}$"
                If (objRegExp.Test(str) = True) Then
                    listeningIP = str
                Else
                    If str <> "all" Then
                        Call Log_Update(logFile, bolWriteLog, "Non IP string for 'listeningip' " & i & " in '" & iniFile & "'")
                    End If
                End If
                
            ElseIf Left(str, 10) = "timetolive" Then
                str = TrimComplete(Mid(str, InStr(str, "=") + 1, Len(str)))
                If IsNumeric(str) Then
                    If str < 20000000 And str > 0 Then
                        ttl1 = Int(str / 16777216)
                        ttl2 = Int(str / 65536)
                        ttl3 = Int(str / 256)
                        ttl4 = Int(str Mod 256)
                        ttl = str
                    Else
                        Call Log_Update(logFile, bolWriteLog, "Invalid time to live (ttl) on line " & i & " in '" & iniFile & "'")
                    End If
                Else
                    Call Log_Update(logFile, bolWriteLog, "Non numeric data on line " & i & " in '" & iniFile & "'")
                End If
            ElseIf Left(str, 17) = "logactivitytofile" Then
                str = TrimComplete(Mid(str, InStr(str, "=") + 1, Len(str)))
                If str = "1" Then
                     bolWriteRequestLog = True
                ElseIf str = "0" Then
                    bolWriteRequestLog = False
                Else
                    Call Log_Update(logFile, bolWriteLog, "Non boolean on line 'logrequests' " & i & " in '" & iniFile & "'")
                End If
            ElseIf Left(str, 11) = "startpaused" Then
                str = TrimComplete(Mid(str, InStr(str, "=") + 1, Len(str)))
                If str = "1" Then
                    bolListening = True
                ElseIf str = "0" Then
                    bolListening = False
                Else
                    Call Log_Update(logFile, bolWriteLog, "Non boolean on line 'createptr' " & i & " in '" & iniFile & "'")
                End If
            ElseIf Left(str, 9) = "createptr" Then
                str = TrimComplete(Mid(str, InStr(str, "=") + 1, Len(str)))
                If str = "1" Then
                     bolCreatePTR = True
                ElseIf str = "0" Then
                    bolCreatePTR = False
                Else
                    Call Log_Update(logFile, bolWriteLog, "Non boolean on line 'createptr' " & i & " in '" & iniFile & "'")
                End If
            End If
        End If
    Loop
    
    Close #1
    
End Sub
