Attribute VB_Name = "modHostsFile"
Option Explicit

Sub Load_Hosts(ByVal bolStart As Boolean)

    Dim str As String
    Dim i As Long
    Dim a As String
    
    numHosts = 0
    
    If Not FileExists(App.Path & "\" & hostFN) Then
        If bolStart Then
            strHostsData = "Missing file: '" & hostFN & "'"
            Form1.txtHostsList.Text = strHostsData
        Else
            'MsgBox ("Missing file: " & hostFN)
        End If
        Call Log_Update(logFile, bolWriteLog, "Error: Missing file: '" & hostFN & "'", True)
        Exit Sub
    End If
    
    On Error Resume Next
    Open App.Path & "\" & hostFN For Input As #1
    If Err.Number <> 0 Then
        Call Log_Update(logFile, bolWriteLog, "Error: Unable to access '" & hostFN & "' file", True)
        strHostsData = "Unable to access '" & hostFN & "' file"
        Form1.txtHostsList.Text = strHostsData
        Exit Sub
    End If

    For i = 1 To colHosts.Count
        colHosts.Remove (1)
    Next
    
    For i = 1 To colPTRHosts.Count
        colPTRHosts.Remove (1)
    Next
    
    i = 0
    Do Until EOF(1)
        i = i + 1
        Line Input #1, str
        str = TrimComplete(str)
        If Check_Hosts_Line(str, i) = 0 Then
            Create_Hosts_Collection (str)
        End If
    Loop
    
    Close #1
    
        
    For i = 1 To colHosts.Count
        'Create gui host string
        If i Mod modHosts = 0 Then
            a = vbTab '... before host name
            If Len(colHosts.Item(i - 1)) < 9 Then
                a = a & vbTab
            End If
            strHostsData = strHostsData & a & colHosts.Item(i) & "" & vbCrLf
        ElseIf i Mod modHosts = 2 Then
            a = vbTab
            strHostsData = strHostsData & a & colHosts.Item(i)
        Else
            strHostsData = strHostsData & colHosts.Item(i)
        End If
    Next
    
    If bolCreatePTR Then
        strHostsData = strHostsData & vbCrLf & vbTab & "-------PTR-------" & vbCrLf
    End If
    
    For i = 1 To colPTRHosts.Count
        'Create gui host string
        If i Mod modHosts = 0 Then
            a = ""
            If Len(colPTRHosts.Item(i - 1)) < 9 Then
                a = vbTab
            End If
            strHostsData = strHostsData & a & colPTRHosts.Item(i) & "" & vbCrLf
        Else
            strHostsData = strHostsData & colPTRHosts.Item(i) & vbTab
        End If
    Next
    
    If numHosts < 1 Then
        Call Log_Update(logFile, bolWriteLog, "No valid host entries found in '" & hostFN & "' file")
        If intTextShow = 1 Then
            Form1.txtHostsList.Text = "No valid host entries found in '" & hostFN & "' file"
        ElseIf intTextShow = 2 Then
            If requestCount = 0 Then
                Form1.txtHostsList.Text = "   -- no requests received --"
            Else
                Form1.txtHostsList.Text = strRequestsData
            End If
        ElseIf intTextShow = 3 Then
            Form1.txtHostsList.Text = strHelp
        End If
    Else
        Call Log_Update(logFile, bolWriteLog, numHosts & " valid IP/Name maps found in '" & hostFN & "'")
        If intTextShow = 1 Then
            Form1.txtHostsList.Text = strHostsData
        ElseIf intTextShow = 2 Then
            If requestCount = 0 Then
                Form1.txtHostsList.Text = "   -- no requests received --"
            Else
                Form1.txtHostsList.Text = strRequestsData
            End If
        ElseIf intTextShow = 3 Then
            Form1.txtHostsList.Text = strHelp
        End If
    End If
    
    Form1.lblNumHosts.Caption = numHosts

End Sub


Sub Create_Hosts_Collection(str As String)

    Dim objRegExp As RegExp
    Dim objMatch As Match
    Dim colMatches As MatchCollection
    Dim spchr As Long, spchr2 As Long, i As Long
    Dim IP As String, strHostName As String
    Dim arySect() As String, a As String, z As String, B As String
    Dim bolFoundPTR As Boolean
    Dim cntChr As Long
    Dim bolDoubleColon As Boolean
    
    i = 0
    strHostsData = ""
    Set objRegExp = New RegExp
    bolFoundPTR = False
    bolDoubleColon = False
    
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    objRegExp.Pattern = "([\s]+[\w.\-]+)"
 
    If (objRegExp.Test(str) = False) Then        'It really should be!
        Exit Sub
    End If
    
    'Add 3 collection objects for each entry
    '1 = IP address
    '2 = host name
    '3 = request counter
    Set colMatches = objRegExp.Execute(str)
    For Each objMatch In colMatches
        i = i + 1
        spchr = objMatch.FirstIndex
        If i = 1 Then
            IP = TrimComplete(Left(str, spchr))
        End If
        strHostName = TrimComplete(objMatch)
        If Len(strHostName) > 512 Then      'RFCs restrict this (UDP)
            Call Log_Update(logFile, bolWriteLog, "Invalid Name (too long) in " & hostFN & ".")
            Exit Sub
        ElseIf strHostName = "." Then 'this valid? don't think so...
            Call Log_Update(logFile, bolWriteLog, "Invalid Name (.) in " & hostFN & ".")
            Exit Sub
        End If
        colHosts.Add 0             'IP
        colHosts.Add IP            'strHostName
        colHosts.Add strHostName   'request counter
        numHosts = numHosts + 1             'increment all host counter
    Next
    Set colMatches = Nothing

    If InStr(str, ".in-addr.arpa") Or InStr(str, ".ip6.int") Or InStr(str, ".ip6.arpa") Then
        bolFoundPTR = True
    End If
        
    '1 PTR per IP if required (if not a ptr record already)
    If bolFoundPTR = True Then
        colPTRHosts.Add 0                       'host name
        colPTRHosts.Add IP           'IP
        colPTRHosts.Add strHostName                      'request counter
        colHosts.Remove (colHosts.Count)
        colHosts.Remove (colHosts.Count)
        colHosts.Remove (colHosts.Count)
    ElseIf bolCreatePTR = True Then
        a = ""
        If InStr(IP, ".") Then
            arySect = Split(IP, ".")
            For i = UBound(arySect) To 0 Step -1
                a = a & arySect(i) & "."
            Next
            B = a & "in-addr.arpa"
        ElseIf InStr(IP, ":") Then
            Dim i4 As Long, j As Long
            i4 = 0
            cntChr = CharCount(IP, ":")
            
            For i = 1 To Len(IP)
                z = Mid(IP, i, 1)
                i4 = i4 + 1
                If z = ":" And i4 = 1 And bolDoubleColon = False Then
                    On Error Resume Next
                    If Mid(IP, i - 1, 1) = ":" Then
                        If Err.Number = 0 Then
                            For j = (cntChr * 4) To (8 * 4) - 1
                                a = "0." & a
                            Next
                            B = a & B
                            a = ""
                            bolDoubleColon = True
                        Else
                            For j = i4 To 4
                                a = a & "0."
                                i4 = i4 + 1
                            Next
                        End If
                    End If
                    i4 = 0
                ElseIf z = ":" Then
                    For j = i4 To 4
                        a = a & "0."
                        i4 = i4 + 1
                    Next
                Else
                    a = z & "." & a
                End If
                If i4 >= 4 Then
                    B = a & B
                    a = ""
                    i4 = 0
                End If
            Next
            If Len(a) < 8 And a <> "" Then 'tidy up last nibbles
                For j = (Len(a) / 2) To (8 / 2) - 1
                    a = a & "0."
                Next
                B = a & B
                a = ""
            End If
            For j = Len(B) / 2 To 31
                B = "0." & B
            Next
            If Len(B) <> 64 Then
                B = "err." & B & "ip6.arpa"
            Else
                B = B & "ip6.arpa"
            End If
            
        End If
        colPTRHosts.Add 0            'host name
        colPTRHosts.Add strHostName                      'IP
        colPTRHosts.Add B                      'request counter
        numHosts = numHosts + 1             'increment all host counter
    End If
 
    
End Sub

Function Check_Hosts_Line(str As String, lineNo As Long)
    'str should have been trimmed before coming here.
    Dim objRegExp As RegExp
    Dim spchr As Long
    Dim a As String
    Dim objMatch As Match
    Dim colMatches As MatchCollection
    Dim i As Long, j As Long

    i = 0
    a = ""
    Set objRegExp = New RegExp

    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    
    If str = "" Then
        Check_Hosts_Line = 1
        Exit Function
    End If
    
    objRegExp.Pattern = "^#"
    If (objRegExp.Test(str) = True) Then
        Check_Hosts_Line = 1
        Exit Function
    End If
    
    If InStr(str, "#") Then 'take off the comment
        str = TrimComplete(Left(str, InStr(str, "#") - 1))
    End If

    'IP4 this only allows 3 domain names after IP.
    objRegExp.Pattern = "^(25[0-5]|2[0-4]\d|[0-1]?\d?\d)(\.(25[0-5]|2[0-4]\d|[0-1]?\d?\d)){3}([\s]+[a-z0-9\-.]+)([\s]*[a-z0-9\-.]*)([\s]*[a-z0-9\-.]*)$"
    If (objRegExp.Test(str) = True) Then
        Check_Hosts_Line = 0
        Exit Function
    End If
    
    'IPv4 PTR
    objRegExp.Pattern = "^[a-z0-9\-.]+[\s]+(25[0-5]|2[0-4]\d|[0-1]?\d?\d)(\.(25[0-5]|2[0-4]\d|[0-1]?\d?\d)){3}(\.in-addr.arpa)$"
    If (objRegExp.Test(str) = True) Then
        Check_Hosts_Line = 0
        Exit Function
    End If
    
    'IPv6 PTR (.arpa)
    objRegExp.Pattern = "^[a-z0-9\-.]+[\s]+([0-9a-f]\.){32}(ip6.arpa)$"
    If (objRegExp.Test(str) = True) Then
        Check_Hosts_Line = 0
        Exit Function
    End If
    
    'IPv6 PTR (.int)
    objRegExp.Pattern = "^[a-z0-9\-.]+[\s]+([0-9a-f]\.){32}(ip6.int)$"
    If (objRegExp.Test(str) = True) Then
        Check_Hosts_Line = 0
        Exit Function
    End If

    i = 0
    'IP6 checks now. I'm sure this could be improved.
    'First, there should only be one space.
    objRegExp.Pattern = "[\s]+"
    Set colMatches = objRegExp.Execute(str)
        For Each objMatch In colMatches
            i = i + 1
            If i = 1 Then
                spchr = objMatch.FirstIndex
            End If
        Next
    Set colMatches = Nothing
    'here is the actual check.
    If i < 1 Or i > 3 Then
GoTo invalid_ip:
    End If

    a = LCase(Left(str, spchr))
    'Check that there are some : chars
    i = CharCount(a, ":")
    If i < 2 Or i > 7 Then
GoTo invalid_ip:
    End If

    objRegExp.Pattern = "([0-9a-f]*)::([0-9a-f]+)::([0-9a-f]*)"
    If (objRegExp.Test(str) = True) Then
GoTo invalid_ip:
    End If

    'make sure there are only hex digits or ':' left.
    i = 1
    j = 0
    objRegExp.Pattern = "[0-9a-f]+"
    Do While i <= Len(a)
        If Mid(a, i, 1) = ":" Then
            j = 0
        ElseIf objRegExp.Test(Mid(a, i, 1)) = False Then
GoTo invalid_ip:
        End If
        i = i + 1
        j = j + 1
        If j > 5 Then
GoTo invalid_ip:
        End If
    Loop
    Check_Hosts_Line = 0
    Exit Function
    
invalid_ip:
    Call Log_Update(logFile, bolWriteLog, "Invalid IP/Name in " & hostFN & " on line: " & lineNo)
    Check_Hosts_Line = 1
    
End Function
