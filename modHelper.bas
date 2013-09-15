Attribute VB_Name = "modHelper"
Option Explicit
'Functions provided by differnect sources on the internet. Thanks.

Public Function HexToDec(HexValue As String, Optional ToInteger As Boolean) As Long
    If ToInteger Then
        ' convert to an integer value, if possible.
        ' Use CInt() if you want to *always* convert to an Integer
        HexToDec = Val("&H" & HexValue)
    Else
        ' always convert to a Long. You can also use the CLng() function.
        HexToDec = Val("&H" & HexValue & "&")
    End If
End Function

Public Function FileExists(sFullPath As String) As Boolean

    Dim oFile As New Scripting.FileSystemObject
    FileExists = oFile.FileExists(sFullPath)
    
End Function

Public Function TrimComplete(ByVal sValue As String) As _
        String

    Dim sAns As String
    Dim sWkg As String
    Dim sChar As String
    Dim lLen As Long
    Dim lCtr As Long

    sAns = sValue
    lLen = Len(sValue)

    If lLen > 0 Then
        'Ltrim
        For lCtr = 1 To lLen
            sChar = Mid(sAns, lCtr, 1)
            If Asc(sChar) > 32 Then Exit For
        Next

        sAns = Mid(sAns, lCtr)
        lLen = Len(sAns)

        'Rtrim
        If lLen > 0 Then
            For lCtr = lLen To 1 Step -1
                sChar = Mid(sAns, lCtr, 1)
                If Asc(sChar) > 32 Then Exit For
            Next
        End If
        sAns = Left$(sAns, lCtr)
    End If

    TrimComplete = sAns

End Function

Public Function CharCount(OrigString As String, _
  Chars As String, Optional CaseSensitive As Boolean = False) _
  As Long

'**********************************************
'PURPOSE: Returns Number of occurrences of a character or
'or a character sequencence within a string

'PARAMETERS:
    'OrigString: String to Search in
    'Chars: Character(s) to search for
    'CaseSensitive (Optional): Do a case sensitive search
    'Defaults to false

'RETURNS:
    'Number of Occurrences of Chars in OrigString

'EXAMPLES:
'Debug.Print CharCount("FreeVBCode.com", "E") -- returns 3
'Debug.Print CharCount("FreeVBCode.com", "E", True) -- returns 0
'Debug.Print CharCount("FreeVBCode.com", "co") -- returns 2
''**********************************************

Dim lLen As Long
Dim lCharLen As Long
Dim lAns As Long
Dim sInput As String
Dim sChar As String
Dim lCtr As Long
Dim lEndOfLoop As Long
Dim bytCompareType As Byte

sInput = OrigString
If sInput = "" Then Exit Function
lLen = Len(sInput)
lCharLen = Len(Chars)
lEndOfLoop = (lLen - lCharLen) + 1
bytCompareType = IIf(CaseSensitive, vbBinaryCompare, _
   vbTextCompare)

    For lCtr = 1 To lEndOfLoop
        sChar = Mid(sInput, lCtr, lCharLen)
        If StrComp(sChar, Chars, bytCompareType) = 0 Then _
            lAns = lAns + 1
    Next

CharCount = lAns

End Function

