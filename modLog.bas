Attribute VB_Name = "modLog"
Option Explicit

Public Const logFile As String = "hoosns.log"
Public Const logRepliesFile As String = "activity.log"

Public Sub Log_Create(ByVal fn As String, ByRef bol As Boolean)

    If bol = False Then
        Exit Sub
    End If
    
    On Error Resume Next
    Open App.Path & "\" & fn For Output As #2
    If Err.Number <> 0 Then
        bol = False
        Exit Sub
    End If
    
    Close #2
    
End Sub

Public Sub Log_Update(ByVal fn As String, ByRef bolDoLog As Boolean, ByVal str As String, Optional bolUpdateGUI As Boolean)

    If bolDoLog = False Then
        Exit Sub
    End If
    
    On Error Resume Next
    Open App.Path & "\" & fn For Append As #2
    If Err.Number <> 0 Then
        bolDoLog = False
        Exit Sub
    End If
    
    Print #2, Now & " " & str
    
    Close #2
    
    If bolUpdateGUI = True Then
        Form1.lblStatusText.Caption = str
    End If
    
End Sub
