VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hoosns"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   Icon            =   "hoosns.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.OptionButton radHelp 
      Caption         =   "Option1"
      Height          =   255
      Left            =   3600
      TabIndex        =   35
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txtHostsList 
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   5415
      Left            =   240
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "hoosns.frx":014A
      Top             =   1680
      Width           =   7215
   End
   Begin VB.CommandButton btnEdit 
      Caption         =   "EDIT"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton btnReloadNames 
      Caption         =   "RELOAD"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton CLose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      ToolTipText     =   "Closes window and stops listening for DNS requests"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton btnStopStart 
      Caption         =   "Pause"
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Name/IP Mapping File"
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   7320
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Status"
      Height          =   735
      Left            =   3120
      TabIndex        =   17
      Top             =   7320
      Width           =   855
      Begin VB.Label lblStatus 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.OptionButton radNameIP 
      BackColor       =   &H8000000B&
      Caption         =   "Option1"
      Height          =   255
      Left            =   360
      TabIndex        =   32
      Top             =   1320
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.OptionButton radRequests 
      BackColor       =   &H8000000A&
      Caption         =   "Option2"
      Height          =   255
      Left            =   2280
      TabIndex        =   33
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      Height          =   255
      Left            =   3840
      TabIndex        =   36
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label lblTabRequests 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Activity"
      Height          =   255
      Left            =   2400
      TabIndex        =   31
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblStatusText 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   5280
      TabIndex        =   29
      Top             =   7440
      Width           =   2055
   End
   Begin VB.Label lblIP 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "-all-"
      Height          =   255
      Left            =   840
      TabIndex        =   28
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "IP:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   375
   End
   Begin VB.Label lblTTL 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2640
      TabIndex        =   27
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "TTL:"
      Height          =   255
      Left            =   2160
      TabIndex        =   26
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "'+' answer ok       '.' unknown host      '?' unknown request type"
      Height          =   255
      Left            =   480
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Running"
      Height          =   255
      Left            =   4440
      TabIndex        =   24
      Top             =   7440
      Width           =   615
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Paused"
      Height          =   255
      Left            =   4440
      TabIndex        =   23
      Top             =   7680
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4200
      TabIndex        =   21
      Top             =   7920
      Width           =   135
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF0000&
      Height          =   135
      Left            =   4200
      TabIndex        =   20
      Top             =   7680
      Width           =   135
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FF00&
      Height          =   135
      Left            =   4200
      TabIndex        =   19
      Top             =   7440
      Width           =   135
   End
   Begin VB.Label lblNumSuccessfullRequest 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4440
      TabIndex        =   15
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblSuccessfulRequests 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Success:"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblNumRequest 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblNumHosts 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hosts:"
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblNoRequests 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Requests:"
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblPortNumber 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblServerName 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label lblPort 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Port:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblHostID 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Server:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Error"
      Height          =   255
      Left            =   4440
      TabIndex        =   22
      Top             =   7920
      Width           =   495
   End
   Begin VB.Label lblTabName 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "    Name/IP Mapping "
      Height          =   255
      Left            =   360
      TabIndex        =   30
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   240
      TabIndex        =   34
      Top             =   1200
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents socket_udp As CSocketMaster
Attribute socket_udp.VB_VarHelpID = -1


Private Sub lblHelp_Click()
    Call radHelp_Click
End Sub

Private Sub lblTabName_Click()
    Call radNameIP_Click
End Sub

Private Sub lblTabRequests_Click()
    Call radRequests_Click
End Sub

Private Sub radHelp_Click()
    Form1.txtHostsList.Text = strHelp
    Form1.radHelp.Value = True
    intTextShow = 3
End Sub

Private Sub radNameIP_Click()
    Form1.txtHostsList.Text = strHostsData
    Form1.radNameIP.Value = True
    intTextShow = 1
End Sub

Private Sub radRequests_Click()
    If requestCount = 0 Then
        Form1.txtHostsList.Text = "   -- no requests received --"
    Else
        Form1.txtHostsList.Text = strRequestsData
    End If
    Form1.radRequests.Value = True
    intTextShow = 2
End Sub

Private Sub socket_udp_DataArrival(ByVal bytesTotal As Long)

    'Process of answering query:
    ' 1. retrieve raw data.
    ' 2. extract requested name
    ' 3. replace delimiter values with '.'
    ' 4. check type and class, hoosns only recognises limited values
    ' 5. check name to IP mapping
    ' 6. check ptr to name mapping, if required
    ' 7. create reply
    
    'strings
    Dim strRAWREQUEST As String         'Raw data from socket
    Dim strRemoteHostIP As String       'Requesting hosts IP
    Dim strDelimDomainName As String    'Domain kept in original format
    'Dim aryDomainNames(0) As String         'Domain name being requested
    Dim a As String, c As String
    
    'bytes
    Dim aryREPLY() As Byte          'Reply array of bytes
    Dim qtype As Byte
    Dim qclass As Byte
    Dim qflag_1 As Byte
    Dim qflag_2 As Byte
    Dim qdatalen As Byte
    Dim aflag_2 As Byte
    Dim qnoquestions As Byte
    
    'booleans
    Dim bolFoundIP As Boolean
    Dim bolInvalidType As Boolean
    
    'longs
    Dim i As Long
    Dim dotDelim As Long
    Dim j As Long
    Dim k As Long
    Dim lenNamePart As Long
    Dim answers As Long
    Dim colonPos As Long
    Dim lastColonPos As Long
    Dim cntChr As Long
    
    'arrays
    Dim aryDots() As Variant
    Dim aryOct As Variant
    Dim aryDomainNames() As Variant

    'initialise some variables
    bolFoundIP = False
    answers = 0
    IP = ""
    i = 0
    aflag_2 = 3     'hoosns doesn't know by default
    a = ""
    colonPos = 1
    lastColonPos = 0
    k = 0
    
    '1. Retrieve the raw DNS request
    On Error Resume Next
    socket_udp.GetData strRAWREQUEST, vbString
    If Err.Number <> 0 Then
        Call Log_Update(logFile, bolWriteLog, "Error: Winsock GetData ", True)
        Exit Sub
    End If
    
    'Is this done type of thing done by winsock?
    strRAWREQUEST = Left(Trim(strRAWREQUEST), 512)
    If Len(strRAWREQUEST) < 19 Then
        Call Log_Update(logFile, bolWriteLog, "Error: packet data too small")
        Exit Sub
    End If
  
    'Give remote IP to socket
    strRemoteHostIP = socket_udp.RemoteHostIP
    socket_udp.RemoteHost = strRemoteHostIP
    
    'Start creating the binary data for the reply.
    'The reply is dynamically lengthened depending on content.
    ReDim Preserve aryREPLY(12) '12* bytes - see below.
    
'===============================================================================
'                       Breakdown of DNS reply packet.
'*******************************************************************************
'1        '    15'16      '     32          1      2         3     4
'¦---------------.---------------¦ <-   ¦---------------.---------------¦
'¦ ID            ¦ Flags         ¦  :   ¦ ae 3f         ¦ SQResp 0x8100 ¦ (129)(0)
'¦---------------.---------------¦  :   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'¦ No of Qstns   ¦ No of RRs     ¦ 12*  ¦    <dissection of flags>
'¦---------------.---------------¦bytes ¦ 1....... message is a response
'¦ No Authrtv RR ¦ No Addtnl RR  ¦  :   ¦ .0000... opcode (0 stdqry, 1 invers, 2 svr status req
'¦-------------------------------¦ <-   ¦ .....0.. auth, not dom auth
'¦           questions           ¦      ¦ ......0. truncated (?not available at mo in hoosns)
'¦-------------------------------¦      ¦ .......1 recursion desired (?not available at mo in hoosns)
'¦            answers            ¦      ¦ ........ 0....... recursion available, (?not available at mo in hoosns)
'¦-------------------------------¦      ¦ ........ .0...... Z reserved
'¦           authority           ¦      ¦ ........ ..0..... Answer authenticated by auth
'¦-------------------------------¦      ¦ ........ ...0.... always 0 (ignored by ethereal)
'¦     additional information    ¦      ¦ ........ ....0000 return code (0 no error, 3 error)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'===============================================================================

    qflag_1 = Asc(Mid(strRAWREQUEST, 3, 1))
    'only use qflaq_1 at mo: qflag_2 = Asc(Mid(strRAWREQUEST, 4, 1))
    qnoquestions = Asc(Mid(strRAWREQUEST, 6, 1))
    
    If (qnoquestions < 1 Or qnoquestions > 12) _
        Or qflag_1 >= 128 Then
GoTo continue_answer:
    End If
    
    '<--2. EXTRACT domain name
    j = 1
    Do While j <= 1 'qnoquestions
        i = j * 13 'position of first character of name
        Do While i < Len(strRAWREQUEST) And Asc(Mid(strRAWREQUEST, i, 1)) <> 0
            a = Mid(strRAWREQUEST, i, 1)
            strDelimDomainName = strDelimDomainName & a
            i = i + 1
        Loop
        'Query type: 1 = A, 12 = PTR
        qtype = Asc(Mid(strRAWREQUEST, i + 2, 1))
        'Query class: hoosns only understands IN (1)
        qclass = Asc(Mid(strRAWREQUEST, i + 4, 1))
        'strDelimDomainName = Right(strDelimDomainName, Len(strDelimDomainName) - 1)
        ReDim Preserve aryDomainNames(j - 1) '12* bytes - see below.
        aryDomainNames(j - 1) = strDelimDomainName
        j = j + 1
    Loop
    '-->
    
    '<--3. REPLACE '.'s in domain name
    i = 0
    ReDim Preserve aryDots(0)                   '3www7openbsd3org0
    dotDelim = Asc(Left(aryDomainNames(0), 1))   'length of first part of domain name
    aryDomainNames(0) = Right(aryDomainNames(0), Len(aryDomainNames(0)) - 1)
    aryDots(0) = dotDelim                       'add what is found to an array, then
    lenNamePart = dotDelim
    Do While dotDelim   'we stop at the end of the name, the last char is 0
        dotDelim = 0
        i = i + 1
        ReDim Preserve aryDots(i)

        On Error Resume Next    'Error returned on check
        dotDelim = Asc(Mid(aryDomainNames(0), dotDelim + lenNamePart + i, 1))
        If Err.Number <> 0 Then
            ReDim Preserve aryDots(i - 1)
            dotDelim = 0
        Else
            lenNamePart = lenNamePart + dotDelim
            aryDots(i) = dotDelim
        End If
    Loop
    For i = 0 To UBound(aryDots) 'this should be change to NOT be a general Replace()
        aryDomainNames(0) = Replace(aryDomainNames(0), Chr(aryDots(i)), ".")
    Next
    '-->
 
    ' 4. If certain flags/types are not correct we continue
    If (qtype <> 1 And qtype <> 12 And qtype <> 28) _
        Or qclass <> 1 Then
        bolInvalidType = True
GoTo continue_answer:
    Else
        strHostsData = ""
    End If
    
    If qtype = 28 Then  'IPv6
        qdatalen = 16
    Else                'IPv4
        qdatalen = 4
    End If              'PTR dealt with later...
    
    strHostsData = ""
    '<-- 5. CHECK name to IP mapping
    For i = 1 To colHosts.Count
        c = ""
        a = LCase(colHosts.Item(i)) 'debug
        If LCase(aryDomainNames(0)) = a And i Mod modHosts = 0 And bolFoundIP = False Then
            IP = colHosts.Item(i - 1)
            If (qtype = 28 And InStr(1, IP, ":")) Or _
                (qtype = 1 And InStr(1, IP, ".")) Then
                    j = colHosts.Item(i - 2) + 1
                    'implement round robin?
                    colHosts.Remove i - 2
                    colHosts.Add j, , i - 2
                    answers = answers + 1
                    bolFoundIP = True
                    aflag_2 = 0
                    c = "<"
            Else
                IP = ""
            End If
        End If
        If i Mod modHosts = 0 Then
            a = vbTab
            If Len(colHosts.Item(i - 1)) < 9 Then
                a = a & vbTab
            End If
            strHostsData = strHostsData & colHosts.Item(i - 2) & c & vbTab & _
                           colHosts.Item(i - 1) & a & colHosts.Item(i) & vbCrLf
        End If
    Next
    '-->END
    
    'Divide name/IP and PTR mappings in GUI
    If bolCreatePTR Then
        strHostsData = strHostsData & vbCrLf & vbTab & "-------PTR-------" & vbCrLf
    End If

    '<-- 6. CHECK PTR to Name mappings
    For i = 1 To colPTRHosts.Count
        a = colPTRHosts.Item(i)
        c = ""
        If LCase(aryDomainNames(0)) = a And i Mod modHosts = 0 And bolFoundIP = False Then
            j = colPTRHosts.Item(i - 2) + 1
            IP = colPTRHosts.Item(i - 1)
            'implement round robin?
            colPTRHosts.Remove i - 2
            colPTRHosts.Add j, , i - 2
            answers = answers + 1
            bolFoundIP = True
            aflag_2 = 0
            qdatalen = Len(IP) + 2
            c = "<"
        End If
        If i Mod modHosts = 0 Then
            a = vbTab
            If Len(colPTRHosts.Item(i - 1)) < 9 Then
                a = a & vbTab
            End If
            strHostsData = strHostsData & colPTRHosts.Item(i - 2) & c & vbTab & _
                           colPTRHosts.Item(i - 1) & a & colPTRHosts.Item(i) & vbCrLf
        End If
    Next
    '-->END
   
continue_answer:

    '7. Start creating the reply (aryREPLY)
    'The first 2 byte field is the DNS transaction ID
    aryREPLY(0) = Asc(Left(strRAWREQUEST, 1))   'id 1st byte
    aryREPLY(1) = Asc(Mid(strRAWREQUEST, 2, 1)) 'id 2nd byte
    aryREPLY(2) = 129       'Flags 1
    aryREPLY(3) = aflag_2   'Flags 2 - either 0 or 3 (no such name)
    aryREPLY(4) = 0         'Questions 1
    aryREPLY(5) = 1         'Questions 2
    aryREPLY(6) = 0         'Answers RRS 1
    aryREPLY(7) = answers   'number of RRS's
    aryREPLY(8) = 0         'Authority RRS 1
    aryREPLY(9) = 0         'Authority RRS 2
    aryREPLY(10) = 0        'Additional RRS 1
    aryREPLY(11) = 0        'Additional RRS 2
    'aryREPLY(12) = Asc(Mid(strRAWREQUEST, 13, 1))
    
    'Insert requested domain name in reply
    i = 1
    Do While i < Len(strDelimDomainName) + 1
        ReDim Preserve aryREPLY(i + 11)
        aryREPLY(i + 11) = Asc(Mid(strDelimDomainName, i, 1))
        i = i + 1
    Loop
    i = i - 1
    ReDim Preserve aryREPLY(i + 16)
    aryREPLY(i + 12) = 0        'end of domain name
    aryREPLY(i + 13) = 0        'Type A (host name) 1
    aryREPLY(i + 14) = qtype    'Type A (host name) 2
    aryREPLY(i + 15) = 0        'Class IN 1
    aryREPLY(i + 16) = 1        'Class IN 2

        
    If IP <> "" Then
        successfulRequestCount = successfulRequestCount + 1
        Form1.lblNumSuccessfullRequest.Caption = successfulRequestCount
        
        ReDim Preserve aryREPLY(i + 28)
        aryREPLY(i + 17) = 192      'Name 1
        aryREPLY(i + 18) = 12       'Name 2
        aryREPLY(i + 19) = 0        'Type A 1
        aryREPLY(i + 20) = qtype    'Type A 2
        aryREPLY(i + 21) = 0        'Class IN 1
        aryREPLY(i + 22) = 1        'Class IN 2
        aryREPLY(i + 23) = ttl1     'TTL 1
        aryREPLY(i + 24) = ttl2     'TTL 2
        aryREPLY(i + 25) = ttl3     'TTL 3
        aryREPLY(i + 26) = ttl4     'TTL 4
        aryREPLY(i + 27) = 0        'Length 1
        aryREPLY(i + 28) = qdatalen   'Length 2
        
        j = i + 29
        
        If qtype = 28 Then 'aaaa
            ReDim Preserve aryREPLY(i + 44)
            
            cntChr = CharCount(IP, ":")
            For i = 1 To 8
                colonPos = InStr(lastColonPos + 1, IP, ":")
                If colonPos = 0 Then 'We have hit the end of ip string
                    colonPos = Len(IP) + 1
                End If
                If (colonPos - lastColonPos) > 3 Then
                    If (colonPos - lastColonPos) > 4 Then
                        aryREPLY(j + (2 * i) - 2) = HexToDec(Mid(IP, colonPos - 4, 2), True)
                    Else
                        aryREPLY(j + (2 * i) - 2) = HexToDec(Mid(IP, colonPos - 3, 1), True)
                    End If
                Else
                    aryREPLY(j + (2 * i) - 2) = 0
                End If
                If (colonPos - lastColonPos) > 1 Then
                    If (colonPos - lastColonPos) > 2 Then
                        aryREPLY(j + (2 * i) - 1) = HexToDec(Mid(IP, colonPos - 2, 2), True)
                    Else
                        aryREPLY(j + (2 * i) - 1) = HexToDec(Mid(IP, colonPos - 1, 1), True)
                    End If
                Else
                    aryREPLY(j + (2 * i) - 1) = 0
                End If
                If Mid(IP, colonPos + 1, 1) = ":" Then
                    For k = i To 8 - cntChr
                        i = i + 1
                        aryREPLY(j + (2 * i) - 2) = 0 'ok?
                        aryREPLY(j + (2 * i) - 1) = 0
                    Next
                    lastColonPos = colonPos + 1
                Else
                    lastColonPos = colonPos
                End If
            Next
        ElseIf qtype = 12 Then 'ptr
            ReDim Preserve aryREPLY(i + 30 + Len(IP))
            cntChr = CharCount(IP, ".")
            If cntChr = 0 Then
                i = 0
                aryREPLY(i + j) = Len(IP)
                For i = 1 To Len(IP)
                    aryREPLY(i + j) = Asc(Mid(IP, i, 1))
                Next
            Else
                i = 0
                aryREPLY(i + j) = InStr(1, IP, ".") - 1
                For i = 1 To Len(IP)
                    If Mid(IP, i, 1) = "." Then
                        If InStr(i + 1, IP, ".") = 0 Then
                            aryREPLY(i + j) = Len(IP) - InStr(i, IP, ".")
                        Else
                            aryREPLY(i + j) = InStr(i + 1, IP, ".") - InStr(i, IP, ".") - 1
                        End If
                    Else
                        a = Asc(Mid(IP, i, 1))
                        aryREPLY(i + j) = Asc(Mid(IP, i, 1))
                    End If
                Next
            End If
            ' aryREPLY(29 + Len(IP)) = 2
        Else
            ReDim Preserve aryREPLY(i + 32)
            aryOct = Split(IP, ".")
            For i = 0 To UBound(aryOct)
                aryREPLY(i + j) = aryOct(i)
            Next
        End If
        a = "+"
    ElseIf bolInvalidType = True Then
        a = "? "
    Else
        a = ". "
    End If

    Select Case qtype
    Case 1:
        a = a & " A "
    Case 2:
        a = a & " NS "
    Case 3:
        a = a & " MD "
    Case 4:
        a = a & " MF "
    Case 5:
        a = a & " CNAME "
    Case 6:
        a = a & " SOA "
    Case 7:
        a = a & " MB "
    Case 8:
        a = a & " MG "
    Case 9:
        a = a & " MR "
    Case 10:
        a = a & " NULL "
    Case 11:
        a = a & " WKS "
    Case 12:
        a = a & " PTR "
    Case 13:
        a = a & " HINFO "
    Case 14:
        a = a & " MINFO "
    Case 15:
        a = a & " MX "
    Case 16:
        a = a & " TXT "
    Case 17:
        a = a & " RP "
    Case 18:
        a = a & " AFSDB "
    Case 19:
        a = a & " X25 "
    Case 20:
        a = a & " ISDN "
    Case 21:
        a = a & " RT "
    Case 22:
        a = a & " NSAP "
    Case 23:
        a = a & " NSAP-PPTR "
    Case 24:
        a = a & " SIG "
    Case 25:
        a = a & " KEY "
    Case 27:
        a = a & " GPOS "
    Case 28:
        a = a & " AAAA "
    Case 29:
        a = a & " LOC "
    Case 30:
        a = a & " NEXT "
    Case 33:
        a = a & " SRV "
    Case 35:
        a = a & " NAPTR "
    Case 36:
        a = a & " KX "
    Case 38:
        a = a & " A6 "
    Case 39:
        a = a & " DNAME "
    Case 43:
        a = a & " DS "
    Case 249:
        a = a & " TKEY "
    Case 250:
        a = a & " TSIG "
    Case Else:
        a = a & " -?- "
    End Select
    
    colRequestsData.Add a & " " & Now & " " & aryDomainNames(0) & " -- " & strRemoteHostIP
    'Form1.lblActivity.BackColor = &HE0E0E0
    
    On Error Resume Next
    socket_udp.SendData (aryREPLY)
    If Err.Number <> 0 Then
        Call Log_Update(logFile, bolWriteLog, "Error: Winsock SendData")
        Exit Sub
    End If

    ' 8. update gui
    requestCount = requestCount + 1
    Form1.lblNumRequest.Caption = requestCount
    
    Call Log_Update(logRepliesFile, bolWriteRequestLog, a & aryDomainNames(0) & " " & strRemoteHostIP)
    
    strRequestsData = ""
    'Print requests, newest first.
    For i = colRequestsData.Count To 0 Step -1
        strRequestsData = strRequestsData & colRequestsData.Item(i) & vbCrLf
    Next
    
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
    
    If colRequestsData.Count > 100 Then
        colRequestsData.Remove (1)
    End If


End Sub

Private Sub btnEdit_Click()

    If FileExists(App.Path & "\" & hostFN) Then
        Shell "notepad " & App.Path & "\" & hostFN, vbNormalFocus
    Else
        MsgBox ("Missing file: " & hostFN)
        Call Log_Update(logFile, bolWriteLog, "Error: Missing file: '" & hostFN & "'", True)
    End If
    
End Sub


Private Sub Form_Load()

    Dim i As Long

    Set socket_udp = New CSocketMaster
    On Error Resume Next
    socket_udp.Protocol = sckUDPProtocol
    If Err.Number <> 0 Then
        Call Log_Update(logFile, bolWriteLog, "Error: Missing file: '" & hostFN & "'")
    End If

    If InStr(Command, "nosystemlog") Then
        bolWriteLog = False
    Else
        bolWriteLog = True
    End If

    ttl1 = 0
    ttl2 = 0
    ttl3 = 0
    ttl4 = 240
    ttl = ttl4
    
    PortNumber = 53
    bolWriteRequestLog = False
    bolCreatePTR = True
    bolGUIHide = False
    listeningIP = ""
    bolListening = False
    intTextShow = 1
    strProgrammeName = "Hoosns"
    strVersion = "0.8"
    
    iniFile = "hoosns.ini"
 
    
    Call Log_Create(logFile, bolWriteLog)
    Call Log_Update(logFile, bolWriteLog, "Starting hoosns")
    Call Parse_INI
    Call Log_Create(logRepliesFile, bolWriteRequestLog)
    
    Form1.Caption = "Listening"
    lblServerName.Caption = socket_udp.LocalHostName
    lblPortNumber.Caption = PortNumber
    lblNumRequest.Caption = requestCount
    lblNumSuccessfullRequest.Caption = successfulRequestCount
    lblTTL.Caption = ttl & " seconds"
    hostFN = "name.map"

    'setup the socket to listen to port XX (default 53),
    Call Listen_to_Requests(bolListening, True)

    'populate the hosts in the name.map file in to an array,
    Call Load_Hosts(True)
    Call Assign_Help
    
    'then show the gui.
    If bolGUIHide = True Then
        Form1.Visible = False
    Else
        Form1.Visible = True
    End If
    
End Sub

Private Sub btnStopStart_Click()

    Call Listen_to_Requests(bolListening)
    
End Sub

Private Sub btnReloadNames_Click()

    Call Load_Hosts(False)
    
End Sub

Private Sub CLose_Click()

    On Error Resume Next
    socket_udp.CloseSck
    Call Log_Update(logFile, bolWriteLog, "Closing hoosns")
    Unload Me
    
End Sub

Sub Listen_to_Requests(bol As Boolean, Optional startup As Boolean)

    If bol = False Then
        On Error Resume Next
        socket_udp.Bind PortNumber, listeningIP
        If Err.Number <> 0 Then
            Call Log_Update(logFile, bolWriteLog, "Error: unable to open port: " & PortNumber, True)
            Form1.btnStopStart.Caption = "Error"
            Form1.lblStatus.BackColor = &HFF&
            Form1.Caption = "Error"
        Else
            Call Log_Update(logFile, bolWriteLog, "Listening on UDP port: " & PortNumber, True)
            Form1.btnStopStart.Caption = "Pause"
            Form1.lblStatus.BackColor = &HFF00&
            Form1.Caption = "Listening"
            bolListening = True
        End If
        
    Else
        Err.Clear
        If startup = False Then
            On Error Resume Next
            socket_udp.CloseSck
        End If
        If Err.Number <> 0 Then
            Call Log_Update(logFile, bolWriteLog, "Error: unable to close port: " & PortNumber, True)
            Form1.btnStopStart.Caption = "Error"
            Form1.lblStatus.BackColor = &HFF&
            Form1.Caption = "Error"
        Else
            Call Log_Update(logFile, bolWriteLog, "Not listening (paused) on port number: " & PortNumber, True)
            Form1.btnStopStart.Caption = "Start"
            Form1.lblStatus.BackColor = &HFF0000
            Form1.Caption = "Not Listening"
            bolListening = False
        End If
    End If

End Sub
