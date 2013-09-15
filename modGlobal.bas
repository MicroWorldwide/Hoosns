Attribute VB_Name = "modGlobal"
Public numHosts As Long
Public colHosts As New Collection
Public colPTRHosts As New Collection
Public Const modHosts As Long = 3
Public PortNumber As Long
Public bolWriteLog As Boolean
Public bolWriteRequestLog As Boolean
Public bolCreatePTR As Boolean
Public strHostsData As String
Public strRequestsData As String
Public colRequestsData As New Collection
Public IP As String
Public hostFN As String
Public bolListening As Boolean
Public bolGUIHide As Boolean
Public ttl1 As Long
Public ttl2 As Long
Public ttl3 As Long
Public ttl4 As Long
Public requestCount As Long
Public successfulRequestCount As Long
Public ttl As Long
Public listeningIP As String
Public strHelp As String
Public strProgrammeName As String
Public strVersion As String
Public iniFile As String
Public intTextShow As Long

