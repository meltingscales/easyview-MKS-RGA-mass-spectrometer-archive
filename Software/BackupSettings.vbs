' Installation script that runs during upgrade to
'	preserve the addin settings for email and ToolLink
'	copy RTCLog.text
'	copy SECS configuration
'(c) MKS Instruments, 2007

Private wshShell, fso
Private SECSPath, RTCPath, BackupPath

Set wshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Make sure that we just continue through errors
On Error Resume Next
If Not GetPaths() Then WScript.Quit


' *************************************************************************
' This is all we do
CopyAddinSettings
CopyRTCLog
CopySECSConfiguration

' Now we have finished
' *************************************************************************

' *******************************************************************************************
' Main Methods
' ***************
Private Sub CopyAddinSettings
	Dim eMailServer, eMailPort
	Dim ToolLinkServer, ToolLinkPort
	Dim ToolLinkLogFile , ToolLinkLogFileSize
	
	On Error Resume Next
	
	' Email settings
	eMailServer = wshShell.RegRead("HKLM\Software\Spectra International\Process Eye\Addins\EMail\SMTP Server")
	If Err.Number <> 0 Then eMailServer = Empty
	Err.Clear
	
	eMailPort = wshShell.RegRead("HKLM\Software\Spectra International\Process Eye\Addins\EMail\SMTP Port")
	If Err.Number <> 0 Then eMailPort = Empty
	Err.Clear
	
	' ToolLink settings
	ToolLinkServer = wshShell.RegRead("HKLM\Software\Spectra International\Process Eye\Addins\ToolLink\Server Address")
	If Err.Number <> 0 Then ToolLinkServer = Empty
	Err.Clear
	
	ToolLinkPort = wshShell.RegRead("HKLM\Software\Spectra International\Process Eye\Addins\ToolLink\Server Port")
	If Err.Number <> 0 Then ToolLinkPort = Empty
	Err.Clear
	
	ToolLinkLogFile = wshShell.RegRead("HKLM\Software\Spectra International\Process Eye\Addins\ToolLink\Log File")
	If Err.Number <> 0 Then ToolLinkLogFile = Empty
	Err.Clear
	
	ToolLinkLogFileSize = wshShell.RegRead("HKLM\Software\Spectra International\Process Eye\Addins\ToolLink\Log File Size")
	If Err.Number <> 0 Then ToolLinkLogFileSize = Empty
	Err.Clear
	
	' Write the settings out
	wshShell.RegDelete "HKLM\Software\Spectra International\Process Eye\Install\SMTP Server" 
	If Not IsEmpty(eMailServer) Then wshShell.RegWrite "HKLM\Software\Spectra International\Process Eye\Install\SMTP Server", eMailServer
	
	wshShell.RegDelete "HKLM\Software\Spectra International\Process Eye\Install\SMTP Port" 
	If Not IsEmpty(eMailPort) Then wshShell.RegWrite "HKLM\Software\Spectra International\Process Eye\Install\SMTP Port", eMailPort
	
	wshShell.RegDelete "HKLM\Software\Spectra International\Process Eye\Install\ToolLink Server" 
	If Not IsEmpty(ToolLinkServer) Then wshShell.RegWrite "HKLM\Software\Spectra International\Process Eye\Install\ToolLink Server", ToolLinkServer
	
	wshShell.RegDelete "HKLM\Software\Spectra International\Process Eye\Install\ToolLink Port" 
	If Not IsEmpty(ToolLinkPort) Then wshShell.RegWrite "HKLM\Software\Spectra International\Process Eye\Install\ToolLink Port", ToolLinkPort
	
	wshShell.RegDelete "HKLM\Software\Spectra International\Process Eye\Install\ToolLink LogFile" 
	If Not IsEmpty(ToolLinkLogFile) Then wshShell.RegWrite "HKLM\Software\Spectra International\Process Eye\Install\ToolLink LogFile", ToolLinkLogFile
	
	wshShell.RegDelete "HKLM\Software\Spectra International\Process Eye\Install\ToolLink LogFileSize" 
	If Not IsEmpty(ToolLinkLogFileSize) Then wshShell.RegWrite "HKLM\Software\Spectra International\Process Eye\Install\ToolLink LogFileSize", ToolLinkLogFileSize
End Sub

Private Sub CopyRTCLog
	Dim s, t, i, bDone

	s = RTCPath & "RTCLog.text"
	If Not fso.FileExists(s) Then Exit Sub
	
	' we keep multiple copies of the log so we have to find a filename that has not been used
	i = 1
	bDone = False
	Do
		t = BackupPath & "RTCLog_" & i & ".text"
		If Not fso.FileExists(t) Then
			fso.CopyFile s, t
			bDone = True
		End If
		i = i + 1
	Loop Until bDone
End Sub

Private Sub CopySECSConfiguration
	Dim i
	
	CopySECSFile "PSecs1.cfg"
	CopySECSFile "PSecs2.cfg"
	CopySECSFile "ProcSECS.ini"
	For i = 1 To 8
		CopySECSFile "ProcSECS" & i & ".ini"
	Next
End Sub

' *****************************************************************************************
' Helper Methods
' *****************
Private Function GetPaths
	Dim AppPath
	Dim p, s
	
	GetPaths = False
	On Error Resume Next
	AppPath = wshShell.RegRead("HKLM\Software\Spectra International\Process Eye\Path: Application")
	BackupPath = wshShell.RegRead("HKLM\Software\Spectra International\Process Eye\Path: Common Settings")
	BackupPath = BackupPath & "Backup"
	If Err.Number <> 0 Then Exit Function
	
	fso.CreateFolder BackupPath
	BackupPath = BackupPath & "\"
	
	SECSPath = AppPath & "Workstation\Addins\ProcSECS\"
	
	s = Left(AppPath, Len(AppPath) - 1)	
	p = InstrRev(s, "\")
	RTCPath = Left(s, p) & "webroot\RealTimeConfiguration\"
	
	GetPaths = True
End Function

Private Sub CopySECSFile (FileName)
	Dim s, t
	
	s = SECSPath & FileName
	t = BackupPath & FileName
	If Not fso.FileExists(s) Then Exit Sub
	
	fso.CopyFile s, t, True
End Sub

