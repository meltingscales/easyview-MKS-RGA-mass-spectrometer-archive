' Installation script to restore preserved addin settings for email and ToolLink
'(c) MKS Instruments, 2007

Private wshShell
Set wshShell = CreateObject("WScript.Shell")

Private eMailServer, eMailPort
Private ToolLinkServer, ToolLinkPort
Private ToolLinkLogFile , ToolLinkLogFileSize

' Make sure we continue through errors
On Error Resume Next

' ********************************************************************************************************
' Read EMail configuration
' **************************
eMailServer = wshShell.RegRead("HKLM\Software\Spectra International\Process Eye\Install\SMTP Server")
If Err.Number <> 0 Then eMailServer = Empty
wshShell.RegDelete "HKLM\Software\Spectra International\Process Eye\Install\SMTP Server" 
Err.Clear

eMailPort = wshShell.RegRead("HKLM\Software\Spectra International\Process Eye\Install\SMTP Port")
If Err.Number <> 0 Then eMailPort = Empty
wshShell.RegDelete "HKLM\Software\Spectra International\Process Eye\Install\SMTP Port" 
Err.Clear


' ********************************************************************************************************
' Read ToolLink configuration
' *****************************
ToolLinkServer = wshShell.RegRead("HKLM\Software\Spectra International\Process Eye\Install\ToolLink Server")
If Err.Number <> 0 Then ToolLinkServer = Empty
wshShell.RegDelete "HKLM\Software\Spectra International\Process Eye\Install\ToolLink Server" 
Err.Clear

ToolLinkPort = wshShell.RegRead("HKLM\Software\Spectra International\Process Eye\Install\ToolLink Port")
If Err.Number <> 0 Then ToolLinkPort = Empty
wshShell.RegDelete "HKLM\Software\Spectra International\Process Eye\Install\ToolLink Port" 
Err.Clear

ToolLinkLogFile = wshShell.RegRead("HKLM\Software\Spectra International\Process Eye\Install\ToolLink LogFile")
If Err.Number <> 0 Then ToolLinkLogFile = Empty
wshShell.RegDelete "HKLM\Software\Spectra International\Process Eye\Install\ToolLink LogFile" 
Err.Clear

ToolLinkLogFileSize = wshShell.RegRead("HKLM\Software\Spectra International\Process Eye\Install\ToolLink LogFileSize")
If Err.Number <> 0 Then ToolLinkLogFileSize = Empty
wshShell.RegDelete "HKLM\Software\Spectra International\Process Eye\Install\ToolLink LogFileSize" 
Err.Clear


' ****************************************************************************************************************
' Now save all the found data back to the correct location
' **********************************************************
If Not IsEmpty(eMailServer) Then wshShell.RegWrite "HKLM\Software\Spectra International\Process Eye\Addins\EMail\SMTP Server", eMailServer
If Not IsEmpty(eMailPort) Then wshShell.RegWrite "HKLM\Software\Spectra International\Process Eye\Addins\EMail\SMTP Port", eMailPort
If Not IsEmpty(ToolLinkServer) Then wshShell.RegWrite "HKLM\Software\Spectra International\Process Eye\Addins\ToolLink\Server Address", ToolLinkServer
If Not IsEmpty(ToolLinkPort) Then wshShell.RegWrite "HKLM\Software\Spectra International\Process Eye\Addins\ToolLink\Server Port", ToolLinkPort
If Not IsEmpty(ToolLinkLogFile) Then wshShell.RegWrite "HKLM\Software\Spectra International\Process Eye\Addins\ToolLink\Log File", ToolLinkLogFile
If Not IsEmpty(ToolLinkLogFileSize) Then wshShell.RegWrite "HKLM\Software\Spectra International\Process Eye\Addins\ToolLink\Log File Size", ToolLinkLogFileSize

