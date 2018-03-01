' Script to restore saved SECS Configuration during upgrade
' (c) MKS Instruments, 2007

Private wshShell, fso
Set wshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

Private SourcePath, TargetPath, i

' Any errors and we just carry on
On Error Resume Next

' Get the common settings location
SourcePath = wshShell.RegRead("HKLM\Software\Spectra International\Process Eye\Path: Common Settings")
If Err.Number <> 0 Then WScript.Quit

TargetPath = wshShell.RegRead("HKLM\Software\Spectra International\Process Eye\Path: Application")
If Err.Number <> 0 Then WScript.Quit

SourcePath = SourcePath & "Backup\"
TargetPath = TargetPath & "Workstation\Addins\ProcSECS\"

' These are the files to move
MoveFile "PSecs1.cfg"
MoveFile "PSecs2.cfg"
MoveFile "ProcSECS.ini"
For i = 1 To 8
	MoveFile "ProcSECS" & i & ".ini"
Next


' *************************************************************************************************
' Helper function to handle the actual copying
Private Sub MoveFile (FileName)
	Dim s, t
	
	s = SourcePath & FileName
	t = TargetPath & FileName
	If Not fso.FileExists(s) Then Exit Sub
	
	If fso.FileExists(t) Then fso.DeleteFile t
	fso.MoveFile s, t
End Sub

