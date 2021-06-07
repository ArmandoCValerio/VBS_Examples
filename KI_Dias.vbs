Set objShell = CreateObject("Shell.Application")
Set wshNetwork = WScript.CreateObject( "WScript.Network" )
Set wshShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFSOdest = CreateObject("Scripting.FileSystemObject")

strComputerName = wshNetwork.computerName
strUserName = wshNetwork.userName
strClient = wshShell.ExpandEnvironmentStrings( "%CLIENTNAME%" )

srvPath = "C:\KI_ProCom\KI_APS\"
apdPath = "/C:\KI_ProCom\KI_APD"
wstPath = "C:\KI_ProCom\KI_APSsrv\" & strClient
login   = "networkuser"

'wshShell.currentdirectory = apdWRK

WScript.Echo("Ihr Computer heißt: " & strComputerName)
WScript.Echo("Der WSTA-Pfad heißt: " & wstPath)
WScript.Echo("Der SRV-Pfad heißt: " & srvPath)
WScript.Echo("Ihr Name ist: " & strUserName)
WScript.Echo("ClientName: " & strClient)

retValue = IsProcessRunning( strComputerName, strUserName, "KI_Dias.EXE" )
If  retValue = True Then
    wshShell.Run(wstPath & "\" & "KI_Dias.EXE")
    WScript.Quit 
    End If

If Not objFSO.FolderExists(wstPath) Then objFSO.CreateFolder(wstPath)
If Not objFSO.FolderExists(wstPath) Then 
   createobject("wscript.shell").popup "Error! Destination Path?", 0, "KI-Application", 64
   WScript.Quit 
end If

If Not objFSO.FolderExists(wstPath) Then 
   createobject("wscript.shell").popup "Cancel! Source Path?", 0, "KI-Application", 64
   WScript.Quit 
End If

retValue = CopyFiles(srvPath,wstPath)
If  retValue = False then
	createobject("wscript.shell").popup "Start KI-Application !!!", 3, "KI-Application", 64
	wshShell.Run(wstPath & "\" & "KI_Dias.exe" & " " & apdPath & " /LOGON=" & login)
Else
    WScript.Echo "Error : " & Err.Number & Err.Description
    End If

WScript.Quit 


Function CopyFiles(srcPath, destPath)
	On Error Resume Next
	createobject("wscript.shell").popup "Start KI-UPDATE!", 3, "KI-Application", 64
	For Each sFile In objFSO.GetFolder(srcPath).Files 
		If Err.Number = 0 then
			If Not objFSO.FileExists(destPath & "\" & objFSO.GetFileName(sFile)) then
				objFSO.GetFile(sFile).Copy destPath & "\" & objFSO.GetFileName(sFile),True 
				'WScript.Echo "Copying : " & Chr(34) & objFSO.GetFileName(sFile) & Chr(34) & " to " & destPath 
			Else
				If  DateDiff("s",objFSOdest.GetFile(destPath & "\" & objFSO.GetFileName(sFile)).DateLastModified,objFSO.GetFile(sFile).DateLastModified) > 0 Then
					objFSO.GetFile(sFile).Copy destPath & "\" & objFSO.GetFileName(sFile),True 
					'WScript.Echo "Copying : " & Chr(34) & objFSO.GetFileName(sFile) & Chr(34) & " to " & destPath 
				End If
			End If
		Else
			CopyFiles = True
		End If 
	next
End Function

Function IsProcessRunning( strComputer, strUser, strProcess )
    Dim Process, strObject
    IsProcessRunning = False
    strObject   = "winmgmts://" & strComputer
    For Each Process in GetObject( strObject ).InstancesOf( "win32_process" )
		colProperties = Process.GetOwner(strNameOfUser,strUserDomain)
		'Wscript.Echo "Process " & Process.Name	& " is owned by " & strUserDomain & "\" & strNameOfUser
		If  UCase( Process.name ) = UCase( strProcess ) And UCase(strUser) = UCase(strNameOfUser) Then
			IsProcessRunning = True
			Exit Function
		End If
    Next
End Function

