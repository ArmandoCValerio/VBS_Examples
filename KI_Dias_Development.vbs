' *********************************************************************
' 
' 				Armando Coelho V A L É R I O
'       			Business Information Systems
'       			2019/11/01  
'
' *********************************************************************

Set objShell = CreateObject("Shell.Application")
Set wshNetwork = WScript.CreateObject( "WScript.Network" )
Set wshShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFSOdest = CreateObject("Scripting.FileSystemObject")

strComputerName = wshNetwork.computerName
strUserName = wshNetwork.userName
strClient = wshShell.ExpandEnvironmentStrings( "%CLIENTNAME%" )

' Get count of arguments
'Argumente = WScript.Arguments.Count
'If Argumente > 0 Then
'   Set Args = WScript.Arguments
'   For i = 0 to Args.Count - 1
'       Params = Params + Chr(10) + Args(i)
'   Next
'   createobject("wscript.shell").popup "Parameter? " & Params, 3, "Parameter", 64
'wshShell.CurrentDirectory = WScript.Arguments(0)
'wstPath = wshShell.CurrentDirectory



createobject("wscript.shell").popup "Start! #1", 2, "KI-Application", 64

srvPath = "Z:\KI_APSws\"
apdPath = "/Z:\KI_APD"
wstPath = "C:\KI_APSws\" '& strClient
curPath = "C:\KI_APSws"
'login   = "networkuser"
login   = "workstation"

wshshell.currentdirectory = curPath

'WScript.Echo("Ihr Computer heißt: " & strComputerName)
'WScript.Echo("Der WSTA-Pfad heißt: " & wstPath)
'WScript.Echo("Der SRV-Pfad heißt: " & srvPath)
'WScript.Echo("Der Current-Pfad heißt: " & curPath)
'WScript.Echo("Ihr Name ist: " & strUserName)
'WScript.Echo("ClientName: " & strClient)

strProcess = "KI_Dias.EXE"

'If  isProcessRunning(strComputerName,strProcess) then
'	wscript.echo strProcess & " is running on computer '" & strComputerName & "'"
'	success =  wshShell.AppActivate(strProcess)
'	If  success then 
'		wshShell.SendKeys ("% x")
'	Else
'		createobject("wscript.shell").popup "KI-Application schon aktiv!!!", 3, "KI-Application", 64
'	End If
'	'wshShell.SendKeys "% r"
'   WScript.Quit 
'Else
'	wscript.echo strProcess & " is NOT running on computer '" & strComputerName & "'"
'	End If

retValue = IsProcessRunning( strComputerName, strProcess)
If  retValue = True Then
    success =  wshShell.AppActivate(strProcess)
	If  success then 
		wshShell.SendKeys ("% x")
	Else
		createobject("wscript.shell").popup "KI-Application schon aktiv!!!", 3, "KI-Application", 64
	End If
	'wshShell.Run(wstPath & "\" & "KI_Dias.EXE")
    WScript.Quit 
    End If

'retValue = IsProcessRunning( strComputerName, strUserName, "KI_Dias.EXE" )
'If  retValue = True Then
'    wshShell.Run(wstPath & "\" & "KI_Dias.EXE")
'    WScript.Quit 
'    End If
	
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

'Function IsProcessRunning( strComputer, strUser, strProcess )
'    Dim Process, strObject
'    IsProcessRunning = False
'    strObject   = "winmgmts://" & strComputer
'    For Each Process in GetObject( strObject ).InstancesOf( "win32_process" )
'		colProperties = Process.GetOwner(strNameOfUser,strUserDomain)
'		Wscript.Echo "Process " & Process.Name	& " is owned by " & strUserDomain & "\" & strNameOfUser
'		If  UCase( Process.name ) = UCase( strProcess ) And UCase(strUser) = UCase(strNameOfUser) Then
'			IsProcessRunning = True
'			Exit Function
'		End If
'   Next
'End Function

' Function to check if a process is running
function isProcessRunning(byval strComputer,byval strProcessName)

	Dim objWMIService, strWMIQuery

	strWMIQuery = "Select * from Win32_Process where name like '" & strProcessName & "'"
	
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" _ 
			& strComputer & "\root\cimv2") 

	if objWMIService.ExecQuery(strWMIQuery).Count > 0 then
		isProcessRunning = true
	else
		isProcessRunning = false
	end if

end function

