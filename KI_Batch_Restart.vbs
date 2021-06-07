Set objShell = CreateObject("Shell.Application")
Set wshNetwork = WScript.CreateObject( "WScript.Network" )
Set wshShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFSOdest = CreateObject("Scripting.FileSystemObject")

strComputerName = wshNetwork.computerName
strUserName = wshNetwork.userName
'strClient = wshShell.ExpandEnvironmentStrings( "%CLIENTNAME%" )

srvPath = "D:\KI_MDE"
apdPath = "/D:\KI_APD\SecurPharm"

'retValue = IsProcessRunning( strComputerName, strUserName, "KI_Batch.EXE" )
'retValue = IsProcRun( strComputerName, strUserName, "KI_Batch.EXE" )
retValue = IsProcessRun( strComputerName, strUserName, "KI_Batch.EXE" )
If  retValue = True Then
    createobject("wscript.shell").popup "Stop KI-Batch !!!", 3, "KI-Batch", 64    
    End If

createobject("wscript.shell").popup "Start KI-Batch !!!", 3, "KI-KI-Batch", 64
wshShell.Run(srvPath & "\" & "KI_Batch.exe" & " " & apdPath)
WScript.Quit 

Function IsProcessRunning( strComputer, strUser, strProcess )
    Dim Process, strObject
    IsProcessRunning = False
    strObject   = "winmgmts://" & strComputer
    For Each Process in GetObject( strObject ).InstancesOf( "win32_process" )
		colProperties = Process.GetOwner(strNameOfUser,strUserDomain)
		'Wscript.Echo "Process " & Process.Name	& " is owned by " & strUserDomain & "\" & strNameOfUser
		If  UCase( Process.name ) = UCase( strProcess ) And UCase(strUser) = UCase(strNameOfUser) Then
			Process.Terminate()
			
			IsProcessRunning = True
			Exit Function
		End If
    Next
End Function

Function IsProcRun( strComputer, strUser, strProcess )
	Dim objWMIService, objProcess, colProcess
	
	Set objWMIService = GetObject("winmgmts:" _
			& "{impersonationLevel=impersonate}!\\" _ 
			& strComputer & "\root\cimv2")
	Set colProcess = objWMIService.ExecQuery _
			("Select * from Win32_Process Where Name = " & strProcess )
			
	For Each objProcess in colProcess
			objProcess.Terminate()
	Next 
	WSCript.Echo "Just killed process " & strProcess _
					& " on " & strComputer
	'WScript.Quit 
	IsProcRun = True
End Function

Function IsProcessRun( strComputer, strUser, strProcess )
    Dim Process, strObject
    IsProcessRun = False
    strObject   = "winmgmts://" & strComputer
    For Each Process in GetObject( strObject ).InstancesOf( "win32_process" )
		'colProperties = Process.GetOwner(strNameOfUser,strUserDomain)
		'Wscript.Echo "Process " & Process.Name	& " is owned by " & strUserDomain & "\" & strNameOfUser
		If  UCase( Process.name ) = UCase( strProcess ) Then
			Process.Terminate()
			
			IsProcessRun = True
			Exit Function
		End If
    Next
End Function
