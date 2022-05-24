Option Explicit

Dim owsh : Set owsh = CreateObject("wscript.shell")
Dim oSAP : Set oSAP = New SAPLauncherV2
Dim oConnFQ2,oSessFQ2
Dim oConnFQ2a,oSessFQ2a
'***********************
'Create connection block
'***********************
On Error Resume Next
	err.Clear
	oSAP.CreateConnectionSession "FQ2 - SAP_VGMF ERP TEST [1010]", oConnFQ2, oSessFQ2
	If err.number <> 0 Then
		'Send admin message here
		debug.WriteLine "Error number: " & err.number & ". Error description: " & err.Description
	End If 
On Error GoTo 0 

WScript.Sleep 5000

WScript.Quit(1)











Class SAPLauncherV2

	' ============= Private members ===========================
	Private oAPP_
	Private oXML_
	Private oFSO_
	Private strLandscape_
	Private boolLogonRunning_ ' Indicates whether SAPlogon is already running i.e sapgui.scriptingctrl.1 has already been instantiated
	Private dictSessionsToClose_  ' Holds sessions that should be closed by Class_Terminate()
	Private dictConnectionsToClose_ ' Holds connections that should be closed by Class_Terminate()
	' ============== Constructor & Destructor ==================
	
	Private Sub Class_Initialize
		
		On Error Resume Next
		Set dictSessionsToClose_ = CreateObject("Scripting.Dictionary")
		Set dictConnectionsToClose_ = CreateObject("Scripting.Dictionary")
		Set oAPP_ = GetObject("SAPGUI").GetScriptingEngine
		On Error GoTo 0
		
		If Not IsObject(oAPP_) Then
			debug.WriteLine "Instantiating Sapgui.ScriptingCtrl.1"
			Set oAPP_ = CreateObject("Sapgui.ScriptingCtrl.1")
			debug.WriteLine "Calling RegisterROT()"
			oAPP_.RegisterROT
			boolLogonRunning_ = False
		Else
			boolLogonRunning_ = True
			debug.WriteLine "Sapgui.ScriptingCtrl.1 already there"
		End If 
		
		
		Set oXML_ = CreateObject("MSXML2.DOMDocument")
		Set oFSO_ = CreateObject("Scripting.FileSystemObject")
		
		
		If Not IsObject(oAPP_) Or Not IsObject(oXML_) Or Not IsObject(oFSO_) Then
			err.Raise 100,"SAPLauncherV2.Class_Initialize","Error creating objects [Sapgui.ScriptingCtrl.1 | MSXML2.DOMDocument | Scripting.FileSystemObject]"
		End If 
		
	End Sub
	
	Private Sub Class_Terminate
		Dim session__,connection__
		If dictSessionsToClose_.Count = 0 Then
			debug.WriteLine "No owned sessions to close"
		ElseIf dictConnectionsToClose_.Count = 0 Then
			debug.WriteLine "No owned connections to close"
		End If 
		
		For Each session__ In dictSessionsToClose_.Keys
			debug.WriteLine "Closing owned session: " & dictSessionsToClose_.Item(session__).Description & " - " & session__
			dictSessionsToClose_.Item(session__).CloseSession(session__)
		Next
		
		For Each connection__ In dictConnectionsToClose_.Keys
			debug.WriteLine "Closing owned connection: " & dictConnectionsToClose_.Item(connection__).Description
			dictConnectionsToClose_.Item(connection__).CloseConnection
		Next 
	End Sub

	
	'************************************************************
	' Function creates a connection and session to the specified
	' system. Returns connection and session objects
	'***********************************************************
	Public Function CreateConnectionSession(strSystemDescription, ByRef oConn, ByRef oSess)
		debug.WriteLine "CreateConnectionSession() called"
		Dim numCount__
		Dim oConn__ 	 ' Private variable holding connection objects. Each oConn__ is a child of oAPP
		Dim oConnChild__ ' Private variable holding session objects. Each oConnChild__ is a child of oConn__
		If boolLogonRunning_ Then ' SAP logon running. Look for an existing connection and returnd oConn__ object
			If oAPP_.Connections.Count > 0 Then ' SAP logon running and at least one connection is opened
				debug.WriteLine "->Looking for a suitable existing connection (" & strSystemDescription & ")"
				For Each oConn__ In oAPP_.Children
					If oConn__.Description = strSystemDescription Then
						debug.WriteLine "->Suitable connection found (" & oConn__.Description & ")"
						Set oConn = oConn__
						numCount__ = oConn__.Sessions.Count
						oConn__.Sessions.Item(0).CreateSession
						Do While oConn__.Sessions.Count <> numCount__ + 1 ' Wait until session count 
							WScript.Sleep 500
						Loop 
	    				Set oSess = oConn__.Sessions.Item(oConn__.Sessions.Count - 1)
	    				dictSessionsToClose_.Add oConn__.Sessions.Item(oConn__.Sessions.Count - 1).Id, oConn__ ' Save ID of the session that should be closed during termination phase
						Exit Function 
					Else
						debug.WriteLine "->No suitable connection found. Opening a new connection (" & strSystemDescription & ")"
						Set oConn = oAPP_.OpenConnection(strSystemDescription)
						Set oSess = oConn.Sessions.Item(0)
						dictSessionsToClose_.Add oConn.Sessions.Item(0).Id, oConn ' Save ID of the session that should be closed during termination phase
						dictConnectionsToClose_.Add oConn.Description, oConn ' Save connection object that should be closed during termination phase
						Exit Function 
					End If  
				Next
			Else ' SAP logon running but 0 connections
				debug.WriteLine "->No connections opened yet. Opening a new connection (" & strSystemDescription & ")"
				Set oConn = oAPP_.OpenConnection(strSystemDescription)
				Set oSess = oConn.Sessions.Item(0)
				dictSessionsToClose_.Add oConn.Sessions.Item(0).Id, oConn ' Save ID of the session that should be closed during termination phase
				If Not dictConnectionsToClose_.Exists(oConn.Description) Then
					dictConnectionsToClose_.Add oConn.Description, oConn ' Save connection object that should be closed during termination phase
				End If 
				Exit Function 
			End If 
		Else ' SAP logon not running. 
			If oAPP_.Connections.Count > 0 Then 
				debug.WriteLine "->Looking for a suitable existing connection (" & strSystemDescription & ")"
				For Each oConn__ In oAPP_.Children
					If oConn__.Description = strSystemDescription Then
						debug.WriteLine "->Suitable connection found (" & oConn__.Description & ")"
						Set oConn = oConn__
						numCount__ = oConn__.Sessions.Count
						oConn__.Sessions.Item(0).CreateSession
						Do While oConn__.Sessions.Count <> numCount__ + 1 ' Wait until session count 
							WScript.Sleep 500
						Loop 
	    				Set oSess = oConn__.Sessions.Item(oConn__.Sessions.Count - 1)
	    				dictSessionsToClose_.Add oConn__.Sessions.Item(oConn__.Sessions.Count - 1).Id, oConn__ ' Save ID of the session that should be closed during termination phase
						Exit Function 
					Else
						debug.WriteLine "->No suitable connection found. Opening a new connection (" & strSystemDescription & ")"
						Set oConn = oAPP_.OpenConnection(strSystemDescription)
						Set oSess = oConn.Sessions.Item(0)
						dictSessionsToClose_.Add oConn.Sessions.Item(0).Id, oConn ' Save ID of the session that should be closed during termination phase
						If Not dictConnectionsToClose_.Exists(oConn.Description) Then
							dictConnectionsToClose_.Add oConn.Description, oConn ' Save connection object that should be closed during termination phase
						End If 
						Exit Function 
					End If  
				Next
			Else
				debug.WriteLine "->No suitable connection found. Opening a new connection (" & strSystemDescription & ")"
				Set oConn = oAPP_.OpenConnection(strSystemDescription)
				Set oSess = oConn.Sessions.Item(0)
				dictSessionsToClose_.Add oConn.Sessions.Item(0).Id, oConn ' Save ID of the session that should be closed during termination phase
				If Not dictConnectionsToClose_.Exists(oConn.Description) Then
					dictConnectionsToClose_.Add oConn.Description, oConn ' Save connection object that should be closed during termination phase
				End If 
				Exit Function
			End If  
		End If 
			
	End Function 
	
	
End Class 