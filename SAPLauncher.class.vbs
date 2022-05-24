Dim sl,system,client,shell
Dim ss ' GLOBAL VARIABLE NEEDS TO BE THIS EXACT NAME or rewrite it in the GetSession subroutine
system = WScript.Arguments.Item(0)
client = WScript.Arguments.Item(1)
Set shell = CreateObject("Wscript.Shell")
Set sl = New SAPLauncher

sl.SetClientName = WScript.Arguments.Item(1)
sl.SetSystemName = WScript.Arguments.Item(0)
sl.SetLocalXML = shell.ExpandEnvironmentStrings("%APPDATA%") & "\SAP\Common\SAPUILandscape.xml"
sl.CheckSAPLogon
sl.FindSAPSession

If sl.SessionFound Then 
	sl.GetSession
	debug.WriteLine "Opened connection to: " & ss.Info.SystemName
End If 















Class SAPLauncher
	
	Private oHTTP
	Private oXML
	Private oWSH
	Private oFSO
	Private oSAPGUI
	Private oAPP
	Private oCON
	Private oSES
	Private strGlobalURL
	Private strLocalLandscapePATH
	Private boolSAPRunning  		' Indicates whether SPA Logon is runniied files
	Private boolSessionFound		' Set to true if session was found or created
	Private strSSN 					' Sap System Name e.g FQ2
	Private strSCN  	    		' Sap Client Name e.g. 105
	Private strSSD          		' Sap System Description. This string is found in the local landscape xml and used to connect to the sap system


	
	
	
	' ============== Constructor & Destructor ==================
	Private Sub Class_Initialize
	
		Set oHTTP = CreateObject("MSXML2.XMLHTTP")
		Set oXML = CreateObject("MSXML2.DOMDocument")
		Set oFSO = CreateObject("scripting.filesystemobject")
		Set oWSH = CreateObject("wscript.shell")
		oSAPGUI = Null
		oAPP = Null
		oCON = Null
		oSES = Null 
		strSSN = Null
		strSCN = Null
		strGlobalURL = Null
		strLocalLandscapePATH = Null
		strSSD = Null
		boolSAPRunning = False
		boolSessionFound = False
		

	End Sub
	
	Private Sub Class_Terminate
	
	End Sub
	
	' ============ P U B L I C  &  P R I V A T E   M E T H O D S  &   S U B R O U T I N E S ===========
	
	

	' ---------- CheckSAPLogon
	Public Sub CheckSAPLogon
	
		Dim oWmi,colProc,proc,oSAP,waitfor
		Set oWmi = GetObject("winmgmts:\\.\root\cimv2")
		Set colProc = oWmi.ExecQuery("SELECT Name, ProcessId FROM Win32_Process")
		
		On Error Resume Next 
		For Each proc In colProc
			If InStr(LCase(proc.Name),"saplogon") > 0 Then 
				Do While True 
					Set oSAPGUI = GetObject("SAPGUI") ' Wait until the object is instantiated
						If IsObject(oSAPGUI) Then
							boolSAPRunning = True
							Exit Sub  ' At this point we can safely assume that SAPLogon is running and SAPGUI object is available
					End If 
				Loop
			End If 
		Next 
		
		On Error GoTo 0 ' Reenable error handling
			
		'Start SAPLogon and open system passed in the command line parameter
		Set proc = oWSH.Exec("C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe")
		Set colProc = oWmi.ExecQuery("SELECT Name, ProcessId FROM Win32_Process")
		
		On Error Resume Next 
		For Each proc In colProc
			If InStr(LCase(proc.Name),"saplogon") > 0 Then 
				Do While True 
					Set oSAPGUI = GetObject("SAPGUI") ' Wait until the object is instantiated
						If IsObject(oSAPGUI) Then
							boolSAPRunning = True
							Exit Sub  ' At this point we can safely assume that SAPLogon is running and SAPGUI object is available
					End If 
				Loop
			End If 
		Next 
		
		On Error GoTo 0 ' Reenable error handling
		

	End Sub 
	
	
	
	' ---------- FindSAPSession
	Public Sub FindSAPSession
		Dim waitPeriod,waitTurns,currentTurn
		waitPeriod = 5000 ' miliseconds
		waitTurns = 5 ' 5 x 5000 = 20000 ms / 20 s
		currentTurn = 1
		
		If Not boolSAPRunning Then
			oSES = Null
			Exit Sub 
		End If  
		
		FindSAPSystemDescription
		
		If IsNull(strSSD) Then
			oSES = Null
			Exit Sub
		End If
		 
		Set oAPP = oSAPGUI.GetScriptingEngine
		Select Case oAPP.Children.Count
	
			Case 0 ' No open connections exist
				For currentTurn = 1 To waitTurns
					Set oCON = oAPP.OpenConnection(strSSD,True,False) ' Open a new connection asynchronously
					On Error Resume Next
					Set oSES = oCON.Children(0) ' Attach to the first session
					On Error GoTo 0
					If Not IsObject(oSES) Or IsNull(oSES) Then
						Debug.WriteLine "No session found, waiting " & currentTurn & " out of " & waitTurns & " turns"
						WScript.Sleep waitPeriod
					ElseIf IsObject(oSES) And IsNull(oSES) Then
						Debug.WriteLine "No session found, waiting " & currentTurn & " out of " & waitTurns & " turns"
						WScript.Sleep waitPeriod
					Else
						Exit For 
					End If 	
				Next
				If Not IsObject(oSES) Or IsNull(oSES) Then
						Debug.WriteLine "No session found after 5 retries"
						boolSessionFound = False
						Exit Sub
				End If 
				
				If IsObject(oSES) And Not IsNull(oSES) Then
					If InStr(oSES.findById("wnd[0]/sbar/pane[0]").text,"No user exists") > 0 Then
						oCON.CloseConnection
						boolSessionFound = False 
						debug.WriteLine "Session found: " & boolSessionFound
						Exit Sub 
					End If 
					boolSessionFound = True
					debug.WriteLine "Session found: " & boolSessionFound
					Exit Sub 
				End If
				
				oCON.CloseConnection
				boolSessionFound = False
				debug.WriteLine "Session found: " & boolSessionFound
				Exit Sub 
				
		
			Case Else ' Atleast one connection exists
		
				For Each oCON In oAPP.Children ' connections
					For Each oSES In oCON.Children ' sessions
						If LCase(oSES.Info.SystemName) = LCase(strSSN) Then
							If InStr(oSES.findById("wnd[0]/sbar/pane[0]").text,"No user exists") > 0 Then
								oCON.CloseConnection
								boolSessionFound = False 
								debug.WriteLine "Session found: " & boolSessionFound
								Exit Sub 
							End If 
							boolSessionFound = True
							debug.WriteLine "Session found: " & boolSessionFound
							Exit Sub ' Stop here. We found our desired system. oCON and oSES objects hold our target system
						End If 
					Next
				Next
			
				
			End Select 
			
			Set oCON = oAPP.OpenConnection(strSSD,True,False) ' Open a new connection asynchronously
			On Error Resume Next 
			Set oSES = oCON.Children(0) ' Attach to the first session
			On Error GoTo 0
			
			If IsObject(oSES) And Not IsNull(oSES) Then
				If InStr(oSES.findById("wnd[0]/sbar/pane[0]").text,"No user exists") > 0 Then
					oCON.CloseConnection
					boolSessionFound = False 
					debug.WriteLine "Session found: " & boolSessionFound
					Exit Sub 
				End If 
					boolSessionFound = True
					debug.WriteLine "Session found: " & boolSessionFound
					Exit Sub
			End If 
			
			oCON.CloseConnection
			debug.WriteLine "Session found: " & boolSessionFound
			
	End Sub 

	
	
	' --------- FindSAPSystemDescription
	Private Sub FindSAPSystemDescription
	
		Dim n_ChildNodes,n_ChildNode,uuid,i,j
		'oHTTP.open "GET",strGlobalURL,False
		'oHTTP.send
		
		'If oHTTP.status <> 200 Then
		'	strSSD = Null
		'	Exit Sub 
		'End If 
		
		'oXML.load(oHTTP.responseXML)
	
		'Set n_ChildNodes = oXML.getElementsByTagName("Messageserver")
	
		'For Each n_ChildNode In n_ChildNodes
		'	If LCase(n_ChildNode.attributes.getNamedItem("name").text) = LCase(strSSN) Then
		'		uuid = n_ChildNode.attributes.getNamedItem("uuid").text ' We found the uuid of the target system
		'		Exit For
		'	End If 
		
		'Next
	
		oXML.load(strLocalLandscapePATH) ' Locally stored XML

		Set n_ChildNodes = oXML.getElementsByTagName("Landscape")
	
		For Each n_ChildNode In n_ChildNodes 
			For Each i In n_ChildNode.childNodes
				If i.baseName = "Services" Then 
					Set n_ChildNodes = i.childNodes
					On Error Resume Next
					For Each j In n_ChildNodes
						
						If Left(LCase(j.attributes.getNamedItem("name").text),3) = LCase(strSSN) Then 
							debug.WriteLine "SAP system name: " & strSSN
							strSSD = j.attributes.getNamedItem("name").text
							debug.WriteLine "SAP system description: " & strSSD
							CheckSAPLogon
							Exit Sub  
						'If j.attributes.getNamedItem("msid").text = uuid Then ' We have a match
						'	strSSD = j.attributes.getNamedItem("name").text
						'	CheckSAPLogon                                                                                                                                                                                                            
						'	Exit Sub ' No need to continue
						End If
					Next
				End If 
			Next
		Next
		strSSD = Null ' Not found
	End Sub 
	
	Public Sub GetSession
	
		Set ss = oSES

	End Sub 
		

	' ================= P R O P E R T I E S ====================
	Public Property Get SAPLogonRunning
		SAPLogonRunning = boolSAPRunning
	End Property 	
		
	Public Property Get SAPSessionExists
		If boolSAPRunning And Not IsNull(oSES) Then
			SAPSessionExists = True
		Else 
			SAPSessionExists = False
		End If
	End Property 
	
	Public Property Get SAPsysName
		SAPsysName = strSSN
	End Property 
	
	Public Property Get SAPcliName
		SAPcliName = strSCN
	End Property 
	
	Public Property Get LandscapeURL
		LandscapeURL = strGlobalURL
	End Property 
	
	
	Public Property Get SAPsysDescription
		SAPsysDescription = strSSD
	End Property 
	
	Public Property Get GetGlobalURL
	
		GetGlobalURL = strGlobalURL
		
	End Property 
	
	Public Property Let SetGlobalURL(url)
	
		strGlobalURL = url
		
	End Property 
	
	Public Property Let SetLocalXML(xml)
	
		strLocalLandscapePATH = xml
		
	End Property 
	
	Public Property Get GetLocalXML
	
		GetLocalXML = strLocalLandscapePATH
		
	End Property 
	
	Public Property Let SetSystemName(sys)
	
		strSSN = sys
		
	End Property 
	
	Public Property Let SetClientName(cli)
	
		strSCN = cli
	
	End Property 
	
	Public Property Get SessionFound
	
		SessionFound = boolSessionFound
		
	End Property 
	
		
	
End Class 