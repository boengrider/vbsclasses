Class RateUpload_v1

	Private oFSO
	Private oWSH
	Private oNET
	Private oSES ' Session should be obtained from SapLauncher.GetSession method
	Private strUserName ' System user name e.g. a293793
	Private strComputerName ' System name e.g. SKSENEW128
	
	' ============== Constructor & Destructor ==================
	Private Sub Class_Initialize
	
		Set oFSO = CreateObject("scripting.filesystemobject")
		Set oWSH = CreateObject("wscript.shell")
		Set oNET = CreateObject("wscript.network")
		oSES = Null 
		strUserName = oNET.UserName
		strComputerName = oNET.ComputerName

	End Sub
	
	Private Sub Class_Terminate
	
	End Sub
	
	' ============ P U B L I C  &  P R I V A T E   M E T H O D S  &   S U B R O U T I N E S =========



	' --------- UploadRates
	Public Function UploadRates(strFiles,strExRateType,boolDoNotNEX) ' strFiles is comma delimited list of files to upload,ex rate type ie YHR2, preserve session. Do not cal /NEX
	
		Dim validfrom,SAPfile,i,ratetype,filename
		i = 0
		ratetype = UCase(strExRateType)
	
		For Each SAPfile In Split(strFiles,",")
		
			If oFSO.FileExists(SAPfile) Then 
				filename = oFSO.GetFileName(SAPfile) ' Returns 20200630.txt 
				validfrom = "" ' Clear
				validfrom = Mid(filename,7,2) & "." & Mid(filename,5,2) & "." & Mid(filename,1,4) ' SAP compatible date format DD.MM.YYYY
				KillPopups(oSES)
				oSES.findById("wnd[0]/tbar[0]/okcd").text = "/NZTC_ZCURR_UPLOAD"
				oSES.findById("wnd[0]").sendVKey 0 ' ENTER
				KillPopups(oSES)
				oSES.findById("wnd[0]/usr/txtP_FILE").text = SAPfile
				oSES.findById("wnd[0]/usr/txtP_KURST").text = ratetype
				oSES.findById("wnd[0]/usr/ctxtP_GDATU").text = validfrom
				oSES.findById("wnd[0]").sendVKey 8
				KillPopups(oSES)
				oSES.findById("wnd[0]").sendVKey 0
				KillPopups(oSES)
		
				Do While oSES.Children.Count > 1
					oSES.findById("wnd[0]").sendVKey 0
				Loop
				i = i + 1
				WScript.Sleep 2000 ' Wait a bit
			End If 	
		Next
		
		If Not boolDoNotNEX Then
			oSES.findById("wnd[0]/tbar[0]/okcd").text = "/NEX" ' Close transaction window
			oSES.findById("wnd[0]").sendVKey 0
		End If 
	
		UploadRates = i ' Return the number of uploaded files or 0 if error occured   

	End Function 
	

				
	''============================================================
'' Program:   SUB Killpopups
'' Desc:      Kill of SAP popup screens which could appear when executing SAP transactions
'' Called by: 
'' Call:      KillPopups(connection.children(0)
'' Arguments: s = connection.children(0)
'' Changes---------------------------------------------------
'' Date		Programmer	Change
'' 2020-06-01	Tomas Chudik(tomas.chudik@volvo.com)	Written as vbscript SUB with arguments; supports kill of "System Message", "Copyright"
''============================================================

	Sub KillPopups(s)
		Do While s.Children.Count > 1
			If InStr(s.ActiveWindow.Text, "System Message") > 0 Then
				s.ActiveWindow.sendVKey 12
		
			ElseIf InStr(s.ActiveWindow.Text, "Copyright") > 0 Then
				s.ActiveWindow.sendVKey 0
				'ElseIF   'Insert next type of popup windows which you want to kill
			Else
				Exit Do
			End If
		Loop
	End Sub

	' ================= P R O P E R T I E S ====================
	Public Property Let SAPSession(s)
		Set oSES = s
	End Property 	
		

	
End Class 