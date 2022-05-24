Class CNBRate
	Private oWSH
	Private strPathToOutputDir
	Private strDate
	Private oHTTP
	Private dictRate
	Private dictQuantity
	Private oOutFile
	Private oTempFile
	Private oFSO
	Private strTCurr
	Private strPathToTempFile
	Private strRateURL
	Private strFCurrList
	Private strErrorSource
	Private boolMakeOutputFile
	Private errno
	
	
	' ---------- Class Constructor ------------
	Sub Class_Initialize
		Set oWSH = CreateObject("Wscript.Shell")
		Set oHTTP = CreateObject("MSXML2.XMLHTTP")
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		Set dictRate = CreateObject("Scripting.Dictionary")
		Set dictQuantity = CreateObject("Scripting.Dictionary")
		strRateURL = "https://www.cnb.cz/cs/financni-trhy/devizovy-trh/kurzy-devizoveho-trhu/kurzy-devizoveho-trhu/denni_kurz.txt"
		strTCurr = "CZK"
		boolMakeOutputFile = False
		strErrorSource = Null
		strPathToOutputDir = Null
		strFCurrList = Null
		strDate = Null
	End Sub 
	
	' ---------- Class Destructor -------------
	Sub Class_Terminate
		
	End Sub 
	
	' -------------------------------------------
	' --------- P u b l i c   M e t h o d s -----
	' -------------------------------------------
	
	' CreateTEMPFile
	' Method creates a unique temp file. 
	' If successfull it returns 0
	' On error it returns -1 and errno is set appropriately
	Public Function CreateTEMPFile
		If IsEmpty(strPathToTempFile) Or strPathToTempFile = "" Then
			boolMakeOutputFile = False 
			errno = 1 ' Can't create a temp file. No path to temp file was provided
			CreateTEMPFile = -1 ' No path provided. File cannot be created
			Exit Function
		End If
		
		Set oTempFile = oFSO.OpenTextFile(strPathToTempFile,2,True)
		oHTTP.open "GET",strRateURL & "?date=" & strDate,False 
		oHTTP.send
		
		
		' In case URL cannot be accessed, close the temp file and delete it. Return HTTP error code
		If oHTTP.status <> 200 Then
			oTempFile.Close
			oFSO.DeleteFile strPathToTempFile
			boolMakeOutputFile = False
			errno = 2 ' Rate file can't be downloaded from the CNB web
			CreateTEMPFile = -1 ' ERROR downloading rate file
			Exit Function
		End If 
		' Write content to the temp file
		oTempFile.Write oHTTP.responseText
		oTempFile.Close
		boolMakeOutputFile = True
		CreateTEMPFile = 0 ' File created successfully
	End Function 
	
	Public Function DeleteTEMPFile
		If Not IsNull(oTempFile) And oFSO.FileExists(strPathToTempFile) Then
			oTempFile.Close
			oFSO.DeleteFile strPathToTempFile
		End If 
	End Function
	
	' Init()
	' strPath -> Absolute path to the temp file
	' strUrl  -> Rate URL or Null
	' strFCurrs -> comma delimited list of the wanted currencies or Null
	' strTargetDate -> target date or Null
	' strPathToOD -> Absolute path to the directory where to put otput files locally e.g C:\ExRate\CZ02
	Public Function Init(strPath,strUrl,strFCurrs,strTargetDate,strPathToOD) 		
		strPathToTempFile = strPath
		If Not IsNull(strPathToOD) Then
			strPathToOutputDir = strPathToOD
		Else ' Try to recover and use C:\ExRate path
				strPathToOutputDir = oWSH.ExpandEnvironmentStrings("%SYSTEMDRIVE%") & "\ExRate"
		End If
	
		Call MakeOutputDir ' Create utput directory
		
		
		If Not IsNull(strTargetDate) Then
			strDate = strTargetDate
		End If 
		
		If strUrl = "" Or IsNull(strUrl) Then
			errno = 1 ' No URL provided to the Init function
			Init = -1
			Exit Function
		Else 
			strRateURL = strUrl
		End If 
		
		If Not strFCurrs = "" Or Not IsNull(strFCurrs) Then
			strFCurrList = strFCurrs
		Else
			' Nothing
		End If 
			
		
	End Function 
	
	' ----------- ParseRateFile() ------------------------
	Public Function ParseRateFile() ' Call this method to add an entry into the dictionary. This function/method will do the parsing. It needs txt file with rates
		Dim key,column,line
	
		
		If IsNull(strDate) Then
			boolMakeOutputFile = False
			errno = 1 ' No Date was provided to the ParseRateFile() function
			ParseRateFile = -1 
			Exit Function
		End If 
   		
   		Set oTempFile = oFSO.OpenTextFile(strPathToTempFile,1,False)
   		line = oTempFile.ReadLine ' 1st line should contain date in DD.MM.YYYY  Compare it to our strDate
   		If Not strDate = Mid(line,1,10) Then
   			oTempFile.Close
   			If oFSO.FileExists(strPathToTempFile) Then
   				oFSO.DeleteFile strPathToTempFile
   			End If 
   			boolMakeOutputFile = False
   			errno = 2 ' Bad date. Target date and downloaded file date dont match
   			ParseRateFile = -1
   			Exit Function 
   		End If 
   		' ------ All OK continue parsing
   		Do While Not oTempFile.AtEndOfStream
   			
   			line = oTempFile.ReadLine ' 1st line should contain date in DD.MM.YYYY  Compare it to our strDate
   			column = Split(line,"|")
   			
   			For Each key In Split(strFCurrList,",")
   				
   				If key = column(3) Then 
  					
   					dictRate.Add column(3),column(4) ' RATE VALUE e.g. AUD 16,421
   						
   				End If 
   					
   			Next
   			
   		Loop
   		boolMakeOutputFile = True
   	End Function
   	
   	
   	Function MakeOutputFile
   		Dim key
   		' Check if the OD exists
   		If Not boolMakeOutputFile Then
   			errno = 1 ' Cant continue making output file
   			MakeOutputFile = -1
   			Exit Function
   		End If 
   		
   		If Not IsNull(strPathToOutputDir) Then
   			
   			Set oOutFile = oFSO.OpenTextFile(strPathToOutputDir & "\" & Right("0000" & Year(Date() + 1),4) & Right("00" & Month(Date() + 1),2) & Right("00" & Day(Date() + 1),2) & ".txt",2,True)
   			For Each key In Split(strFCurrList,",")
   				
   				If dictQuantity.Exists(key) Then
   					
   					oOutFile.WriteLine strTCurr & vbTab & key & vbTab & FormatNumber(Round((1 / dictRate(key)),5),5) & vbTab & "1" & vbTab & "100"
   					oOutFile.WriteLine key & vbTab & strTCurr & vbTab & FormatNumber(Round(dictRate(key),5),5) & vbTab & "100" & vbTab & "1"
   					
   				Else 
   				
   					oOutFile.WriteLine strTCurr & vbTab & key & vbTab & FormatNumber(Round((1 / dictRate(key)),5),5) & vbTab & "1" & vbTab & "1"
   					oOutFile.WriteLine key & vbTab & strTCurr & vbTab & FormatNumber(Round(dictRate(key),5),5) & vbTab & "1" & vbTab & "1"
   					
   				End If 
   					
   			Next
   			
   			oOutFile.WriteLine "EUR" & vbTab & "SEK" & vbTab & FormatNumber(Round((dictRate("EUR") / dictRate("SEK")),5),5) & vbTab & "1" & vbTab & "1"
   			oOutFile.WriteLine "SEK" & vbTab & "EUR" & vbTab & FormatNumber(Round((1 / ((dictRate("EUR") / dictRate("SEK")))),5),5) & vbTab &  "1" & vbTab & "1"
   			oOutFile.Close
   			MakeOutputFile = 0
   		End If 
   			
   		
   		
   	End Function 
   	
   	
   	Private Function MakeOutputDir
	Dim comps,i,l,path 
	comps = Split(strPathToOutputDir,"\")
	l = UBound(comps) ' save len
	i = 0
	
	Do While Not i = l + 1
		
		If oFSO.GetDriveName(comps(i)) = comps(i) Then
			path = comps(i)
			i = i + 1
		End If
		
		path = path & "\" & comps(i)
		If Not oFSO.FolderExists(path) Then
			oFSO.CreateFolder path
		End If
		
		i = i + 1
		
	Loop 
	
End Function

   
   			
   	
   	' ----------------------------------------------
	' -------------- P r o p e r t i e s -----------
	' ----------------------------------------------
	Public Property Get RateURL
		RateURL = strRateURL
	End Property 
	
	Public Property Get PathToTempFile
		PathToTempFile = strPathToTempFile
	End Property
	
	Public Property Get Tcurr 
		Tcurr = strTCurr
	End Property 
	
	Public Property Let Tcurr(strCurr)
		strTCurr = strCurr
	End Property 
	
	Public Property Get Fcurrs
		Fcurrs = strFCurrList
	End Property 
	
	Public Property Let Fcurrs(strCurrs)
		strFCurrList = strCurrs
	End Property 
	
	Public Property Get Ddate 
		Ddate = strDate
	End Property 
	
	Public Property Get GetRate(strCurrency)
		GetRate = dictRate.Item(strCurrency)
	End Property 
	
	Public Property Let OverrideQuantity(strQuantity)
		Dim key
		For Each key In Split(strQuantity,",")
			dictQuantity.Add key,100
		Next
	End Property 
	
	Public Property Get ErrorCode
		ErrorCode = errno
	End Property 
	
	
End Class 
