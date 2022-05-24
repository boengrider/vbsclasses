'*********************** ECBRate v1 *************************************
' Added public read property --> Public Property Get FoundDateInXML
' Property returns true if the first non tcd date is found wihtin
' xml90
' This way we can loop over previous dates until we find the rate for the required date
' Example: 
' Dim oTCD,oECBR,oDF
' Set oDF = new DateFormatter
' Set oTCD = new TCDCalendar
' Set oECBR = new ECBRate
' oTCD.FindNonTCDDate($OUR_TARGET_DATE)	' Find the first non TCD date
' Do While oECBR.FoundDateInXML = False 
'	oECBR.MakeOutputFile True,oDF.ToYearMonthDayWithDashes(oTCD.FirstNonTCD)
'	oTCD.FindNonTCDDate ($OUR_TARGET_DATE - 1)
' Loop 
	
		
'-------------------- ECBRate Class ---------------------------------------
Class ECBRate 

	Private strXmlUrl
	Private strNameSpace
	Private dictQuantity        
	Private dictFcurrs					' Dictionary hodling target currencies
	Private dictOutputFiles				' Dictionary holding processed output file paths i.e C:\ExRate\SK01\20200101.txt C:\ExRate\SK01\20200102.txt ...
	Private strOutputDir
	Private strOutputFile
	Private strTcurr
	Private boolFoundDateInXml
	Private oOutFile
	Private oWSH
	Private oXML
	Private oHTTP
	Private oFSO
	Private errno

	
	' Constructor and destructor
	Private Sub Class_Initialize
	
		Set dictFcurrs = CreateObject("Scripting.Dictionary")
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		Set oHTTP = CreateObject("MSXML2.XMLHTTP")
		Set oXML = CreateObject("MSXML2.DOMDocument")
		Set dictQuantity = CreateObject("Scripting.Dictionary")
		Set oWSH = CreateObject("Wscript.Shell")
		
		boolFoundDateInXml = False
		strTcurr = "EUR"
		strNameSpace = "xmlns:gesmes='http://www.gesmes.org/xml/2002-08-01' xmlns='http://www.ecb.int/vocabulary/2002-08-01/eurofxref'"
		strXmlUrl = ""
		strOutputDir = ""
		
	End Sub 
	
	Private Sub Class_Terminate
	
	End Sub 
	
	' Public methods
	
	
	' Init()
	Public Function Init(strUrl,strOutDir)
	
		strOutputDir = strOutDir
		strXmlUrl = strUrl
		
	End Function
	
	' OverrideQuantity()
	Public Function OverrideQuantity(strCurrs)
	
		Dim curr
		For Each curr In Split(strCurrs,",")
		
			dictQuantity(curr) = "100"
			
		Next
		
	End Function 
	
	' MakeOutputFile()
	Public Function MakeOutputFile(boolClearIECache,strDate)
	
		Dim ChildNodes,ChildNode,Attributes,Attribute,i,key,delim
	
		If boolClearIECache Then
		
			ClearIE					' Clear internet explorer cache first
			
		End If 
		
		
		' Open http connection
		
		oHTTP.open "GET",strXmlUrl,False
		oHTTP.send
		
		If oHTTP.status <> 200 Then
			
			errno = 2        		' Set errno
			MakeOutputFile = -1 	' ERROR downloading rate file
			Exit Function
			
		End If 
		
		' Http request OK, continue loading xml
		oXML.load oHTTP.responseXML
		oXML.setProperty "SelectionNamespaces", strNameSpace
		Set ChildNodes = oXML.getElementsByTagName("Cube")
		
		i = 0
		
		Do While ChildNodes.item(i).attributes.length <> 1
		
			i = i + 1
			
		Loop
		
		
		On Error Resume next
		Do While IsObject(ChildNodes.item(i))
		
			debug.WriteLine ChildNodes.item(i).attributes.getNamedItem("time").text
			
			If strDate = ChildNodes.item(i).attributes.getNamedItem("time").text Then
				boolFoundDateInXml = True
				debug.WriteLine "Date found in the XML"
				Exit Do 				' Exit loop
			End If 
			
			i = i + (ChildNodes.item(i).childNodes.length) + 1
			
			If Not ChildNodes.item(i).hasChildNodes Then
			
				boolFoundDateInXml = False
				Exit Do 
				
			End If 
			
		Loop
		
		If Not boolFoundDateInXml Then
		
			errno = 3
			MakeOutputFile = -1
			Exit Function 
			
		End If 
		On Error GoTo 0 
		
		' Found the target date 
		' Select rates
		Set ChildNode = ChildNodes.item(i)
		Set ChildNodes = ChildNode.childNodes
		
		For Each ChildNode In ChildNodes 
		
			If dictFcurrs.Exists(ChildNode.attributes.getNamedItem("currency").text) Then
			
				dictFcurrs(ChildNode.attributes.getNamedItem("currency").text) = ChildNode.attributes.getNamedItem("rate").text
				
			End If 
			
		Next 
	
		' Dictionary hold pairs CURRENCY RATE
	
		If InStr(1 / 2,",") >= 1 Then
			delim = ","
		Else 
			delim = "."
		End If 
		
		For Each key In dictFcurrs.Keys
	
			dictFcurrs(key) = Replace(dictFcurrs.Item(key),".",delim) ' replace . with whatever is the system delimiter
			
		Next
		 
		Set oOutFile = oFSO.OpenTextFile(strOutputDir & "\" & strOutputFile,2,True)
		
		For Each key In dictFcurrs.Keys
		
			oOutFile.WriteLine key & vbTab & strTcurr & vbTab & FormatNumber(Round(1/dictFcurrs.Item(key),5) * dictQuantity.Item(key),5) & vbTab & dictQuantity.Item(key) & vbTab & "1"
		Next 
		
		MakeOutputFile = 0
		'CopyOutputFile
		
		
	End Function
	
	' CopytOutputFile()
	Private Sub CopyOutputFile
	
		oFSO.CopyFile strOutputDir & "\" & strOutputFile, "\\vcn.ds.volvo.net\cli-sd\sd1294\046629\output\01_SK01_ExRateProcessing\SK01\" & strOutputFile
	
	End Sub 
	
	Private Sub ClearIE
	
		oWSH.Run "RunDLL32.exe InetCpl.cpl,ClearMyTracksByProcess 8",0,True
		
	End Sub 
		
	
	' Properties
	Public Property Get FoundDateInXML
	
		FoundDateInXML = boolFoundDateInXml
		
	End Property 
	
	Public Property Let SetXmlURL(url)
	
		strXmlUrl = url
		
	End Property  
	
	Public Property Let SetOutputDirectory(dir)
	
		strOutputDir = dir
		
	End Property 
	
	Public Property Let SetOutputFile(file)
	
		strOutputFile = file
		
	End Property 
	
	Public Property Get GetOutputFile
	
		GetOutputFile = strOutputFile
		
	End Property 
	
	Public Property Get GetOutputDirectory
	
		GetOutputDirectory = strOutputDir
		
	End Property  
	
	Public Property Get GetXmlURL
	
		GetXmlURL = strXmlUrl
		
	End Property 
	
	Public Property Get GetErrorCode
	
		GetErrorCode = errno
		
	End Property 
	
	Public Property Let AddTargetCurrency(currs)
	
		Dim curr
		For Each curr In Split(currs,",") 
		
			dictFcurrs.Add curr,""
			dictQuantity.Add curr,"1"
			
		Next
		
	End Property
	
	
End Class 