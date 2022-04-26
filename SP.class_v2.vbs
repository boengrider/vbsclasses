Option Explicit

Dim retval
Dim oSP : Set oSP = New SharePoint
retval = oSP.SharePoint("https://volvogroup.sharepoint.com/sites/unit-rc-sk-bs-it","40784cc3-ba68-45d0-9891-f3dfa8f04d15","cDES2gLLi%2BBRI/FcUizAyZuGQFQ5p%2B6rrknc3kMBWmE=",True)
oSP.UpdateSingleListItemJ "WDAPP","$select=Title&$filter=(Title eq 'SK01_IBANcheckVAT')","{""ComputerName"":""Test""}"



Class SharePoint
	
	Private oRX 
	Private oXML
	Private strAuthUrlPart1
	Private strAuthUrlPart2
	Private vti_bin_clientURL
	Private oHTTP
	Private oFSO
	Private strClientID
	Private strSecurityToken
	Private strClientSecret
	Private strFormDigestValue
	Private strTenantID
	Private strResourceID
	Private strURLbody
	Private strSiteURL
	Private strDomain
	Private numHTTPstatus
	Private strSite
	Private errDescription
	Private errNumber
	Private errSource
	Private boolRaise

	
		Private Sub Class_Initialize
		errDescription = ""
		errNumber = 0
		numHTTPstatus = 0
		strAuthUrlPart1 = "https://accounts.accesscontrol.windows.net/"
		strAuthUrlPart2 = "/tokens/OAuth/2"
		boolRaise = False 
		Set oRX = New RegExp
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		Set oXML = CreateObject("MSXML2.DOMDocument")
	End Sub 
	

	Public Function SharePoint(sSiteUrl,sClientID,sClientSecret,bRaise)
		Dim strErrSource : strErrSource = "SharepointLite.SharePointLite()"
		Dim tmp,retval
		boolRaise = bRaise
		oRX.Global = True
		oRX.Multiline = True
		oRX.IgnoreCase = True
		oRX.Pattern = "(http:\/\/|https:\/\/)([^\/])*\/sites\/([^\/])*\/{0,1}"
		
		If oRX.Test(sSiteUrl) Then
			tmp = oRX.Execute(sSiteUrl)(0)
			If Right(tmp,1) <> "/" Then
				strSiteURL = tmp & "/"
			Else
				strSiteURL = tmp
			End If 
		ElseIf boolRaise Then
			err.Raise 100, strErrSource, "Bad URL -> " & sSiteUrl
		Else
			errSource = strErrSource
			errNumber = 100
			errDescription = "Bad URL -> " & sSiteUrl
			SharePoint = 100
			Exit Function
		End If 
		
		
		vti_bin_clientURL = strSiteURL & "_vti_bin/client.svc"
		tmp = Split(strSiteURL,"/")
		strDomain = tmp(2)
		strClientID = sClientID
		strClientSecret = sClientSecret
		
		retval = GetTenantID      ' Obtain the Tenant/Realm ID
		If retval <> 0 Then
			SharePoint = retval
			Exit Function 
		End If 
		
		retval = GetSecurityToken ' Obtain the Security Token
		If retval <> 0 Then
			SharePoint = retval
			Exit Function
		End If 
		
		retval = GetXDigestValue  ' Obtain the form digest value
		If retval <> 0 Then
			SharePoint = retval
			Exit Function
		End If 
		
		
		SharePoint = 0
	
	End Function 
		
	'********************** P R I V A T E   F U N C T I O N S ************************
	
	
	'##############################
	'######### GetTenantID ########
	'##############################
	Private Function GetTenantID()
		Dim rxResult
		Dim strErrSource : strErrSource = "Sharepoint.GetTenantID()"
		
		With oHTTP
			.open "GET",vti_bin_clientURL,False
			.setRequestHeader "Authorization","Bearer"
			.send
		End With
		
		If Not oHTTP.status = 401 And boolRaise Then
			err.Raise oHTTP.status, strErrSource, oHTTP.responseText
		ElseIf Not oHTTP.status = 401 And Not boolRaise Then 
			GetTenantID = oHTTP.status
			Exit Function
		End If 
		
		oRX.Pattern = "Bearer realm=""([a-zA-Z0-9]{1,}-)*[a-zA-Z0-9]{12}"
		If oRX.Test(oHTTP.getResponseHeader("WWW-Authenticate")) Then
			Set rxResult = oRX.Execute(oHTTP.getResponseHeader("WWW-Authenticate"))
			oRX.Pattern = "[a-fA-F0-9]{8}-([a-fA-F0-9]{4}-){3}[a-fA-F0-9]{12}"
			If oRX.Test(rxResult(0)) Then 
				strTenantID = oRX.Execute(rxResult(0))(0)
			ElseIf boolRaise Then
				err.Raise 1000, strErrSource, "Bearer realm not found"
			End If 
		ElseIf boolRaise Then
			err.Raise 1000, strErrSource, "Bearer realm not found"
		Else
			errSource = strErrSource
			errNumber = 1000
			errDescription = "Bearer realm not found"
			GetTenantID = 1000
			Exit Function
		End If 
		
		oRX.Pattern = "client_id=""[a-fA-F0-9]{8}-([a-fA-F0-9]{4}-){3}[a-fA-F0-9]{12}"
		If oRX.Test(oHTTP.getResponseHeader("WWW-Authenticate")) Then
			Set rxResult = oRX.Execute(oHTTP.getResponseHeader("WWW-Authenticate"))
			oRX.Pattern = "[a-fA-F0-9]{8}-([a-fA-F0-9]{4}-){3}[a-fA-F0-9]{12}"
			If oRX.Test(rxResult(0)) Then 
				strResourceID = oRX.Execute(rxResult(0))(0)
			ElseIf boolRaise Then
				err.Raise 1000, strErrSource, "Client_id not found"
			Else
				GetTenantID = 1000
				Exit Function 
			End If  
		ElseIf boolRaise Then
			err.Raise 1000, strErrSource, "Client_id not found"
		Else
			errSource = strErrSource
			errNumber = 1000
			errDescription = "Client_id not found"
			GetTenantID = 1000
			Exit Function 
		End If 
		
		GetTenantID = 0
	End Function
	
	
	'##############################
	'####### GetXDigestValue ######
	'##############################
	Private Function GetXDigestValue()
		Dim strErrSource : strErrSource = "Sharepoint.GetXDigestValue()"
		Dim colNodes
		
		With oHTTP
			oHTTP.open "POST", strSiteURL & "_api/contextinfo", False 
			oHTTP.setRequestHeader "accept","application/atom+xml;odata=verbose"
			oHTTP.setRequestHeader "authorization", "Bearer " & strSecurityToken
			oHTTP.send
		End With 
		
		If Not oHTTP.status = 200 And boolRaise Then
			err.Raise oHTTP.status, strErrSource, oHTTP.responseText
		ElseIf Not oHTTP.status = 200 Then
			errSource = strErrSource
			errNumber = oHTTP.status
			errDescription = oHTTP.responseText
			GetXDigestValue = oHTTP.status
			Exit Function 
		End If 
		
		oXML.loadXML oHTTP.responseText
		
		Set colNodes = oXML.selectNodes("//d:FormDigestValue")
		
		If colNodes.length = 0 And boolRaise Then
			err.Raise 1100, strErrSource, "FormDigestValue not found"
		ElseIf colNodes.length = 0 Then
			errSource = strErrSource
			errNumber = 1100
			errDescription = "FormDigestValue not found"
			GetXDigestValue = 1100
			Exit Function
		Else 
			strFormDigestValue = colNodes.item(0).text
		End If 
		
		GetXDigestValue = 0	
	End Function
	
	
	'##############################
	'###### GetSecurityToken ######
	'##############################
	Private Function GetSecurityToken()
		Dim rxResult
		Dim strErrSource : strErrSource = "Sharepoint.GetSecurityToken()"
		Dim strURLbody : strURLbody = "grant_type=client_credentials&client_id=" & strClientID & "@" & strTenantID & "&client_secret=" & strClientSecret & "&resource=" & strResourceID & "/" & strDomain & "@" & strTenantID
		
		With oHTTP
			.open "POST", strAuthUrlPart1 & strTenantID & strAuthUrlPart2, False
			.setRequestHeader "Host","accounts.accesscontrol.windows.net"
			.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
			.setRequestHeader "Content-Length", CStr(Len(strURLbody))
			.send strURLbody
		End With 
		
		If Not oHTTP.status = 200 And boolRaise Then
			err.Raise oHTTP.status, strErrSource, oHTTP.responseText
		ElseIf Not oHTTP.status = 200 Then
			errSource = strErrSource
			errNumber = oHTTP.status
			errDescription = oHTTP.responseText
			GetSecurityToken = oHTTP.status
			Exit Function 
		End If 

		oRX.Pattern = "access_token"":"".*"
		If oRX.Test(oHTTP.responseText) Then
			Set rxResult = oRX.Execute(oHTTP.responseText)
			rxResult = Split(rxResult(0),":")
			rxResult(1) = Replace(rxResult(1),"""","")
			rxResult(1) = Replace(rxResult(1),"}","")
			strSecurityToken = rxResult(1) ' Save the token 
		ElseIf boolRaise Then
			 err.Raise 1200, strErrSource, "Access token not found"
		Else
			errSource = strErrSource
			errNumber = 1200
			errDescription = "Access token not found"
			GetSecurityToken = 1200
			Exit Function
		End If 
		
		GetSecurityToken = 0 	
	End Function 
	
	
	
	
	Private Function Strip(sString)
		If Right(sString,1) = "/" Then
			sString = Mid(sString,1,Len(sString) - 1)
		End If 
		If Left(sString,1) = "/" Then
			sString = Mid(sString,2,Len(sString) - 1)
		End If
		
		Strip = sString
	End Function 
	
	
	'********************** P U B L I C   F U N C T I O N S ************************
	'Error codes 2xxx
	
	'sQuery example: "$select=Title&$filter=(Title eq 'SK01_IBANcheckVAT')"
	
	'Function updates only the first found result, make your query
	Public Function UpdateSingleListItemJ(sListName,sQuery,sJsonPatch)
		Dim strErrSource : strErrSource = "Sharepoint.UpdateSingleListItemJ"
		
		With oHTTP
			.open "GET", strSiteURL & "_api/web/lists/getbytitle('" & sListName & "')/items?" & sQuery      
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.setRequestHeader "X-RequestDigest", strFormDigestValue
			.send
		End With 
		
		If Not oHTTP.status = 200 And boolRaise Then
			err.Raise oHTTP.status, strErrorSource, oHTTP.responseText
		ElseIf Not oHTTP.status = 200 Then
			UpdateSingleListItemJ = oHTTP.status
			Exit Function
		End If 
		
		debug.WriteLine oHTTP.responseText
		oXML.loadXML oHTTP.responseText

	End Function 




	Public Function GetListItem(sListName,sFieldName,sFieldValue)
		Dim oHTTP
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		With oHTTP
			.open "GET", strSiteURL & "_api/web/lists/GetByTitle('" & sListName & "')/items?$select=" & sFieldName & "&$filter=" & sFieldName & " eq " & "'" & sFieldValue & "'", False 
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With
		
		If Not oHTTP.status = 200 Then
			GetListItem = False ' Something went wrong. Assume the item doesn't exist an owervrite it. Or lose it !
			Exit Function
		End If 
		
		oXML.loadXML oHTTP.responseText
		
		If oXML.getElementsByTagName("d:Title").length > 0 Then
			If sFieldValue = oXML.getElementsByTagName("d:Title").item(0).text Then
				GetListItem = True
				Exit Function
			Else
				GetListItem = False
				Exit Function
			End If
		Else
			GetListItem = False
			Exit Function
		End If 
	End Function
	
	Public Function UpdateList()
	
		Dim oHTTP,body,oSTREAM
		body = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">" _
		& "<soap12:Body><UpdateListItems xmlns=""http://schemas.microsoft.com/sharepoint/soap/"">" _
		& "<listName>SK01_Manual_Payments_QA</listName><updates><Field Name=""ID"">21<Field><Field Name=""Title"">HELLO</Field></updates></UpdateListItems></soap12:Body></soap12:Envelope>"
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
		oHTTP.open "POST","https://volvogroup.sharepoint.com/sites/unit-rc-sk-bs-it/_vti_bin/Lists.asmx",False
		oHTTP.setRequestHeader "Host","volvogroup.sharepoint.com"
		oHTTP.setRequestHeader "Content-Type","application/soap+xml; charset=utf-8"
		oHTTP.setRequestHeader "Content-Length",Len(body)
		oHTTP.send body 
	End Function 
	
	Public Function GetFileInfo(sServerRelFilePath)
		
		If Not Left(sServerRelFilePath,1) = "/" Then
			sServerRelFilePath = "/" & sServerRelFilePath
		End If 
		 
		With oHTTP
			.open "GET", strSiteURL & "_api/web/getFileByServerRelativeUrl('/sites/" & strSite & sServerRelFilePath & "')/Properties"
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With 
		
		If Not oHTTP.status = 200 Then
			GetFileInfo = -1
			Exit Function
		End If 
		
		oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
		oXML.loadXML oHTTP.responseText
		debug.WriteLine oXML.getElementsByTagName("d:vti_x005f_filesize").item(0).text
		
	End Function 
	
			
			
	Public Function DownloadFile(sServerRelFilePath,sSaveAsPath)
	
		Dim oHTTP,oSTREAM
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP")
		With oHTTP
			.open "GET", strSiteURL & "_layouts/15/download.aspx?SourceUrl=https://" & strDomain & sServerRelFilePath
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.send
		End With
		
		debug.WriteLine oHTTP.responseText
		debug.WriteLine oHTTP.status
		
		If oHTTP.status = 200 Then
			Set oSTREAM = CreateObject("ADODB.Stream")
			oSTREAM.Open
			oSTREAM.Type = 1
			oSTREAM.Write oHTTP.responseBody
			oSTREAM.SaveToFile sSaveAsPath
			oSTREAM.Close
		Else
			debug.WriteLine oHTTP.status
			debug.WriteLine oHTTP.responseText
		End If 
		
	End Function 
	
			
	Public Function GetFileCount(sServerRelDirPath)
	
		Dim oHTTP,oXML,colElements
		sServerRelDirPath = Strip(sServerRelDirPath)
		Set oHTTP = CreateObject("MSXML2.XMLHTTP")
		Set oXML = CreateObject("MSXML2.DOMDocument")
		
		With oHTTP
			.open "GET", oSP.GetSiteURL & "_api/web/GetFolderByServerRelativeUrl('/sites/" & strSite & sServerRelDirPath & "')/Files", False
			.setRequestHeader "Authorization", "Bearer " & oSP.GetToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With
		
		If Not oHTTP.status = 200 Then 
			GetFileCount = -1
			Exit Function
		End If 

		
		oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
		oXML.loadXML oHTTP.responseText
		Set colElements = oXML.getElementsByTagName("d:Name")
		
		GetFileCount = colElements.length
				
	End Function 
	
	Public Function FolderExists(sRelDirPath)
		
		Dim oHTTP,oXML,colElements
		Set oHTTP = CreateObject("MSXML2.XMLHTTP")
		Set oXML = CreateObject("MSXML2.DOMDocument")
		oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
		With oHTTP
			.open "GET", strSiteURL & "_api/web/GetFolderByServerRelativeUrl('" & sRelDirPath & "')", False
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With
		
		If oHTTP.status = 200 Then
			oXML.loadXML oHTTP.responseText
			Set colElements = oXML.getElementsByTagName("d:Exists")
			If colElements.length > 0 Then
				If LCase(colElements.item(0).text) = "true" Then
					FolderExists = True
					Exit Function
				Else
					FolderExists = False
					Exit Function
				End If 
			End If  
		Else
			FolderExists = False
		End If 
	
	End Function 
	
	Public Function DownloadFilesA(sServerRelDirPath,sDestinationFolder,ByRef arrFiles)
		
		Dim item,nodes
		Dim path
		Dim oSTREAM : Set oSTREAM = CreateObject("ADODB.Stream")
		
		For item = 0 To UBound(arrFiles)
			
			With oHTTP
				'.open "GET", strSiteURL & "_layouts/15/download.aspx?SourceUrl=" & strSiteURL & sServerRelDirPath & "/" & arrFiles(item), False
				.open "GET", strSiteURL & "_api/web/GetFolderByServerRelativeUrl('" & sServerRelDirPath & "')/Files('" & arrFiles(item) & "')", False
				.setRequestHeader "Authorization", "Bearer " & strSecurityToken
				.send
			End With
			
			
			If Not oHTTP.status = 200 Then
				errSource = "MSXML2.XMLHTTP"
				errDescription = oHTTP.responseText
				errNumber = oHTTP.status
				DownloadFilesA = -1
				Exit Function
			End If 
			
			oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
			oXML.loadXML oHTTP.responseText
			
			Set nodes = oXML.getElementsByTagName("d:ServerRelativeUrl")
			
			If Not nodes.length > 0 Then
				errSource = "MSXML2.XMLHTTP"
				errDescription = "d:serverRelativeUrl node missing. Affected file: " & arrFiles(item)
				errNumber = Hex(1000)
				DownloadFilesA = -1
				Exit Function
			End If 
			
			path = nodes.nextNode.text ' Save relative URl
			debug.WriteLine strSiteURL & "_layouts/15/download.aspx?SourceUrl=https://" & strDomain & path
			With oHTTP
				.open "GET", strSiteURL & "_layouts/15/download.aspx?SourceUrl=https://" & strDomain & path, False
				.setRequestHeader "Authorization", "Bearer " & strSecurityToken
				.send
			End With
			
			If Not oHTTP.status = 200 Then
				errSource = "MSXML2.XMLHTTP"
				errDescription = oHTTP.responseText
				errNumber = oHTTP.status
				DownloadFiles = -1
				Exit Function
			End If 
			
			
			oSTREAM.Open
			oSTREAM.Type = 1
			oSTREAM.Write oHTTP.responseBody
			On Error Resume Next 
			oSTREAM.SaveToFile oFSO.BuildPath(sDestinationFolder,oFSO.GetFileName(path))
			
			If err.number > 0 Then 
				errSource = err.Source
				errDescription = err.Description
				errNumber = err.number
				DownloadFiles = -1
				oSTREAM.Close
				Exit Function
			End If 
			
			oSTREAM.Close
			
		Next
		
	End Function 
	
	
	
	
	
	
	Public Function DownloadFiles(sServerRelDirPath,sDestinationFolder)
	
		Dim item,nodes
		Dim path
		Dim oSTREAM : Set oSTREAM = CreateObject("ADODB.Stream")
		
		With oHTTP
			.open "GET", strSiteURL & "_api/web/GetFolderByServerRelativeUrl('/sites/" & strSite & sServerRelDirPath & "')/Files", False
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.send
		End With 
		
		If Not oHTTP.status = 200 Then
			errDescription = oHTTP.responseText
			errNumber = oHTTP.status
			DownloadFiles = -1
			Exit Function
		End If 
		
		oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
		oXML.loadXML oHTTP.responseText
		
		Set nodes = oXML.getElementsByTagName("d:ServerRelativeUrl")
		
		For item = 0 To nodes.length - 1
			path = nodes.nextNode.text
			With oHTTP
				.open "GET", strSiteURL & "_layouts/15/download.aspx?SourceUrl=https://" & strDomain & path, False
				.setRequestHeader "Authorization", "Bearer " & strSecurityToken
				.send
			End With
			
			If Not oHTTP.status = 200 Then
				errDescription = oHTTP.responseText
				errNumber = oHTTP.status
				DownloadFiles = -1
				Exit Function
			End If 
			
			
			oSTREAM.Open
			oSTREAM.Type = 1
			oSTREAM.Write oHTTP.responseBody
			On Error Resume Next 
			oSTREAM.SaveToFile oFSO.BuildPath(sDestinationFolder,oFSO.GetFileName(path))
			
			If err.number > 0 Then 
				errSource = err.Source
				errDescription = err.Description
				errNumber = err.number
				DownloadFiles = -1
				Exit Function
			End If 
			
			oSTREAM.Close
			
		Next 
		
		DownloadFiles = nodes.length
		
	End Function 
	
	Public Function GetFilesA(sServerRelDirPath,ByRef dictFiles)
	
		Dim item,nodes
		With oHTTP
			.open "GET", strSiteURL & "_api/web/GetFolderByServerRelativeUrl('/sites/" & strSite & sServerRelDirPath & "')/Files", False
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With 
		
		If Not oHTTP.status = 200 Then
			GetFilesA = -1
			Exit Function
		End If 
		
		oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
		oXML.loadXML oHTTP.responseText
		
		Set nodes = oXML.getElementsByTagName("d:Name")
		
		For item = 0 To nodes.length - 1
			dictFiles.Add nodes.nextNode.text,""
		Next 
		
		GetFilesA = dictFiles.Count
	
	End Function 
		
	Public Function GetFiles(sServerRelDirPath,ByRef colFiles) ' sType "json" or "atom+xml"
	
		Dim dictFilesInSourceDir
		Dim dictFiles
		Dim oHTTP,oXML,colItems,item,colPaths,path
		Set dictFiles = CreateObject("Scripting.Dictionary")
		Set oXML = CreateObject("MSXML2.DOMDocument")
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If Left(sServerRelDirPath,1) = "/" Then
			sServerRelDirPath = Right(sServerRelDirPath,Len(sServerRelDirPath) - 1)
		End If
		If Right(sServerRelDirPath,1) = "/" Then
			sServerRelDirPath = Left(sServerRelDirPath,Len(sServerRelDirPath) - 1)
		End If 
		
		With oHTTP
			.open "GET", strSiteURL & "_api/web/GetFolderByServerRelativeUrl('" & sServerRelDirPath & "')/Files", False
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With
	
		If oHTTP.status = 200 Then
			oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
			oXML.loadXML oHTTP.responseText
			Set colItems = oXML.getElementsByTagName("d:Name")
			Set colPaths = oXML.getElementsByTagName("d:ServerRelativeUrl")
		
			For item = 0 To colItems.length - 1
				colFiles.add colItems.item(item).text,colPaths.item(item).text
			Next
			
'			colFiles = dictFiles
'			Exit Function
		End If 
	
	End Function
			
			
	Public Function MoveFile2(sSourceRelDirPath,sDestRelDirPath)

		If Left(sSourceRelDirPath,1) = "/" Then
			sSourceRelDirPath = Right(sSourceRelDirPath,Len(sSourceRelDirPath) - 1)
		End If 
		If Left(sDestRelDirPath,1) = "/" Then
			sDestRelDirPath = "/" & Right(sDestRelDirPath,Len(sDestRelDirPath) - 1)
		End If 
		 
		Dim oHTTP,strBody
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		'strBody = "{ ""srcPath"": { ""__metadata"": ""SP.ResourcePath"" },""DecodeUrl"":" & strSiteURL & sSourceRelDirPath & """},""destPath"": { ""__metadata"": ""SP.ResourcePath"" },""DecodeUrl"":" & strSiteURL & sDestRelDirPath & """ } }"
		strBody = "{""srcPath"": {""__metadata"": {""type"": ""SP.ResourcePath""},""DecodedUrl"": """ & strSiteURL & sSourceRelDirPath & """},""destPath"": {""__metadata"": {""type"": ""SP.ResourcePath""},""DecodedUrl"": """ & strSiteURL & sDestRelDirPath & """}}"
		
		With oHTTP
			.open "POST","https://volvogroup.sharepoint.com/sites/unit-rc-sk-bs-it/_api/SP.MoveCopyUtil.MoveFileByPath(overwrite=@a1)?@a1=true"
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/json;odata=nometadata"
			.setRequestHeader "Content-Type", "application/json;odata=verbose"
			.setRequestHeader "Content-Length", Len(strBody)
			.send strBody
		End With
		
		If oHTTP.status = 200 Then
			MoveFile2 = 1 ' Return 1 or True if successfull
			Exit Function
		Else
			MoveFile2 = 0 ' Return 0 or False if failed
			Exit Function
		End If 
	
	End Function
	
	
	Public Function AddListItem(sListName,sJsonRequest)
'		To do this operation, you must know the ListItemEntityTypeFullName property of the list And
'		pass that as the value of type in the HTTP request body. Following is a sample rest call to get the ListItemEntityTypeFullName

		Dim oHTTP,oXML,strEntityTypeFullName,colElements,request
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		Set oXML = CreateObject("MSXML2.DOMDocument")
		oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
		
		With oHTTP
			.open "GET", strSiteURL & "_api/web/lists/GetByTitle('" & sListName & "')?$select=ListItemEntityTypeFullName", False
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With 
		
		If oHTTP.status = 200 Then
			oXML.loadXML oHTTP.responseText
			Set colElements = oXML.getElementsByTagName("d:ListItemEntityTypeFullName")
			If colElements.length >= 1 Then
				strEntityTypeFullName = colElements.item(0).text
			Else
				AddListItem = -1 ' Couldn't obtain the EntityTypeFullName
				Exit Function 
			End If
		Else 
			AddListItem = -2 ' http error
			Exit Function 
		End If
		
		
		sJsonRequest = "{""__metadata"": { ""type"": """ & strEntityTypeFullName & """ }," & sJsonRequest ' Prepend metadata part
		With oHTTP
			.open "POST", strSiteURL & "_api/web/lists/GetByTitle('" & sListName & "')/items", False
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/json;odata=verbose"
			.setRequestHeader "Content-Type", "application/json;odata=verbose"
			.setRequestHeader "If-None-Match", "*"
			.setRequestHeader "Content-Length", Len(sJsonRequest)
			.setRequestHeader "X-RequestDigest", strFormDigestValue
			.send sJsonRequest
		End With
		
		If oHTTP.status = 201 Then
			AddListItem = 0 ' Success
			Exit Function
		Else 
			AddListItem = oHTTP.status
			debug.WriteLine oHTTP.responseText
			Exit Function
		End If 
		
		
	End Function 
	
	Public Function DeleteAllItemsInList(sListName)
		
		Dim oHTTP,oXML,element
		Set oXML = CreateObject("MSXML2.DOMDocument")
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		
		With oHTTP
			.open "GET", strSiteURL & "_api/web/lists/GetByTitle('" & sListName & "')/items", False 
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accpet", "application/json;odata=verbose"
			.setRequestHeader "Content-Type", "application/json"
'			.setRequestHeader "If-Match", "{etag or *}"
'			.setRequestHeader "X-HTTP-Method", "DELETE"
			.send
		End With 
		
		oXML.loadXML oHTTP.responseText
		For Each element In oXML.getElementsByTagName("d:Id")
			With oHTTP
				.open "POST", strSiteURL & "_api/web/lists/GetByTitle('" & sListName & "')/items(" & element.text & ")", False
				.setRequestHeader "Authorization", "Bearer " & strSecurityToken
				.setRequestHeader "Accpet", "application/json;odata=verbose"
				.setRequestHeader "Content-Type", "application/json"
				.setRequestHeader "If-Match", "*"
				.setRequestHeader "X-HTTP-Method", "DELETE"
				.send
			End With 
		Next 
	End Function 
		
			
	
	
	
	
	' ************************** P R O P E R T I E S ******************************
	Public Property Get GetDigest
		
		GetDigest = strFormDigestValue
		
	End Property 
	
	Public Property Get GetToken
	
		GetToken = strSecurityToken
		
	End Property 
	
	Public Property Get GetHttpResponse
		GetHttpResponse = oHTTP.responseText
	End Property 
	
	Public Property Get GetHttpResponseHeaders(strHeader) ' If strHeader "*" then get all headers
		If strHeader = "*" Then
		
			GetHttpResponseHeaders = oHTTP.getAllResponseHeaders
			Exit Property
		
		End If
		
		GetHttpResponseHeaders = oHTTP.getResponseHeader(strHeader)
		
	End Property 
	
	Public Property Get GetRealmTenantID
		GetRealmTenantID = strTenantID
	End Property
	
	Public Property Get GetClientID
		GetClientID = strClientID
	End Property 
	
	Public Property Get GetResourceID
		GetResourceID = strResourceID
	End Property 
	
	Public Property Get GetClientSecret
		GetClientSecret = strClientSecret
	End Property 
	
	Public Property Get GetAuthURL
		GetAuthURL = strAuthUrlPart1 & strTenantID & strAuthUrlPart2
	End Property 
	
	Public Property Get GetSiteURL
		GetSiteURL = strSiteURL
	End Property 
	
	Public Property Get GetSiteDomain
		GetSiteDomain = strDomain
	End Property 
	
	Public Property Get LastErrorNumber
		LastErrorNumber = errNumber
	End Property
	
	Public Property Get LastErrorDesc
		LastErrorDesc = errDescription
	End Property
	
	Public Property Get LastErrorSource
		LastErrorSource = errSource
	End Property 
	
End Class 