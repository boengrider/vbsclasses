Class SP
	
	Private oXML
	Private strAuthUrlPart1
	Private strAuthUrlPart2
	Private vti_bin_clientURL
	Private oHTTP
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
	
	Private Sub Class_Initialize
		
		numHTTPstatus = 0
		strAuthUrlPart1 = "https://accounts.accesscontrol.windows.net/"
		strAuthUrlPart2 = "/tokens/OAuth/2"
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		Set oXML = CreateObject("MSXML2.DOMDocument")
		
	End Sub 
	
	Public Function Init(sSiteUrl,sDomain,sClientID,sClientSecret)
	
		If Right(sSiteUrl,1) = "/" Then
			strSiteURL = sSiteUrl
		Else
			strSiteURL = sSiteUrl & "/"
		End If 
		
		If Left(sDomain,1) = "/" Then
			sDomain = Right(sDomain,Len(sDomain) - 1)
		End If
		If Right(sDomain,1) = "/" Then
			sDomain = Left(sDomain,Len(sDomain) - 1)
		End If 
		strDomain = sDomain
		strClientID = sClientID
		strClientSecret = sClientSecret
		
		If Right(sSiteUrl,1) = "/" Then
			vti_bin_clientURL = sSiteUrl & "_vti_bin/client.svc"
		Else
			vti_bin_clientURL = sSiteUrl & "/_vti_bin/client.svc"
		End If 
		
		GetTenantID      ' Obtain the Tenant/Realm ID
		GetSecurityToken ' Obtain the Security Token
		GetXDigestValue  ' Obtain the form digest value
	
	End Function 
		
	'********************** P R I V A T E   F U N C T I O N S ************************
	Private Function GetTenantID()
	
		Dim part,parts,header
		oHTTP.open "GET",vti_bin_clientURL,False
		oHTTP.setRequestHeader "Authorization","Bearer"
		oHTTP.send
	
		parts = Split(oHTTP.getResponseHeader("WWW-Authenticate"),",")
	
		For Each part In parts 
	
			If InStr(part,"Bearer realm") > 0 Then
				header = Split(part,"=")
				strTenantID = header(1)
				strTenantID = Mid(strTenantID,2,Len(strTenantID) - 2)
			End If 
		
			If InStr(part,"client_id") > 0 Then
				header = Split(part,"=")
				strResourceID = header(1)
				strResourceID = Mid(strResourceID,2,Len(strResourceID) - 2)
			End If 		
		Next
	
	End Function
	
	Private Function GetXDigestValue()
		
		Dim colElements
		oHTTP.open "POST", strSiteURL & "_api/contextinfo", False 
		oHTTP.setRequestHeader "accept","application/atom+xml;odata=verbose"
		oHTTP.setRequestHeader "authorization", "Bearer " & strSecurityToken
		oHTTP.send
		oXML.loadXML oHTTP.responseText
		Set colElements = oXML.getElementsByTagName("d:FormDigestValue")
		strFormDigestValue = colElements.item(0).text 
		
	End Function
	
	Private Function GetSecurityToken
	
		Dim oHTTP,part,parts,tokens,token
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP")
		strURLbody = "grant_type=client_credentials&client_id=" & strClientID & "@" & strTenantID & "&client_secret=" & strClientSecret & "&resource=" & strResourceID & "/" & strDomain & "@" & strTenantID
		oHTTP.open "POST", strAuthUrlPart1 & strTenantID & strAuthUrlPart2, False
		oHTTP.setRequestHeader "User-Agent","Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)"
		oHTTP.setRequestHeader "Host","accounts.accesscontrol.windows.net"
		oHTTP.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		oHTTP.setRequestHeader "Content-Length", CStr(Len(strURLbody))
		oHTTP.send strURLbody
		parts = Split(oHTTP.responseText,",")
		For Each part In parts
			If InStr(part,"access_token") > 0 Then
				tokens = Split(part,":")
				Exit For
			End If
		Next
		
		token = Mid(tokens(1),2,Len(tokens(1)) - 3)
		strSecurityToken = token
		
		
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
	
		
	
	'******************** P U B L I C   F U N C T I O N S ***********************

	Public Function ItemExists(sListName,sQuery,sFieldNameToCount)
		With oHTTP
			.open "GET", strSiteURL & "_api/web/lists/GetByTitle('" & sListName & "')/items?$" & sQuery, False 
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With
		
		If Not oHTTP.status = 200 Then
			ItemExists = False
			Exit Function
		End If 
	
		oXML.loadXML oHTTP.responseText

		If oXML.getElementsByTagName(sFieldNameToCount).length > 0 Then
			ItemExists = True
			Exit Function 
		Else
			ItemExists = False
			Exit Function
		End If 
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
			GetListItem = False
			Exit Function
		End If 
		
		oXML.loadXML oHTTP.responseText
		
		If oXML.getElementsByTagName("d:" & sFieldName).length > 0 Then
			If sFieldValue = oXML.getElementsByTagName("d:" & sFieldName).item(0).text Then
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
	
	
	Public Function DownloadFile(sServerRelFilePath,sSaveAsPath)
	
		Dim oHTTP
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP")
		With oHTTP
			.open "GET", strSiteURL & "_layouts/15/download.aspx?SourceUrl=https://" & strDomain & sServerRelFilePath
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.send
		End With
		
		If oHTTP.status = 200 Then
			Set oSTREAM = CreateObject("ADODB.Stream")
			oSTREAM.Open
			oSTREAM.Type = 1
			oSTREAM.Write oHTTP.responseBody
			oSTREAM.SaveToFile sSaveAsPath
			oSTREAM.Close
		End If
		
	End Function 
	
	Public Function GetFileCountXML(sServerRelDirPath,boolRetColObject,ByRef colObject)
	
		Dim oHTTP,oXML,colElements
		sServerRelDirPath = Strip(sServerRelDirPath)
		Set oHTTP = CreateObject("MSXML2.XMLHTTP")
		Set oXML = CreateObject("MSXML2.DOMDocument")
		
		With oHTTP
			.open "GET", oSP.GetSiteURL & "_api/web/GetFolderByServerRelativeUrl('" & sServerRelDirPath & "')/Files", False
			.setRequestHeader "Authorization", "Bearer " & oSP.GetToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With

		If oHTTP.status = 200 Then
			oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
			oXML.loadXML oHTTP.responseText
			Set colElements = oXML.getElementsByTagName("d:Name")
			If Not boolRetColObject Then ' Return file count only
				GetFileCount = colElements.length ' return the file count
				Exit Function
			Else 
				Set colObject = colElements
				GetFileCount = colElements.length
				Exit Function
			End If 
		End If 
		
		GetFileCount = -1 ' Return error
			
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
			debug.WriteLine "HTTP->OK"
			Exit Function
		Else 
			AddListItem = oHTTP.status
			debug.WriteLine "HTTP-> " & oHTTP.responseText
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
	Public Property Get Errcode
	
		Errcode = numHTTPstatus
		
	End Property
	
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
	
End Class 
