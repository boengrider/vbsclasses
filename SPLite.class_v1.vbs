Option Explicit

Dim retval,token
Dim oSPL : Set oSPL = New SharePointLite
retval = oSPL.SharePointLite("https://volvogroup.sharepoint.com/sites/unit-rc-sk-bs-it","40784cc3-ba68-45d0-9891-f3dfa8f04d15","cDES2gLLi%2BBRI/FcUizAyZuGQFQ5p%2B6rrknc3kMBWmE=",False)

If retval <> 0 Then
	debug.WriteLine oSPL.LastErrorNumber
	debug.WriteLine oSPL.LastErrorSource
	debug.WriteLine oSPL.LastErrorDesc
Else ' Initialization OK
	token oSPL.AccessToken
End If 

'Your code goes here



Class SharePointLite
	
	Private oRX 
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
	

	Public Function SharePointLite(sSiteUrl,sClientID,sClientSecret,bRaise)
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
			SharePointLite = 100
			Exit Function
		End If 
		
		
		vti_bin_clientURL = strSiteURL & "_vti_bin/client.svc"
		tmp = Split(strSiteURL,"/")
		strDomain = tmp(2)
		strClientID = sClientID
		strClientSecret = sClientSecret
		
		retval = GetTenantID      ' Obtain the Tenant/Realm ID
		If retval <> 0 Then
			SharePointLite = retval
			Exit Function 
		End If 
		
		retval = GetSecurityToken ' Obtain the Security Token
		If retval <> 0 Then
			SharePointLite = retval
			Exit Function
		End If 
		
		retval = GetXDigestValue  ' Obtain the form digest value
		If retval <> 0 Then
			SharePointLite = retval
			Exit Function
		End If 
		
		
		SharePointLite = 0
	
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
	
	Public Property Get AccessToken
		AccessToken = strSecurityToken
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