Option Explicit

Dim vat : Set vat = New EUVATChecker
If vat.CheckVatA("SK","1020349826") Then
	debug.WriteLine "VAT valid"
Else
	debug.WriteLine "VAT invalid"
End If 

	
	
Class EUVATChecker

	Private oHTTP
	Private oXML
	Private oRX
	Private strVatNumPattern
	Private strCtryCodePattern
	Private oRxDeleteWhiteSpace
	
	
	Private Sub Class_Initialize()
		Set oHTTP = CreateObject("MSXML2.XMLHTTP")
		Set oXML = CreateObject("MSXML2.DOMDOCUMENT")
		Set oRX = CreateObject("Vbscript.Regexp")
		Set oRxDeleteWhiteSpace = CreateObject("Vbscript.Regexp")
		strVatNumPattern = "^[a-zA-Z]{1,3}[0-9]*$"
		strCtryCodePattern = "^[a-zA-Z]{1,3}"
		oRxDeleteWhiteSpace.Pattern = "\s+"
		oRxDeleteWhiteSpace.Global = true
	End Sub
	
	Private Sub Class_Terminate
		Set oHTTP = Nothing
		Set oXML = Nothing
		Set oRX = Nothing
	End Sub 
	
	Public Function DeleteWhiteSpace(strVatNum)
		DeleteWhiteSpace = oRxDeleteWhiteSpace.Replace(strVatNum,"")
	End Function 
	
	Public Function CheckVatA(strCtryCode, strVatNum)
		Dim arrRxRslt
		Dim oXmlNode
		
		strCtryCode = Trim(strCtryCode) ' Just in case. If any white space is left in the string vies will return invalid input
		strVatNum = Trim(strVatNum)
		
		Dim strHtmlBody : strHtmlBody = "<s11:Envelope xmlns:s11='http://schemas.xmlsoap.org/soap/envelope/'>" _
		& "<s11:Body><tns1:checkVat xmlns:tns1='urn:ec.europa.eu:taxud:vies:services:checkVat:types'>" _
        & "<tns1:countryCode>" & strCtryCode & "</tns1:countryCode><tns1:vatNumber>" & strVatNum & "</tns1:vatNumber>" _
    	& "</tns1:checkVat></s11:Body></s11:Envelope>"
    	
    	With oHTTP
			.open "POST", "http://ec.europa.eu/taxation_customs/vies/services/checkVatService", False
			.setRequestHeader "Accept", "application/xml"
			.setRequestHeader "Content-Type", "application/xml; charset=utf-8"
			.setRequestHeader "Content-Length", Len(strHtmlBody)
			.send strHtmlBody
		End With
		
		If Not oHTTP.Status = 200 Then
			CheckVatA = -1 ' HTTP error
			Exit Function 
		End If 
		
		debug.WriteLine oHTTP.ResponseText
		oXML.LoadXML oHTTP.ResponseText
		
		For Each oXmlNode In oXML.getElementsByTagName("valid")
			If oXmlNode.text = "true" Then
				CheckVatA = 1 ' VAT valid
				Exit Function
			Else
				CheckVatA = 0 ' VAT invalid
				Exit Function
			End If 
		Next 
		
		CheckVatA = 0

    End Function
    
    	
	Public Function CheckVat(strCtryCodeVatNumber)
		Dim strCtryCode : strCtryCode = ""
		Dim strVatNum : strVatNum = ""
		Dim arrRxRslt
		Dim oXmlNode
		
		strCtryCodeVatNumber = Trim(strCtryCodeVatNumber) ' Just in case. If any white space is left in the string vies will return invalid input
		oRX.Pattern = strVatNumPattern
		oRX.IgnoreCase = True
		
		If Not oRX.Test(strCtryCodeVatNumber) Then
			CheckVat = -2 ' Invalid VAT format
			Exit Function
		End If 
		
		oRX.Pattern = strCtryCodePattern
		oRX.IgnoreCase = True
		
		Set arrRxRslt = oRX.Execute(strCtryCodeVatNumber)
		
		strCtryCode = arrRxRslt(0) ' Since VAT is valid, tere should really be only one match
		strVatNum = Right(strCtryCodeVatNumber,Len(strCtryCodeVatNumber) - Len(strCtryCode))
		
		Dim strHtmlBody : strHtmlBody = "<s11:Envelope xmlns:s11='http://schemas.xmlsoap.org/soap/envelope/'>" _
		& "<s11:Body><tns1:checkVat xmlns:tns1='urn:ec.europa.eu:taxud:vies:services:checkVat:types'>" _
        & "<tns1:countryCode>" & strCtryCode & "</tns1:countryCode><tns1:vatNumber>" & strVatNum & "</tns1:vatNumber>" _
    	& "</tns1:checkVat></s11:Body></s11:Envelope>"
    	
    	With oHTTP
			.open "POST", "http://ec.europa.eu/taxation_customs/vies/services/checkVatService", False
			.setRequestHeader "Accept", "application/xml"
			.setRequestHeader "Content-Type", "application/xml; charset=utf-8"
			.setRequestHeader "Content-Length", Len(strHtmlBody)
			.send strHtmlBody
		End With
	
		If Not oHTTP.Status = 200 Then
			CheckVat = -1 ' HTTP error
			Exit Function
		End If 
		
		oXML.LoadXML oHTTP.ResponseText
		
		For Each oXmlNode In oXML.getElementsByTagName("valid")
			If oXmlNode.text = "true" Then
				CheckVat = 1 ' VAT valid
				Exit Function
			Else
				CheckVat = 0 ' VAT invalid
				Exit Function
			End If 
		Next 
		
		CheckVat = 0

    End Function
    
End Class 
