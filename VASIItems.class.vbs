Class MyItems
	
	Private listItems__
	Private invoice__
	Private oRX__
	Private ccpItems__
	Private outItems__
	Private otherItems__
	
	Private Sub Class_Initialize()
		ccpItems__ = 0
		outItems__ = 0
		otherItems__ = 0
		Set oRX__ = New RegExp
		oRX__.Pattern = "[0-9]{1,}"
		oRX__.Global = True 
		Set listItems__ = CreateObject("Scripting.Dictionary")
	End Sub
	
	Private Sub Class_Terminate()
	End Sub 
	
	Public Function Consume(xmlNode__,xmlConfigNode__)
		invoice__ = xmlNode__.selectSingleNode("content").selectSingleNode("m:properties").selectSingleNode("d:Title").text
		
		Select Case UCase(xmlNode__.SelectSingleNode("//content/m:properties/d:DocType").text)
		
			Case "CCP"
				
				If Not listItems__.Exists(invoice__) Then
					listItems__.Add invoice__, New MyItem
					ccpItems__ = ccpItems__ + 1
				End If
				
				'*************************************
				'Header line 
				'************************************* 
				'DocDate -> header
				listItems__.Item(invoice__).HeaderDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
				'PostingDate -> header
				listItems__.Item(invoice__).HeaderPostingDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
				'Currency -> header
				listItems__.Item(invoice__).HeaderCurrency = xmlConfigNode__.SelectSingleNode("//Currency").text
				'Reference -> header
				listItems__.Item(invoice__).HeaderReference = invoice__
				'Parma -> header
				listItems__.Item(invoice__).HeaderParma = xmlConfigNode__.SelectSingleNode("//VendorParma").text
				'TotalAmount -> header
				listItems__.Item(invoice__).HeaderTotalAmount = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SubTotal").text
				'TaxCode -> header
				listItems__.Item(invoice__).HeaderTaxCode = xmlConfigNode__.SelectSingleNode("CaseCCP").SelectSingleNode("TaxCode").text
				'TaxAmount -> header
				listItems__.Item(invoice__).HeaderTaxAmount = "0,00"
				'AmountInLocCurrency -> header
				listItems__.Item(invoice__).HeaderAmInLocCur = ""
				'TradingPartner -> header
				listItems__.Item(invoice__).HeaderTradingPartner = xmlConfigNode__.SelectSingleNode("//TradingPartner").text
				'LineText -> header
				listItems__.Item(invoice__).HeaderLineText = ""
				'PaymentTerms -> header
				listItems__.Item(invoice__).HeaderPaymentTerms = xmlConfigNode__.SelectSingleNode("//PaymentTerms").text
				'*************************************
				'GL Lines (must be in order)
				'*************************************
				'DocDate -> line item
				listItems__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
				'PostingDate -> line item
				listItems__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
				'Currency -> line item
				listItems__.Item(invoice__).LineItemCurrency = xmlConfigNode__.SelectSingleNode("//Currency").text
				'Reference -> line items
				listItems__.Item(invoice__).LineItemReference = invoice__
				'GL -> line item
				listItems__.Item(invoice__).LineItemGLAccount = xmlConfigNode__.SelectSingleNode("//CaseCCP/GL").text
				'TotalAmount -> line item
				listItems__.Item(invoice__).LineItemTotalAmount = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SubTotal").text
				'TaxCode -> line item
				listItems__.Item(invoice__).LineItemTaxCode = xmlConfigNode__.SelectSingleNode("CaseCCP").SelectSingleNode("TaxCode").text
				'TaxAmount -> line item
				listItems__.Item(invoice__).LineItemTaxAmount = "0,00"
				'CostCenter -> line item
				listItems__.Item(invoice__).LineItemCostCenter = xmlConfigNode__.SelectSingleNode("CaseCCP").SelectSingleNode("CC").text
				'ProfitCenter -> line item
				listItems__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseCCP").SelectSingleNode("PC").text
				'Allocation -> line item
				listItems__.Item(invoice__).LineItemAllocation = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:RequestNumber").text
				'TradingPartner -> line item
				listItems__.Item(invoice__).LineItemTradingPartner = ""
				'AmountInLocCurrency -> line item
				listItems__.Item(invoice__).LineItemAmInLocCur = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SubTotal").text
				
				'Additonal properties
				listItems__.Item(invoice__).ItemType = invoice__
				
			Case "OUT"
			
				If Not listItems__.Exists(invoice__) Then
					'Create a new instance of MyItem
					listItems__.Add invoice__, New MyItem
					outItems__ = outItems__ + 1
				End If
				
			
			Case Else
				otherItems__ = otherItems__ + 1 
			
		End Select 
		
	End Function
	
	Public Property Get CCPCount
		CCPCount = ccpItems__
	End Property 
	
	Public Property Get OUTCount
		OUTCount = outItems__
	End Property
	
	Public Property Get Count
		Count = listItems__.Count
	End Property 
	
	Public Property Get Items
		Set Items = listItems__
	End Property 
	 
End Class
'This class represents one invoice
Class MyItem
	
	Public 	outHeader__(13) ' Header line array
	Private outBuffer__ 	' GL lines string delimited with CRLF
	'Following variables will be used tu build header line
	Private invoice__ ' invoice string
	Private isInvoice__       ' Invoice or credit note. If true -> invoice else credite note
	Private rx__
	
	Private Sub Class_Initialize()
		Set rx__ = New RegExp
		rx__.Global = True
		outBuffer__ = ""
	End Sub
	
	'***************
	'Properties
	'***************
	
	'Setters
	'Index 0 DocDate
	Public Property Let HeaderDocdate(d)
		rx__.Pattern = "[0-9]{4}-[0-9]{2}-[0-9]{2}"
		outHeader__(0) = Replace(rx__.Execute(d)(0),"-","")
	End Property
	
	'Index 1 PostingDate
	Public Property Let HeaderPostingDate(d)
		rx__.Pattern = "[0-9]{4}-[0-9]{2}-[0-9]{2}"
		outHeader__(1) = Replace(rx__.Execute(d)(0),"-","")
	End Property
	
	'Index 2 Currency
	Public Property Let HeaderCurrency(c)
		outHeader__(2) = c
	End Property 
	
	'Index 3 Reference
	Public Property Let HeaderReference(r)
'		rx__.Pattern = "[0-9]{1,}"
'		outHeader__(3) = rx__.Execute(r)(0)
		outHeader__(3) = r
	End Property 
	
	'Index 4 Parma
	Public Property Let HeaderParma(p)
		outHeader__(4) = p
	End Property 
	
	'Index 5 & 7 TotalAmount + TAX amount
	Public Property Let HeaderTotalAmount(n)
		outHeader__(5) = outHeader__(5) + CDbl(n)
		outHeader__(7) = outHeader__(5)
	End Property 
	
	'Index 6 TaxCode
	Public Property Let HeaderTaxCode(t)
		outHeader__(6) = t
	End Property 
	
	'Index 8 LineText
	Public Property Let HeaderLineText(t)
		outHeader__(8) = t
	End Property 
	
	'Index 9 PaymentTerms
	Public Property Let HeaderPaymentTerms(p)
		outHeader__(9) = p
	End Property
	
	'Index 10 TradingPartner
	Public Property Let HeaderTradingPartner(p)
		outHeader__(10) = p
	End Property
	
	'Index 11 AmountInLocCurr
	Public Property Let HeaderAmInLocCur(a)
		outHeader__(11) = a
	End Property
	
	'Index 12 TaxAmount
	Public Property Let HeaderTaxAmount(a)
		outHeader__(12) = a
	End Property
	
	'Sets isInvoice__ and invoice__
	Public Property Let ItemType(t)
		If Left(UCase(t),1) = "I" Then
			invoice__ = t
			isInvoice__ = True
		Else
			invoice__ = t 
			isInvoice__ = False
		End If
	End Property
	
	Public Property Let LineItemDocDate(d)
		rx__.Pattern = "[0-9]{4}-[0-9]{2}-[0-9]{2}"
		outBuffer__ = outBuffer__ & Replace(rx__.Execute(d)(0),"-","") & ";"
	End Property
	
	Public Property Let LineItemCurrency(c)
		outBuffer__ = outBuffer__ & c & ";"
	End Property
	
	Public Property Let LineItemReference(r)
'		rx__.Pattern = "[0-9]{1,}"
'		outBuffer__ = outBuffer__ & rx__.Execute(r)(0) & ";"
		outBuffer__ = outBuffer__ & r & ";"
	End Property
	
	Public Property Let LineItemGLAccount(a)
		outBuffer__ = outBuffer__ & a & ";;;"
	End Property 
		
	Public Property Let LineItemTotalAmount(a)
		If isInvoice__ Then
			If InStr(a,",") > 0 Then
				outBuffer__ = outBuffer__ & Replace(a,",","") & ";"
			Else
				outBuffer__ = outBuffer__ & a & "00;"
			End If 
		Else
			If InStr(a,",") > 0 Then
				outBuffer__ = outBuffer__ & Replace(a,",","") & "-;"
			Else
				outBuffer__ = outBuffer__ & a & "00-;"
			End If 
		End If 
	End Property 
	
	Public Property Let LineItemTaxCode(t)
		outBuffer__ = outBuffer__ & t & ";"
	End Property
	
	Public Property Let LineItemTaxAmount(a)
		outBuffer__ = outBuffer__ & a & ";"
	End Property 
	
	Public Property Let LineItemCostCenter(c)
		outBuffer__ = outBuffer__ & c & ";"
	End Property 
	
	Public Property Let LineItemProfitCenter(p)
		outBuffer__ = outBuffer__ & p & ";;;;;;;;;;;"
	End Property 
	
	Public Property Let LineItemAllocation(a)
		outBuffer__ = outBuffer__ & a & ";;"
	End Property 
	
	Public Property Let LineItemTradingPartner(p)
		outBuffer__ = outBuffer__ & p & ";"
	End Property 
	
	Public Property Let LineItemAmInLocCur(a)
		If isInvoice__ Then
			If InStr(a,",") > 0 Then
				outBuffer__ = outBuffer__ & Replace(a,",","") & ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;" & vbCrLf
			Else
				outBuffer__ = outBuffer__ & a & "00;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;" & vbCrLf
			End If 
		Else
			If InStr(a,",") > 0 Then
				outBuffer__ = outBuffer__ & Replace(a,",","") & "-;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;" & vbCrLf
			Else
				outBuffer__ = outBuffer__ & a & "00-;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;" & vbCrLf
			End If 
		End If 
	End Property 
	
	
	
	
	
	'Getters
	'Returns header string
	Public Property Get GetHeader
		GetHeader = outHeader__(0) & ";" & outHeader__(1) & ";" & outHeader__(2) & ";" & outHeader__(3) _
		          & ";;" & outHeader__(4) & ";;" & GetTotalAmount & ";" & outHeader__(6) & ";0,00" _
		          & ";;;;;;" & outHeader__(8) & ";;" & outHeader__(9) & ";;;;" & outHeader__(10) & ";;" _
		          & outHeader__(11) & ";" & outHeader__(12) & ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;" & vbCrLf 
	End Property 
	
	'Return line items string
	Public Property Get GetLineItems
		GetLineItems = outBuffer__
	End Property 
	
	'Returns True if item is invoice, otherwise false
	Public Property Get IsInvoice
		IsInvoice = isInvoice__
	End Property 
	
	Private Property Get GetTotalAmount
		Dim tmp : tmp = Round(outHeader__(5),2)
		
		If isInvoice__ Then
			If InStr(CStr(tmp),",") > 0 Then
				GetTotalAmount = Replace(tmp,",","") & "-"
			Else
				GetTotalAmount = tmp & "00-"
			End If 
		Else
			If InStr(CStr(tmp),",") > 0 Then
				GetTotalAmount = Replace(tmp,",","")
			Else
				GetTotalAmount = tmp & "00"
			End If 
		End If 
	End Property 
	
End Class