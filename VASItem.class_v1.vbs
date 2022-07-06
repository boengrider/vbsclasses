'This class represents one invoice
Class MyItem
	
	Public 	outHeader__(27) ' Header line array
	'0 -> skip
	'1 -> DocumentDate; 2 -> PostingDate; 3 -> Currency; 4 -> Reference; 5 -> GLAccount; 6 -> PARMA; 7 -> SpecialG/L; 8 -> AmountDocumentCurrency
	'9 -> TAXCode; 10 -> TaxAmount; 11 -> CostCenter; 12 -> ProfitCenter; 13 -> Order; 14 -> Serial/ChassiNumber; 15 -> ProductVariant; 16 -> DueDate
	'17 -> Quantity; 18 -> LineText; 19 -> NumberOfDays; 20 -> PaymentTerms; 21 -> PaymentBlock; 22 -> PaymentMethod; 23 -> Allocation; 24 -> TradingPartner
	'25 -> ExchangeRate; 26 -> AmountInLocCur; 27 -> TaxAmountInLocCur
	
	Private outBuffer__ 	' GL lines string delimited with CRLF
	Private invoice__ 		' invoice string
	Private isInvoice__     ' Invoice or credit note. If true -> invoice else credite note
	Private rx__
	Private ok__            ' Initially False, set to true by Consume method indicating this invoice fits the conditions specified for GL lines
	Private ids__			' collection of sharepoint IDs associated with each invoice. CCP has multiple IDs OUT usually only one
	Private type__			' This field holds info about invoice type (CCP,OUT or other). For now we consider only CCP and OUT types
	Private conditionsMatched__ ' This field is for testing purposes. Concatenated conditions e.g 0 | 01 | 02 | 03 - items w/o condition or items only w/ condition 0 are not considered valid
	
	Private Sub Class_Initialize()
		Set ids__ = CreateObject("Scripting.Dictionary")
		Set rx__ = New RegExp
		rx__.Global = True
		ok__ = False 
		outBuffer__ = ""
		type__ = ""
	End Sub
	
	'***************
	'Properties
	'***************
	
	'Setters
	'Index 1 DocDate
	Public Property Let HeaderDocdate(d)
		rx__.Pattern = "[0-9]{4}-[0-9]{2}-[0-9]{2}"
		outHeader__(1) = Replace(rx__.Execute(d)(0),"-","")
	End Property
	
	'Index 2 PostingDate
	Public Property Let HeaderPostingDate(d)
		rx__.Pattern = "[0-9]{4}-[0-9]{2}-[0-9]{2}"
		outHeader__(2) = Replace(rx__.Execute(d)(0),"-","")
	End Property
	
	'Index 3 Currency
	Public Property Let HeaderCurrency(c)
		outHeader__(3) = c
	End Property 
	
	'Index 4 Reference
	Public Property Let HeaderReference(r)
'		rx__.Pattern = "[0-9]{1,}"
'		outHeader__(3) = rx__.Execute(r)(0)
		outHeader__(4) = r
	End Property 
	
	'Index 5 GLAccount
	Public Property Let HeaderGLAccount(a)
		outHeader__(5) = a
	End Property 
	
	'Index 6 Parma
	Public Property Let HeaderParma(p)
		outHeader__(6) = p
	End Property 
	
	'Index 7 SpecialG/L
	Public Property Let HeaderSpecialGL(g)
		outHeader__(7) = g
	End Property 
	
	'Index 8 TotalAmount
	Public Property Let HeaderTotalAmount(n)
		outHeader__(8) = outHeader__(8) + CDbl(Replace(n,".",","))
	End Property 
	
	'Index 9 TaxCode
	Public Property Let HeaderTaxCode(t)
		outHeader__(9) = t
	End Property 
	
	'Index 10 TaxAmount
	Public Property Let HeaderTaxAmount(t)
		outHeader__(10) = t
	End Property
	
	'Index 11 CostCenter
	Public Property Let HeaderCostCenter(c)
		outHeader__(11) = c
	End Property 
	
	'Index 12 ProfitCenter
	Public Property Let HeaderProfitCenter(p)
		outHeader__(12) = p
	End Property
	
	'Index 13 Order
	Public Property Let HeaderOrder(o)
		outHeader__(13) = o
	End Property 
	
	'Index 14 Serial/ChassiNumber
	Public Property Let HeaderSerialChassiNumber(s)
		outHeader__(14) = s
	End Property
	
	'Index 15 ProductVariant
	Public Property Let HeaderProductVariant(v)
		outHeader__(15) = v
	End Property
	
	'Index 16 DueDate
	Public Property Let HeaderDueDate(d)
		outHeader__(16) = d
	End Property 
	
	'Index 17 Quantity
	Public Property Let HeaderQuantity(q)
		outHeader__(17) = q
	End Property
	
	'Index 18 LineText
	Public Property Let HeaderLineText(l)
		outHeader__(18) = l
	End Property
	
	'Index 19 NumberOfDays
	Public Property Let HeaderNumberOfDays(n)
		outHeader__(19) = n
	End Property
	
	'Index 20 PaymentTerms
	Public Property Let HeaderPaymentTerms(p)
		outHeader__(20) = p
	End Property
	
	'Index 21 PaymentBlock
	Public Property Let HeaderPaymentBlock(b)
		outHeader__(21) = b
	End Property 
	
	'Index 22 PaymentMethod
	Public Property Let HeaderPaymentMethod(m)
		outHeader__(22) = m
	End Property
	
	'Index 23 Allocation
	Public Property Let HeaderAllocation(a)
		outHeader__(23) = a
	End Property
	
	'Index 24 TradingPartner
	Public Property Let HeaderTradingPartner(p)
		outHeader__(24) = p
	End Property
	
	'Index 25 ExchangeRate
	Public Property Let HeaderExchangeRate(r)
		outHeader__(25) = r
	End Property 
	
	'Index 26 AmountInLocCurr
	Public Property Let HeaderAmInLocCur(a)
		outHeader__(26) = a
	End Property
	
	'Index 27 TaxAmount
	Public Property Let HeaderTaxAmountInLocCur(a)
		outHeader__(27) = a
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
		outBuffer__ = outBuffer__ & a & ";"
	End Property
	
	Public Property Let LineItemSpecialGL(g)
		outBuffer__ = outBuffer__ & g & ";"
	End Property 
	
	Public Property Let LineItemPARMA(p)
		outBuffer__ = outBuffer__ & p & ";"
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
		outBuffer__ = outBuffer__ & p & ";"
	End Property
	
	Public Property Let LineItemOrder(o)
		outBuffer__ = outBuffer__ & o & ";"
	End Property
	
	Public Property Let LineItemSerialChassiNumber(s)
		outBuffer__ = outBuffer__ & s & ";"
	End Property
	
	Public Property Let LineItemProductVariant(v)
		outBuffer__ = outBuffer__ & v & ";"
	End Property
	
	Public Property Let LineItemDueDate(d)
		outBuffer__ = outBuffer__ & d & ";"
	End Property 
	 
	Public Property Let LineItemQuantity(q)
		outBuffer__ = outBuffer__ & q & ";"
	End Property
	
	Public Property Let LineItemLineText(t)
		outBuffer__ = outBuffer__ & t & ";"
	End property
	
	Public Property Let LineItemNumberOfDays(d)
		outBuffer__ = outBuffer__ & d & ";"
	End Property
	
	Public Property Let LineItemPaymentTerms(t)
		outBuffer__ = outBuffer__ & t & ";"
	End Property 
	
	Public Property Let LineItemPaymentBlock(b)
		outBuffer__ = outBuffer__ & b & ";"
	End Property
	
	Public Property Let LineItemPaymentMethod(m)
		outBuffer__ = outBuffer__ & m & ";"
	End Property 
	
	Public Property Let LineItemAllocation(a)
		outBuffer__ = outBuffer__ & a & ";;;"
	End Property 
	
	Public Property Let LineItemTradingPartner(p)
		outBuffer__ = outBuffer__ & p & ";"
	End Property 
	
	Public Property Let LineItemExchangeRate(e)
		outBuffer__ = outBuffer__ & e & ";"
	End Property
	
	Public Property Let LineItemAmInLocCur(a)
		outBuffer__ = outBuffer__ & a & ";" 
	End Property 
	
	Public Property Let LineItemTaxAmountInLocCur(a)
		outBuffer__ = outBuffer__ & a & ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;" & vbCrLf
	End Property
	
	Public Property Let OK(b)
		ok__ = b
	End Property
	
	Public Property Let DocType(t)
		type__ = t
	End Property 
	
	Public Property Let Id(i)
		ids__.Add i,""
	End Property
	
	Public Property Let ItemCondition(c)
		conditionsMatched__ = conditionsMatched__ & c
	End Property 
	
	
	'Getters
	Public Property Get GetDocType
		GetDocType = type__
	End Property 
	
	Public Property Get GetIds
		GetIds = ids__.Keys
	End Property 
	
	'Returns header string
	Public Property Get GetHeader
		GetHeader = outHeader__(1) & ";" & outHeader__(2) & ";" & outHeader__(3) & ";" & outHeader__(4) & ";" & outHeader__(5) _
		          & ";" & outHeader__(6) & ";" & outHeader__(7) & ";" & GetHeaderTotalAmount & ";" & outHeader__(9) & ";" & outHeader__(10) _
		          & ";" & outHeader__(11) & ";" & outHeader__(12) & ";" & outHeader__(13) & ";" & outHeader__(14) & ";" & outHeader__(15) _
		          & ";" & outHeader__(16) & ";" & outHeader__(17) & ";" & outHeader__(18) & ";" & outHeader__(19) & ";" & outHeader__(20) _
		          & ";" & outHeader__(21) & ";" & outHeader__(22) & ";" & outHeader__(23) & ";" & outHeader__(24) & ";" & outHeader__(25) _
		          & ";" & outHeader__(26) & ";" & outHeader__(27) & ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;" & vbCrLf
	End Property 
	
	'Return line items string
	Public Property Get GetLineItems
		GetLineItems = outBuffer__
	End Property 
	
	'Returns True if item is invoice, otherwise false
	Public Property Get IsInvoice
		IsInvoice = isInvoice__
	End Property 
	
	Private Property Get GetHeaderTotalAmount
		If isInvoice__ Then
			If InStr(CStr(outHeader__(8)),",") > 0 Then
				GetHeaderTotalAmount = Replace(CStr(outHeader__(8)),",","") & "-"
			Else
				GetHeaderTotalAmount = CStr(outHeader__(8)) & "00-"
			End If 
		Else
			If InStr(CStr(outHeader__(8)),",") > 0 Then
				GetHeaderTotalAmount = Replace(CStr(outHeader__(8)),",","")
			Else
				GetHeaderTotalAmount = CStr(outHeader__(8)) & "00"
			End If 
		End If 
	End Property
	
	Public Property Get IsOK
		IsOK = ok__
	End Property
	
	Public Property Get GetInvoice
		GetInvoice = invoice__
	End Property 
	
	Public Property Get GetItemCondition
		GetItemCondition = Left(conditionsMatched__,Len(conditionsMatched__) - 1)
	End property
End Class
