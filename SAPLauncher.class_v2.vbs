Option Explicit
'No SAPLogon used in this class
Dim oSAP : Set oSAP = CreateObject("Sapgui.ScriptingCtrl.1")
Dim oCON : Set oCON = oSAP.OpenConnection("FQ2 - SAP_VGMF ERP TEST [1010]")
debug.WriteLine oCON.Description

WScript.Sleep 10000

WScript.Quit