If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If

If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If

If Not IsObject(session) Then
   Set session = connection.Children(0)
End If

If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject application, "on"
End If


' DATA: POPRZEDNI MIESIĄC 
Dim rok, miesiac
rok = Year(Date)
miesiac = Month(Date) - 1

If miesiac = 0 Then
    miesiac = 12
    rok = rok - 1
End If

Dim dataOd, dataDo
dataOd = DateSerial(rok, miesiac, 1)
dataDo = DateSerial(rok, miesiac + 1, 0)


' FORMAT DATY POD SAP 
Function FormatDateSAP(d)
	FormatDateSAP = Right("0" & Day(d), 2) & "." & _
                        Right("0" & Month(d), 2) & "." & _
                        Year(d)
End Function

Function FormatDateMB51(d)
    FormatDateMB51 = Year(d) & "." & Right("0" & Month(d), 2) & "." & Right("0" & Day(d), 2)
End Function


' NAZWA PLIKU 
Dim nazwaPliku
nazwaPliku = FormatDateMB51(Date) & " MB51.xls"


' SAP 
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "MB51"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "Y001"
session.findById("wnd[0]/usr/ctxtBWART-LOW").text = "101"
session.findById("wnd[0]/usr/ctxtBWART-HIGH").text = "102"
session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = FormatDateSAP(dataOd)
session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = FormatDateSAP(dataDo)
session.findById("wnd[0]/usr/ctxtALV_DEF").text = "/Intrastat"
session.findById("wnd[0]/usr/ctxtALV_DEF").setFocus
session.findById("wnd[0]/usr/ctxtALV_DEF").caretPosition = 4
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nazwaPliku
session.findById("wnd[1]/tbar[0]/btn[0]").press