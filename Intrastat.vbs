If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If

If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If

If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If

If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
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

' FORMAT DATY SAP 
Function FormatDateSAP(d)
	FormatDateSAP = Right("0" & Day(d), 2) & "." & _
                        Right("0" & Month(d), 2) & "." & _
                        Year(d)
End Function

' FORMAT NAZW PLIKOW
Function FormatDateMaterial(d)
    FormatDateMaterial = Year(d) & "." & Right("0" & Month(d), 2) & "." & Right("0" & Day(d), 2)
End Function

Function FormatDateZakupy(d)
    FormatDateZakupy = Year(d) & "." & Right("0" & Month(d), 2) & "." & Right("0" & Day(d), 2)
End Function

Function FormatDateSprzedaz(d)
    FormatDateSprzedaz = Year(d) & "." & Right("0" & Month(d), 2) & "." & Right("0" & Day(d), 2)
End Function


' NAZWA PLIKOW 
Dim nazwaPliku
nazwaPliku = FormatDateMaterial(Date) & " Sq01_material006.xls"
nazwaPliku1 = FormatDateZakupy(Date) & " SQ01_zakupy007.xls"
nazwaPliku2 = FormatDateSprzedaz(Date) & " SQ01_sprzedaz008.xls"

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "Sq01"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").setFocus
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "Material_0006"

session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem "&PC"
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nazwaPliku
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 31
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").currentCellRow = 1
session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell").firstVisibleRow = 27
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "ZAKUPY_0007"


session.findById("wnd[0]/usr/ctxtRS38R-QNUM").caretPosition = 11
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/ctxtSP$00004-LOW").text = "300000"
session.findById("wnd[0]/usr/ctxtSP$00004-HIGH").text = "399999"
session.findById("wnd[0]/usr/ctxtSP$00009-LOW").text = FormatDateSAP(dataOd)
session.findById("wnd[0]/usr/ctxtSP$00009-HIGH").text = FormatDateSAP(dataDo)
session.findById("wnd[0]/usr/ctxt%LAYOUT").text = "/INTRASTAT"
session.findById("wnd[0]/usr/ctxt%LAYOUT").setFocus
session.findById("wnd[0]/usr/ctxt%LAYOUT").caretPosition = 10
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem "&PC"
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nazwaPliku1
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = Len(nazwaPliku1)
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "SPRZEDAZ_0008"
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").caretPosition = 13
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/ctxtSP$00005-LOW").text = FormatDateSAP(dataOd)
session.findById("wnd[0]/usr/ctxtSP$00005-HIGH").text = FormatDateSAP(dataDo)
session.findById("wnd[0]/usr/ctxtSP$00015-LOW").text = "ue"
session.findById("wnd[0]/usr/ctxtSP$00015-LOW").setFocus
session.findById("wnd[0]/usr/ctxtSP$00015-LOW").caretPosition = 2
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxt%LAYOUT").text = "/INTRASTAT"
session.findById("wnd[0]/usr/ctxt%LAYOUT").setFocus
session.findById("wnd[0]/usr/ctxt%LAYOUT").caretPosition = 10
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem "&PC"
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nazwaPliku2
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = Len(nazwaPliku2)
session.findById("wnd[1]/tbar[0]/btn[0]").press
