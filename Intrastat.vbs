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


' ===== DATA: POPRZEDNI MIESIĄC =====
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


' ===== FORMAT DATY POD SAP =====
Function FormatDateSAP(d)
	FormatDateSAP = Right("0" & Day(d), 2) & "." & _
                        Right("0" & Month(d), 2) & "." & _
                        Year(d)
End Function

Function FormatDateMB51(d)
    FormatDateMB51 = Year(d) & "." & Right("0" & Month(d), 2) & "." & Right("0" & Day(d), 2)
End Function


Function FormatDateMaterial(d)
    FormatDateMaterial = Year(d) & "." & Right("0" & Month(d), 2) & "." & Right("0" & Day(d), 2)
End Function

Function FormatDateZakupy(d)
    FormatDateZakupy = Year(d) & "." & Right("0" & Month(d), 2) & "." & Right("0" & Day(d), 2)
End Function

Function FormatDateSprzedaz(d)
    FormatDateSprzedaz = Year(d) & "." & Right("0" & Month(d), 2) & "." & Right("0" & Day(d), 2)
End Function

' ===== NAZWA PLIKOW =====
Dim nazwaPliku
nazwaPliku = FormatDateMaterial(Date) & " Sq01_material006.xls"
nazwaPliku1 = FormatDateZakupy(Date) & " SQ01_zakupy007.xls"
nazwaPliku2 = FormatDateSprzedaz(Date) & " SQ01_sprzedaz008.xls"
nazwaPliku3 = FormatDateMB51(Date) & " MB51.xls"

' ===== SAP =====
' ===== FUNKCJE POMOCNICZE =====
Sub WaitBusy()
    Do While session.Busy
        WScript.Sleep 200
    Loop
    WScript.Sleep 300
End Sub

Function WaitFor(id)
    Dim obj
    Do
        On Error Resume Next
        Set obj = session.findById(id)
        On Error GoTo 0
        If Not obj Is Nothing Then Exit Do
        WScript.Sleep 200
    Loop
    Set WaitFor = obj
End Function


' ===== MB51 =====
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "MB51"
session.findById("wnd[0]").sendVKey 0
WaitBusy

session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "Y001"
session.findById("wnd[0]/usr/ctxtBWART-LOW").text = "101"
session.findById("wnd[0]/usr/ctxtBWART-HIGH").text = "102"
session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = FormatDateSAP(dataOd)
session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = FormatDateSAP(dataDo)
session.findById("wnd[0]/usr/ctxtALV_DEF").text = "/Intrastat"

session.findById("wnd[0]/tbar[1]/btn[8]").press
WaitBusy

session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select

WaitFor("wnd[1]/usr/ctxtDY_FILENAME")

session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/tbar[0]/btn[0]").press

WaitFor("wnd[1]/usr/ctxtDY_FILENAME").Text = nazwaPliku3
session.findById("wnd[1]/tbar[0]/btn[0]").press

WaitBusy

session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press


' ===== SQ01 =====
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "SQ01"
session.findById("wnd[0]").sendVKey 0
WaitBusy

session.findById("wnd[0]/tbar[1]/btn[19]").press

WaitFor("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell
WaitBusy

session.findById("wnd[0]/tbar[1]/btn[19]").press

WaitFor("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell
WaitBusy

session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "Material_0006"
session.findById("wnd[0]/tbar[1]/btn[8]").press
WaitBusy

session.findById("wnd[0]/tbar[1]/btn[8]").press
WaitBusy

' ===== EXPORT =====
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem "&PC"

WaitFor("wnd[1]/usr/ctxtDY_FILENAME")

session.findById("wnd[1]/tbar[0]/btn[0]").press
WaitFor("wnd[1]/usr/ctxtDY_FILENAME").Text = nazwaPliku
session.findById("wnd[1]/tbar[0]/btn[0]").press

WaitBusy

session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press


' ===== ZAKUPY =====
session.findById("wnd[0]/tbar[1]/btn[19]").press

WaitFor("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell
WaitBusy

session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "ZAKUPY_0007"
session.findById("wnd[0]/tbar[1]/btn[8]").press
WaitBusy

WaitFor("wnd[0]/usr/ctxtSP$00004-LOW").Text = "300000"
session.findById("wnd[0]/usr/ctxtSP$00004-HIGH").text = "399999"
session.findById("wnd[0]/usr/ctxtSP$00009-LOW").text = FormatDateSAP(dataOd)
session.findById("wnd[0]/usr/ctxtSP$00009-HIGH").text = FormatDateSAP(dataDo)

session.findById("wnd[0]/usr/ctxt%LAYOUT").text = "/INTRASTAT"

session.findById("wnd[0]/tbar[1]/btn[8]").press
WaitBusy

session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem "&PC"

WaitFor("wnd[1]/usr/ctxtDY_FILENAME")

session.findById("wnd[1]/tbar[0]/btn[0]").press
WaitFor("wnd[1]/usr/ctxtDY_FILENAME").Text = nazwaPliku1
session.findById("wnd[1]/tbar[0]/btn[0]").press

WaitBusy
