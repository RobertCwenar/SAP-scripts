' SAP GIU Connection
If month = 0 Then
    month = 12
    year = year - 1
End If

Dim dataOd, dataDo
dataOd = DateSerial(year, month, 1)
dataDo = DateSerial(year, month + 1, 0)


' Data formats to SAP
Function FormatDateSAP(d)
    FormatDateSAP = Right("0" & Day(d), 2) & "." & _
                    Right("0" & Month(d), 2) & "." & _
                    Year(d)
End Function


' Name file
Dim namefile
namefile = FormatDateSAP(Date) & " mb51_Gus.xls"


' SAP 
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "MB51"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "Y001"
session.findById("wnd[0]/usr/ctxtBWART-LOW").text = "101"
session.findById("wnd[0]/usr/ctxtBWART-HIGH").text = "102"
session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = FormatDateSAP(dataOd)
session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = FormatDateSAP(dataDo)
session.findById("wnd[0]/usr/ctxtALV_DEF").text = "/P02"
session.findById("wnd[0]/usr/ctxtALV_DEF").setFocus
session.findById("wnd[0]/usr/ctxtALV_DEF").caretPosition = 4
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = namefile
session.findById("wnd[1]/tbar[0]/btn[0]").press
