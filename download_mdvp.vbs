'Check and connect to SAP GUI Scripting engine
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
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

'Maximize the main SAP window
session.findById("wnd[0]").maximize

'Navigate to transaction
session.findById("wnd[0]/tbar[0]/okcd").text = "mdvp"
session.findById("wnd[0]").sendVKey 0

'Open selection screen for material or stock view
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").setCurrentCell 2,"TEXT"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "2"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell

'Set the stock parameter
session.findById("wnd[0]/usr/txtS_RECKST-HIGH").text = "4"
session.findById("wnd[0]/usr/txtS_RECKST-HIGH").setFocus
session.findById("wnd[0]/usr/txtS_RECKST-HIGH").caretPosition = 6
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID_HEAD/shellcont/shell").setCurrentCell -1,""
session.findById("wnd[0]/usr/cntlGRID_HEAD/shellcont/shell").selectAll
session.findById("wnd[0]/tbar[1]/btn[27]").press
session.findById("wnd[1]/tbar[0]/btn[6]").press
session.findById("wnd[1]/usr/chkIOCSEL-AV_FIX_PL").selected = false
session.findById("wnd[1]/usr/radIOCSEL-AV_CHECK_ATP").select
session.findById("wnd[1]/usr/radIOCSEL-AV_CHECK_ATP").setFocus
session.findById("wnd[1]/tbar[0]/btn[5]").press

'Sort the results by document number (BSFAK) in descending order
session.findById("wnd[0]/usr/cntlGRID_HEAD/shellcont/shell").setCurrentCell -1,"BSFAK"
session.findById("wnd[0]/usr/cntlGRID_HEAD/shellcont/shell").selectColumn "BSFAK"
session.findById("wnd[0]/usr/cntlGRID_HEAD/shellcont/shell").selectedRows = ""
session.findById("wnd[0]/usr/cntlGRID_HEAD/shellcont/shell").pressToolbarButton "&SORT_DSC"