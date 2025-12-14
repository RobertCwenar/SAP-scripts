'Dynamic name and generation for COOIS exception report export
Dim dzien, miesiac, rok, dzis
dzien = Right("0" & Day(Date), 2)
miesiac = Right("0" & Month(Date), 2)
rok = Right(Year(Date), 2)
dzis = dzien & "." & miesiac & "." & rok

Dim fso, folderPath, numer, fileName, f
Set fso = CreateObject("Scripting.FileSystemObject")
folderPath = "C:\Users\robert.cwenar\" '<-- Change to your folder path
If Right(folderPath,1) <> "\" Then folderPath = folderPath & "\"

'Added unique number if file exists
numer = 0
fileName = "COOIS wyjatki z dnia " & dzis & ".xls"
Do While fso.FileExists(folderPath & fileName)
    numer = numer + 1
    fileName = "COOIS wyjatki z dnia " & dzis & " (" & numer & ").xls"
Loop

'Connection to SAP GUI and session
Dim SapGuiAuto, application, connection, potentialSession, sessionsList, i, choice, session, sessionFound
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine

'Get list of active sessions
sessionsList = Array()
i = 0
For Each connection In application.Children
    For Each potentialSession In connection.Children
        ReDim Preserve sessionsList(i)
        Set sessionsList(i) = potentialSession
        i = i + 1
    Next
Next

If i = 0 Then
    MsgBox "No active SAP session found.Please log in to SAP and try again."
    WScript.Quit
End If

'Choose session
msg = "Choose the session to connect:" & vbCrLf
For i = 0 To UBound(sessionsList)
    msg = msg & i+1 & ". Sesja #" & i+1 & vbCrLf
Next

choice = InputBox(msg, "Choose session", "1")
If choice = "" Then WScript.Quit
choice = CInt(choice) - 1

'Try to connect to chosen session or next available
sessionFound = False
For i = choice To UBound(sessionsList)
    On Error Resume Next
    Set session = sessionsList(i)
    session.findById("wnd[0]").maximize
    If Err.Number = 0 Then
        sessionFound = True
        Exit For
    Else
        Err.Clear
    End If
Next

'Added msgBox if no session found
If Not sessionFound Then
    MsgBox "Failed to connect to the selected or next SAP session."
    WScript.Quit
End If

If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject application, "on"
End If

MsgBox "Connected to SAP session #" & i+1

'Coois report generation and export
On Error Resume Next
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "COOIS"
session.findById("wnd[0]").sendVKey 0

'Setting report parameters
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E1").selected = True
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").selected = True
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1").text = "ZTCH"
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST2").text = "ZAK"
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").setFocus

'Execute report
session.findById("wnd[0]/tbar[1]/btn[8]").press

'Export to Excel
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton "&NAVIGATION_PROFILE_TOOLBAR_EXPAND"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem "&PC"

session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press

'Save file with dynamic name
session.findById("wnd[1]/usr/ctxtDY_PATH").text = folderPath
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fileName
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = Len(fileName)
session.findById("wnd[1]/tbar[0]/btn[0]").press

On Error GoTo 0
