Option Explicit

'Date formatting
Dim dzien, miesiac, rok, dzis
dzien = Right("0" & Day(Date), 2)
miesiac = Right("0" & Month(Date), 2)
rok = Right(Year(Date), 2)
dzis = dzien & "." & miesiac & "." & rok

'Folder path and file name
Dim folderPath, fileName
Dim fso, shell

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

folderPath = "C:\Users\robert.cwenar\"

If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
If Not fso.FolderExists(folderPath) Then
    fso.CreateFolder folderPath
End If

fileName = "prio dsmetal z d " & dzis & ".xls"

'SAP GUI Scripting - connection to session
Dim SapGuiAuto, application, connection
Dim session, potentialSession
Dim sessionsList()
Dim i, choice
Dim sessionFound

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine

'Download list of active sessions
ReDim sessionsList(0)
i = 0

For Each connection In application.Children
    For Each potentialSession In connection.Children
        ReDim Preserve sessionsList(i)
        Set sessionsList(i) = potentialSession
        i = i + 1
    Next
Next

If i = 0 Then
    MsgBox "No active SAP session found."
    WScript.Quit
End If

'Choose session
Dim msg
msg = "Choose SAP session:" & vbCrLf

For i = 0 To UBound(sessionsList)
    msg = msg & i + 1 & ". Session #" & i + 1 & vbCrLf
Next

choice = InputBox(msg, "Choose SAP session", "1")
If choice = "" Then WScript.Quit
If Not IsNumeric(choice) Then WScript.Quit

choice = CInt(choice) - 1
If choice < 0 Or choice > UBound(sessionsList) Then WScript.Quit

sessionFound = False

For i = choice To UBound(sessionsList)
    On Error Resume Next
    Set session = sessionsList(i)
    session.findById("wnd[0]").maximize
    If Err.Number = 0 Then
        sessionFound = True
        Exit For
    End If
    Err.Clear
    On Error GoTo 0
Next

If Not sessionFound Then
    MsgBox "No active SAP session found in SAP."
    WScript.Quit
End If

If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject application, "on"
End If

MsgBox "Connected to SAP session #" & i + 1

'COOIS Report Execution
On Error Resume Next

session.findById("wnd[0]/tbar[0]/okcd").Text = "COOIS"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/" & _
    "ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E1").Selected = True

session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/" & _
    "ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").Selected = True

session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/" & _
    "ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_MATNR-LOW").Text = "P3064440-Z1-S2B2"

session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/" & _
    "ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1").Text = "ZTCH"

session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/" & _
    "ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST2").Text = "ZAK"

session.findById("wnd[0]/tbar[1]/btn[8]").press

'Export to Excel
Set shell = session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell")

shell.pressToolbarButton "&NAVIGATION_PROFILE_TOOLBAR_EXPAND"
shell.pressToolbarContextButton "&MB_EXPORT"
shell.selectContextMenuItem "&PC"

session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press

'SAve file with dynamic name
session.findById("wnd[1]/usr/ctxtDY_PATH").text = folderPath
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fileName
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = Len(fileName)
session.findById("wnd[1]/tbar[0]/btn[0]").press

On Error GoTo 0