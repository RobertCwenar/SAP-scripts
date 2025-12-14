'==== DYNAMICZNA DATA I UNIKALNA NAZWA PLIKU ====
Dim dzien, miesiac, rok, dzis
dzien = Right("0" & Day(Date), 2)
miesiac = Right("0" & Month(Date), 2)
rok = Right(Year(Date), 2)
dzis = dzien & "." & miesiac & "." & rok 

Dim fso, folderPath, numer, fileName, f, re, matches, maxNum
Set fso = CreateObject("Scripting.FileSystemObject")
folderPath = "C:\Users\robert.cwenar\"

' Upewnij się, że folderPath kończy się backslashem
If Right(folderPath,1) <> "\" Then folderPath = folderPath & "\"

'==== RegExp do nadawania numerów istniejącym plikom ====
Set re = New RegExp
re.Pattern = "ydrzewo (\d+) z d " & dzis & "\.xls"
re.IgnoreCase = True
re.Global = False

maxNum = 0
For Each f In fso.GetFolder(folderPath).Files
    If re.Test(f.Name) Then
        If CInt(re.Execute(f.Name)(0).SubMatches(0)) > maxNum Then
            maxNum = CInt(re.Execute(f.Name)(0).SubMatches(0))
        End If
    End If
Next

numer = maxNum + 1
fileName = "ydrzewo " & numer & " z d " & dzis & ".xls"

'==== START SAP GUI ====
Dim SapGuiAuto, application, connection, potentialSession, sessionsList(), i, choice, session, sessionFound
Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine

'==== Pobranie wszystkich sesji ====
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
    MsgBox "Nie znaleziono żadnej aktywnej sesji SAP."
    WScript.Quit
End If

'==== Wyświetlenie listy sesji i wybór ====
Dim msg
msg = "Wybierz sesje SAP do uzycia:" & vbCrLf
For i = 0 To UBound(sessionsList)
    msg = msg & i+1 & ". Sesja #" & i+1 & vbCrLf
Next

choice = InputBox(msg, "Wybór sesji", "1")
If choice = "" Then WScript.Quit
choice = CInt(choice) - 1

'==== Próba połączenia z wybraną sesją lub kolejną wolną ====
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

If Not sessionFound Then
    MsgBox "Nie udało się połaczyc z wybrana ani kolejna sesja SAP."
    WScript.Quit
End If

If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject application, "on"
End If

MsgBox "Polaczono z sesja SAP #" & i+1

'==== EKSPORT DANYCH Z SAP ====
On Error Resume Next
session.findById("wnd[0]").maximize

' Sprawdzenie czy kontener GRID istnieje
If session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell", False) Is Nothing Then
    MsgBox "Nie znaleziono GRID1 w SAP. Sprawdz transakcje."
    WScript.Quit
End If


session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,"DELKZ"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "DELKZ"
session.findById("wnd[0]/tbar[1]/btn[29]").press

session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "FE"
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 2
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[45]").press

session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fileName
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = Len(fileName)
session.findById("wnd[1]/tbar[0]/btn[0]").press

