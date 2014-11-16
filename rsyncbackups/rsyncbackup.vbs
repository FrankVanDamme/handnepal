'......................................................................................
'... rsyncBackup.vbs 1.04 .................. Autor: Karsten Violka kav@ctmagazin.de ...
'... c't 9/06 .........................................................................
'......................................................................................
'
'--------------------------------------------------------------------------------------
' Bekannte Probleme:
'   -- rsync kopiert keine geöffneten Dateien
'   -- rsync kopiert nur Pfade bis zu einer Länge von 260 Zeichen.
'   -- rsync kopiert keine NTFS-Spezialitäten (Junctions, Streams, Sparse Files)

' Skript mit niedriger Priorität starten: 
' 	start /min /belownormal cscript.exe rsyncBackup.vbs
'--------------------------------------------------------------------------------------

Option Explicit
'--------------------------------------------------------------------------------------
'----- Konfiguration ------------------------------------------------------------------
'--------------------------------------------------------------------------------------
Dim sourceFolders
' Quellverzeichnisse
' Wichtig: Geben Sie bei den Quellpfaden keinen abschließenden Backslash an, damit
'   rsync im Backup-Ziel für jede Quelle einen separatenUnterordner erstellt. 
'Beispiel:
'sourceFolders = Array("e:\text", "c:\Dokumente und Einstellungen\Karsten")
sourceFolders = Array("S:\backups\sharada-E\Documents")

Dim excludeFiles
excludeFiles = Array("Cache", "parent.lock", "Temp*")

' Das Zielverzeichnis sollte sich auf einem mit NTFS formatierten Laufwerk befinden
' Beispiel:
'const DESTINATION="e:\rsyncbackup"
const DESTINATION="c:\backups"

' Anzahl der aufbewahrten Backups:
const STAGE0_HOURLY =8 
const STAGE1_DAILY = 14
const STAGE2_WEEKLY = 10

'Ergänzung: Wenn Sie die Konstante COMPARE_CHECKSUMS auf true setzen,
'ruft das Skript rsync mit dem Schalter --checksum auf (siehe Manpage). Um die Menge
'der Dateien zu ermitteln, die es beim inkrementellen Backup kopiert,
'orientiert sich rsync normalerweise am Zeitpunkt der letzten Änderung. Mit dem gesetzten
'Schalter liest es stattdessen alle Dateien komplett ein, erstellt Prüfsummen und
'vergleicht den tatsächlichen Inhalt.

'Dieser Modus kann aber erheblich mehr Zeit in Anspruch nehmen.
 
'Die Option kann als Ersatz für die fehlende Verify-Funktion dienen: Wenn Sie in der
'log-Datei feststellen, dass rsync Dateien erneut kopiert, obwohl sie seit der
'letzten Sicherung nicht geändert wurden, könnten die Dateien auf dem Backupmedium
'verfälscht worden sein.

const COMPARE_CHECKSUMS=false

'Wenn Sie mehrere Quellordner sichern, die denselben Namen tragen, vermischt rsync
'deren Inhalte standardmäßig im selben Backupverzeichnis. Die Konstante FULL_PATHNAME
'aktiviert den rsync-Parameter "R", der bewirkt, dass rsync für jeden Quellpfad den
'absoluten Pfad im Zielverzeichnis anlegt.
'Wenn Sie beispielsweise zwei Ordner namens "text" auf den Laufwerken E: und F: in
'den Zielordner U:\backup sichern, sieht das Ergebnis etwa so aus:
'
'	U:\backup\2006-05-08~15\cygdrive\e\text
'	U:\backup\2006-05-08~15\cygdrive\f\text

const FULL_PATHNAME=false

'--------------------------------------------------------------------------------------
'----- ENDE Konfiguration -------------------------------------------------------------
'--------------------------------------------------------------------------------------
const STAGE1_DAILY_FOLDER =  "\_daily"
const STAGE2_WEEKLY_FOLDER=  "\_weekly"
const STAGE3_MONTHLY_FOLDER= "\_monthly"
' Konstanten für ADO
const adVarChar = 200
const adDate = 7
' Feldnamen fürs RecordSet
Dim rsFieldNames
rsFieldNames = Array("name", "date")

'---- Global verwendete Variablen
Dim fso, wsh, wshEnv

set fso = CreateObject("Scripting.FileSystemObject")
set wsh = CreateObject("WScript.Shell")

' Wenn die Umgebungsvariable CYGWIN=NONTSEC gesetzt ist, verändert rsync die Zugriffsrechte
' der Backups nicht. Normalerweise setzt die Cygwin-Bibliothek eigene ACLs,
' um die Unix-Zugriffsrechte abzubilden.

Set wshEnv = wsh.Environment("process") 
wshEnv("CYGWIN")= "NONTSEC"

'---- Die Log-Datei wird im Profilverzeichnis erstellt, etwa:
'---- c:\Dokumente und Einstellungen\Klaus\rsyncBackup.log
Dim logFile
logFile = wsh.ExpandEnvironmentStrings("%userprofile%") & "\rsyncBackup.log"

Dim strSourceFolder, recentBackupFolder, strDateFolder, strDestinationFolder
Set recentBackupFolder = Nothing
Dim strCmd, cmdResult

logAppend(vbCRLf & "-------- Start: " & Now & " --------------------------------------------")

checkFolders()
strDateFolder = getDateFolderName()

strDestinationFolder = DESTINATION & "\~" & strDateFolder ' Zielordner zunächst Tilde voranstellen
Set recentBackupFolder = getRecentFolder(DESTINATION)

'-- per Dry-Run prüfen, ob sich der Inhalt eines der Quellordner geändert hat
If sourceChanged() Then
	strCmd=getRsyncCmd(false)
	logAppend("--- rsync-Befehlszeile:")
	logAppend(strCmd)
	cmdResult=callCmd(strCmd)
	logAppend("--- Ausgabe von rsync:" & vbCrLf & toCrLf(removePathLines(cmdResult(1))))
	
	If Len(cmdResult(2)) > 0 Then	
		logAppend("--- Fehlermeldungen:" & vbCrLf & toCrLf(cmdResult(2)))
	End If
	
	logAppend("--- Errorlevel: " & cmdResult(0))
	' Zielordner umbenennen und Tilde entfernen
	fso.MoveFolder strDestinationFolder, DESTINATION & "\" & strDateFolder
Else
	logAppend("--- nichts Neues")
End If

'-- Backups rotieren und alte Backups löschen
rotate getFolderObject(DESTINATION), _
		getFolderObject(DESTINATION & STAGE1_DAILY_FOLDER), STAGE0_HOURLY, "d"
rotate getFolderObject(DESTINATION & STAGE1_DAILY_FOLDER), _
		getFolderObject(DESTINATION & STAGE2_WEEKLY_FOLDER), STAGE1_DAILY, "ww"
rotate getFolderObject(DESTINATION & STAGE2_WEEKLY_FOLDER), _
		getFolderObject(DESTINATION & STAGE3_MONTHLY_FOLDER), STAGE2_WEEKLY, "m"
logAppend("-------- Fertig: " & Now & " --------------------------------------------")

'---------------------------------------------------------------------------------------
'--- Funktionen ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------

'--- checkFolders() -------------------------------------------------------------------
' Prüft ob die eingetragenen Pfade plausibel sind.
Function checkFolders()
	Dim aSourceFolder
	For Each aSourceFolder in sourceFolders
		If Not fso.FolderExists(aSourceFolder) Then
			criticalErrorHandler "checkFolders()", "Quellordner '" & aSourceFolder & "' existiert nicht.", 0, ""
		End If
	Next
	
	If Not fso.DriveExists(fso.getDriveName(DESTINATION)) Then
		criticalErrorHandler "checkFolders()", "Ziellaufwerk " & fso.getDriveName(DESTINATION) & " nicht gefunden", 0, ""
	End If
	
	dim d, f
	
	If Not fso.getDrive(fso.getDriveName(DESTINATION)).FileSystem = "NTFS" Then
		logAppend("--- Warnung: Zielpfad " & DESTINATION & " liegt nicht auf einem NTFS-Laufwerk!")
		logAppend("--- Warnung: rsync erstellt dort keine Hard-Links, sondern vollständige Kopien")
	End If
End Function

'--- sourceChanged() -------------------------------------------------------------------
' Liefert "true", wenn ein Problelauf von rsync ermittelt, dass in den Quellordnern
' seit dem letzten Backup Dateien geändert wurden.
Function sourceChanged()
	dim strCmd, cmdResult, arrayOutput
	cmdResult = callCmd(getRsyncCmd(true)) ' Kommando mit dryRun aufbauen
	strCmd=removePathLines(cmdResult(1))
	arrayOutput=Split(strCmd, "" & Chr(10) & "", -1, 1)
	'-- wenn schon in der vierten Zeile "sent" steht, hat sich nichts geändert
	If Left(arrayOutput(3), 4) = "sent" Then
	  sourceChanged=false
	Else
	  sourceChanged=true
	End If
End Function

'--- getRsyncCmd() ----------------------------------------------------------------------
' Baut das rsync-Kommando zusammen. Der Parameter "true" schaltet den dryRun-Modus ein,
' der einen Probelauf startet.
'
' In Version 1.01 habe ich den Schalter "b" wieder entfernt: Er bewirkt, dass
' rsync in neuen Ordnern Backup-Dateien geänderter Dateien vorhält, die auf eine
' Tilde "~" enden. Ohne den Schalter wird die Ausgabe von rsync allerdings sehr
' unübersichtlich: rsync listet dann jedes Mal alle durchsuchten Quellverzeichnisse auf,
' egal, ob es dort etwas Neues gibt. Die Funktion removePathLines() filtert diese
' überflüssigen Zeilen wieder raus.

' Verwendete rsync-Parameter:
'   a   Archiv-Modus   Quellen rekursiv und vollständig kopieren
'   v   Verbose        Ausführliche Ausgabe, listet alle neu übertragenen Dateien auf
'   c                  Optional, rsync berechnet Checksummen und vergleicht damit die
'                      Dateiinhalte, um die Menge der zu kopierenden Dateien zu bestimmen
'   R 	relative       Legt im Ziel für jeden Quellordner den vollen Pfad an
'   n   Dryrun


Function getRsyncCmd(dryRun)
	dim cmd, aSourceFolder, aExcludeFile
	cmd = wsh.ExpandEnvironmentStrings("%comspec%") & " /c rsync -av"
	
	If (FULL_PATHNAME = true) Then
		cmd = cmd & "R"
	End If
	
	If (COMPARE_CHECKSUMS = true) Then
		cmd = cmd & "c"
	End If
	
	If (dryRun = true) Then
		cmd = cmd & "n"
	End If
	
	If Not recentBackupFolder Is Nothing Then
		cmd = cmd & " --link-dest=""" _
			& toCygwinPath(recentBackupFolder.Path) & """"
	End If
	
	For Each aExcludeFile in excludeFiles
		cmd = cmd & " --exclude """ & aExcludeFile & """"
	Next
	
	For Each aSourceFolder in sourceFolders
		cmd = cmd & " """ & toCygwinPath(aSourceFolder) & """"
	Next
	
	cmd = cmd & " """ & toCygwinPath(strDestinationFolder) & """"
	
	getRsyncCmd = cmd
End Function

'--- getDateFolderName()------------------------------------------------------------
' Generiert einen Ordnernamen mit dem aktuellen Datum und der Uhrzeit.
Function getDateFolderName()
	Dim jetzt
	jetzt = Now()
	getDateFolderName = Year(jetzt) & "-" & addLeadingZero(Month(jetzt))_
		& "-" & addLeadingZero(Day(jetzt))_
		& "_"	& addLeadingZero(Hour(jetzt))_
		& "~" & addLeadingZero(Minute(jetzt))
End Function

'--- addLeadingZero(number) -------------------------------------------------------------
' Fügt bei Zahlen < 10 führende Null ein.
Function addLeadingZero(number)
	If number < 10 Then
		number = "0" & number
	End If 
	addLeadingZero = number
End Function

'--- getFolderObject(path) -------------------------------------------------------------
' Liefert zum übergebenen Pfad-String ein WSH-Objekt vom Typ Folder
' Wenn das Verzeichnis noch nicht existiert, wird es angelegt.
Function getFolderObject(path)
	If (fso.FolderExists(path)) Then
		Set getFolderObject = fso.GetFolder(path)
	Else
		logAppend("--- Erstelle Ordner: " & path)
		On Error Resume Next
		Set getFolderObject = fso.CreateFolder(path)
		
		If Err.Number <> 0 Then
			On Error Goto 0
			criticalErrorHandler "getFolderObject()", "Konnte Zielordner nicht erstellen", Err.Number, Err.Description
		End If
		
		On Error Goto 0
	End If
End Function

'--- toCygwinPath(String) -----------------------------------------------------------------
' Wandelt einen Windows-Pfad in das Format, das Cygwin erwartet
Function toCygwinPath(path)
	Dim driveLetter, newPath
	driveLetter = Left(fso.GetDriveName(path), 1)
	newPath = Replace(path, "\", "/")
	newPath = Mid(newPath, 4)
	toCygwinPath = "/cygdrive/" & driveLetter & "/" & newPath
End Function

'--- toCrLf(String) -----------------------------------------------------------------------
' Ersetzt den von rsync ausgegebenen Unix-Zeilenumbruch (LF)
' durch das Windows-übliche Format (CRLF)
Function toCrLf(strText)
	toCrLf = Replace(strText, vbLf, vbCrLf)
End Function

'--- removePathLines(String) -----------------------------------------------------------------------
' Entfernt alle Zeilen, die auf einen Backslash enden.
' rsync gibt normalerweise alle Pfade aus, die es auf neue Dateien überprüft,
' auch wenn sich dort gar nichts geändert hat. Diese Routine entfernt diese Zeilen,
' damit die Log-Datei übersichtlich bleibt.
Function removePathLines(strText)
	Dim arrayText, line
	arrayText=Split(strText, "" & Chr(10) & "", -1, 1) ' Die Ausgabe muss im Unix-Format
							' vorliegen, mit LF als Zeilentrenner.
	For Each line in arrayText
		If Not Right(line, 1) = "/" Then
			removePathLines = removePathLines & line & vbLF
		End If
	Next
End Function

'--- logAppend(String) --------------------------------------------------------------------
' hängt den übergebenen Text an die Log-Datei an
Function logAppend(string)
	const forAppend = 8
	dim f, errnum
	
	On Error Resume Next	
	Set f = fso.OpenTextFile(logFile, forAppend, true)
	errnum = Err.Number
	
	On Error Goto 0
	If errnum = 0 Then
		f.WriteLine(string)
		f.Close()
	Else
		Err.Raise 1, "logAppend", "Konnte Logdatei nicht öffnen"
	End If
End Function

'--- getRecentFolder(String) ---------------------------------------------------------------
' Sortiert die im übergebenen Pfad enthaltenen Ordner nach Datum und liefert das jüngste
' Ordner-Objekt zurück
' Parameter: Pfad als String
Function getRecentFolder(path)
	Dim destinationFolder, rs
	Set destinationFolder = getFolderObject(path)
	Set rs=newFolderRecordSet(destinationFolder)
	
	If Not (rs.Eof) Then
		rs.sort = "date DESC"		' absteigend nach Erstellungszeitpunkt sortieren 
		rs.MoveFirst
		Set getRecentFolder= fso.GetFolder(rs.fields("name"))
	Else
		Set getRecentFolder = Nothing
	End If
	rs.Close
	Set rs = Nothing
End Function

'--- newFolderRecordSet(Folder-Objekt) -----------------------------------------------------
' Füllt die Unterordner der übergebenen Folder-Objekts in ein neues RecordSet-Objekt,
' das zum Sortieren verwendet wird.

Function newFolderRecordSet(folder)
	Dim aFolder
	Set newFolderRecordSet = CreateObject("ADODB.Recordset")
	newFolderRecordSet.Fields.Append "name", adVarChar, 255
	newFolderRecordSet.Fields.Append "date", adDate
    newFolderRecordSet.Open
	
	For Each aFolder in folder.SubFolders
		If Left(aFolder.Name, 2) = "20" Then	' nur die Datumsordner in die Liste aufnehmen
			newFolderRecordSet.addnew rsFieldNames, Array(aFolder.Path, aFolder.DateCreated)
		End if
	Next	
End Function

'--- rotate(fromFolder, toFolder, numberToKeep, diffInterval) ------------------------------
' Verschiebt oder löscht die Backup-Ordner. Fürjedes Zeitintervall (Tag, Woche, Monat) wird
' jeweils das zuletzt erstellte Backup archiviert.
'
Function rotate(fromFolder, toFolder, numberToKeep, diffInterval)
	Dim rs, aFolder, lastFolder, i, recentBackup, errNr
	Set rs=newFolderRecordSet(fromFolder)
	If Not (rs.Eof) Then
		rs.Sort = "date DESC"
		rs.MoveFirst
		i = 0
		Do until rs.Eof
			If i >= numberToKeep Then
				'MsgBox("übrig:" & rs.fields("name"))
				'Das jüngste Backup dieses Datums aus dem toFolder holen. Wenn neuer, ersetzen.
				Set recentBackup = getRecentBackupForDate(toFolder, rs.fields("date"), diffInterval)
				On Error Resume Next
				If Not recentBackup Is Nothing Then
					' Wenn das gewählte Backup vom selben Zeitintervall (Tag) ist und
					' später erstellt wurde, soll es das Backup im Zielordner ersetzen.
					If DateDiff("s", recentBackup.DateCreated, rs.fields("date")) > 0 Then
						'MsgBox("selber Tag & neuer: bewegen")
						logAppend("--- bewege " & rs.fields("name") & " nach " & toFolder.Path)
						fso.MoveFolder fso.GetFolder(rs.fields("name")), toFolder.Path & "\"
						If Err.Number <> 0 Then 
							ErrNr=Err.Number
							On Error Goto 0
							criticalErrorHandler "rotate()", "Konnte Ordner nicht bewegen", Err.Number, Err.Description
						End If
						'MsgBox("Vorgänger löschen.")
						logAppend("--- Vorgänger löschen " & recentBackup)
						fso.DeleteFolder recentBackup, true
						If Err.Number <> 0 Then
							On Error Goto 0
							criticalErrorHandler "rotate()", "Konnte Ordner nicht löschen", Err.Number, Err.Description
						End If					
					Else
						logAppend("--- lösche " & rs.fields("name"))
						'MsgBox("selber Tag & älter: weg damit.")
						fso.DeleteFolder fso.GetFolder(rs.fields("name")), true
					
						If Err.Number <> 0 Then 
							On Error Goto 0
							criticalErrorHandler "rotate()", "Konnte Ordner nicht löschen", Err.Number, Err.Description
						End If
					End If
				Else
					' Vom diesem Tag existiert noch kein Backup
					'MsgBox("noch nicht da, bewegen!")
					logAppend("--- bewege " & rs.fields("name") & " nach " & toFolder.Path)
					fso.MoveFolder fso.GetFolder(rs.fields("name")), toFolder.Path & "\"
					If Err.Number <> 0 Then 
						On Error Goto 0
						criticalErrorHandler "rotate()", "Konnte Ordner nicht bewegen", Err.Number, Err.Description
					End If	
				End If
				On Error Goto 0
			End If
			i = i + 1
			rs.MoveNext
		Loop
	End If
	rs.Close
	Set rs = Nothing
End Function

'--- getRecentBackupForDate(folderObj, aDate, diffInterval) -----------------------------
' Sortiert die Unterverzeichnisse mit Hilfe des ADO RecordSet und liefert
' das das letzte Backup des angegeben Tages/der Woche/des Monats --> diffInterval
Function getRecentBackupForDate(folderObj, aDate, diffInterval)
	Dim rs, exitLoop
	Set getRecentBackupForDate = Nothing
	Set rs=newFolderRecordSet(folderObj)
	If Not (rs.Eof) Then
		rs.Sort = "date DESC"
		rs.MoveFirst
		exitLoop=false 
		Do until rs.Eof Or exitLoop
			If DateDiff(diffInterval, rs.fields("date"), aDate) = 0 Then
				set getRecentBackupForDate = fso.GetFolder(rs.fields("name"))
				exitLoop = true
			End If
		   rs.MoveNext
		Loop	  
	End If
	rs.Close
	Set rs = Nothing
End Function

'--- criticalErrorHandler(source, description, errNumber, errDescription) ---------------
' Kritischen Fehler loggen und Programm abbrechen. Vor dem Aufruf muss die
' Fehlerbehandlung mit "On Error Goto 0" wieder eingeschaltet werden, damit das Skript
' mit dem neu erzeugten Fehler abbricht.
Function criticalErrorHandler(source, description, errNumber, errDescription)
	logAppend("--- Fehler: Funktion " & source & ", " & description)
	logAppend("            Err.Number: " & errNumber & " Err.Description:" & errDescription)
	logAppend("-------- Stop: " & Now & " --------------------------------------------")
	Err.Raise 1, source, description
End Function


'--- callCmd(strCommand) ----------------------------------------------------------------
' Führt Kommandozeilenbefehl aus und liefert Array zurück:
' Index 0: Errorlevel
' Index 1: Ausgabe
' Index 2: Fehlerausgabe
Function callCmd(strCommand)
	Dim strTmpFile, strTmpFile2, outputFile, result, strOutput, strOutput2, failed
	
	strTmpFile = fso.GetSpecialFolder(2) & "\" & fso.GetTempName
	strTmpFile2 = fso.GetSpecialFolder(2) & "\" & fso.GetTempName
	
	strOutput = ""
	strOutput2 = ""
	strCommand = strCommand & " 1>""" & strTmpFile & """ 2>""" & strTmpFile2 & """"
	
	result = wsh.Run(strCommand, 0, true)
	
	If fso.FileExists(strTmpFile2) Then
		If fso.GetFile(strTmpFile2).Size > 0 Then
			Set outputFile = FSO.OpenTextFile(strTmpFile2)
			strOutput2 = outputFile.Readall
			outputFile.Close
			deleteInsistently(strTmpFile2)
		End If
	End If
	
	If fso.FileExists(strTmpFile) Then
		If fso.GetFile(strTmpFile).Size > 0 Then
			Set outputFile = FSO.OpenTextFile(strTmpFile)
			strOutput = outputFile.Readall
			outputFile.Close
			callCmd = Array(result, strOutput, strOutput2)
			deleteInsistently(strTmpFile)
		Else
			failed=true
		End If
	Else
		failed=true
	End If
	
	If failed=true Then
		criticalErrorHandler "callCmd()", "Kommando fehlgeschlagen: " & strCommand _
						& vbCrLf & "--- Fehlermeldung: " & strOutput2, 0, ""
	End If
End Function


'--- deleteInsistently(strFileName)  -----------------------------------------------------
' Auf einigen Testsystemen trat ein Fehler auf, weil die Funktion callCmd() ihre
' temporären Dateien nicht wieder löschen konnte. Vermutlich blockierte gerade ein
' Virenscanner die Datei. Die Funktion deleteInsistently() unternimmt deshalb 10 Versuche,
' die übergebene Datei zu löschen. Wenn ein Versuch fehlschlägt, probiert es das Skript 5
' Sekunden später erneut.
Function deleteInsistently(strFileName)
	Dim noOfTries, successful
	
	On Error Resume Next
	noOfTries=0			
	successful=false
		
	While noOfTries < 10 And Not successful
		Err.Clear
		If fso.FileExists(strFileName) Then
			fso.DeleteFile(strFileName)
				If Err.Number <> 0 Then
					successful=false
					noOfTries = noOfTries + 1
					logAppend("--- Warnung: Konnte temporäre Datei " & strFileName & " nicht löschen, Versuch " & noOfTries)
					Wscript.Sleep(5000)
				Else
					successful=true
				End If
		Else
				successful=true
		End If
	Wend
	On Error Goto 0
	If Not successful Then
		logAppend("--- Warnung: Ich geb's auf.")
	End If
End Function
