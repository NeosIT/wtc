# WTC - Word Template Corrector

## Deutsch
### Was ist WTC?
Erzeugt man Dokumente mit Microsoft Word auf Basis von Vorlagen, so wird im
Word-Dokument auch der Pfad die dieser Vorlage gespeichert. In Unternehmen
wird oft eine gemeinsame Sammlung von Vorlagen, die auf einem Server liegt,
genutzt. So weit so gut.

### Das Problem
Ändert sich nun der Pfad zum Vorlagenverzeichnis wird die Sache unangenehm.
Insbesondere die Änderung der Serverangabe ist problematisch. Zum besseren 
Verständnis hier mal ein Beispiel:

* Alter Pfad zu den Vorlagen: \\\\alter-server\\share\\templates\\
* Neuer Pfad zu den Vorlagen: \\\\neuer-server\\share\\templates\\
* Der alte Server existiert nicht mehr.

Öffnet man nun ein Word-Dokument, dass mit einer Vorlage vom alten Server
erzeugt wurde, versucht Word die Vorlage von dort zu laden. Da der Server
aber gar nicht mehr existiert, versucht Word das so lange, bis es in einen
Timeout läuft. Und das kann leider gefühlt, sehr lange dauern.

### Die Lösung
Die Lösung ist naheliegend. In den betroffenen Dokumenten muss **nur** der
Pfad zur Dokumentenvorlage geändert werden. Dazu findet man im Netz viele
Lösungen, die aber fast alle eines gemein haben:
* basieren auf VBA und laufen innerhalb von Word oder Excel,
* sind langsam,
* verarbeiten keine Verzeichnisbäume
* geben keine vernünftigen Fehlermeldungen aus,
* können nicht adäquat von der Kommandozeile genutzt werden.

WTC macht das anders, was allerdings zu einer Einschränkung führt. **WTC
funktioniert nur für Dokumente im Format Office Open XML, das mit Office 2003
eingeführt wurde. Die gebräuchlichen Dateiendungen sind .docx, dotx, docm und
dotm.** Diese Dateien sind im Grunde eine Sammlung von
XML-Dateien, die in ein ZIP-Archiv verpackt sind. Ja tatsächlich. Man kann die
Dokumente ganz einfach mit einem Programm wie 7Zip öffnen oder entpacken.

### So arbeitet WTC
Das Prinzip ist simpel. Datei für Datei wird entpackt und dann wird in der
Einstellungsdatei `word\_rels\settings.xml.rels` wird die alte und unerwünschte
Pfadangabe zur Vorlage durch die neue, korrekte ersetzt. Danach wird alles
wieder eingepackt und die Originaldatei ersetzt. Standardmäßig wird dabei eine
Sicherungskopie des Originals mit der Endung .bak erzeugt.

### Die Optionen
* `--help`  
Ausgabe der möglichen Optionen
* `-d, --directory` (required)  
Alle Dokumente in diesem Verzeichnis werden bearbeitet.
* `-o, --old` (required)  
Pfad oder Teil des Pfades zur Vorlage bzw. dem Vorlagenverzeichnis der ersetzt werden soll.
* `-n, --new` (required)  
Durch diesen Pfad oder Teilpfad wird ersetzt.
* `-r, recursive` (optional)  
Bearbeite auch die Dokumente in den Unterverzeichnissen.
* `-b, --nobackup` (optional)  
Es werden **keine** Sicherungsdateien für korrigierte Dokumente angelegt.
* `-t, --dry-run` (optional)  
Testmodus benutzen. Es wird nach zu ändernden Dokumenten gesucht, sie werden aber nicht verändert.
* `-v, --verbose` (optional)  
Es werden ausführlichere Fehlermeldung ausgegeben.

#### Beispiel
`wtc -d \\server\share\documents -o \\alter-server\share\templates\ -n \\server\share\templates\ -r`


## English

Translation following soon.