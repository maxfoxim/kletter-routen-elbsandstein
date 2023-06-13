# kletter-routen-elbsandstein
Auswertung und Visualiserung der Daten von Teufelsturm.de. Ursprünglich kam mir die Idee, als ich auf die Bewertungen und Kommentare in der sächsischen Schweiz zugreifen wollte. Wegen schlechtem bis gar keinem Handyempfang war dies leider nicht möglich. Also ist die Idee entstanden die ganzen Informationen offline auf dem Handy zur Verfügung zur stellen.

# Was kann es:
- Die Kommentare von Teufelsturm als PDF speichern.
- Eine grafische Darstellung von Schwierigkeit gegen Bewertungsqualität.
- Die Routen der 3 kürzesten Parkplätze.
- Die PDFs in er KMZ-Datei speichern, sodass diese in LocusMap (eventuell auch andere Offline-Karten-Apps) auf Karten dargestellt werden können. Hat den Vorteil, dass man direkt im Umkreis die passenden Dokumente hat.

# Welche Dateien sind relevant?
- Hauptversion_Elbi.xlsx: enthält eine Liste aller Gebiete, Gipfel und Wege und können beispielsweise genutzt werden um die "besten" Gipfel zu finden.
- Routen_Distanz.py: berechnet die Routen von 3 angegebenen Parkplätzen zum Gipfel
- Tabelle_To_PDF.py: erstellt aus den Kommentaren und den Bewertungen der Gipfel und Routen ein PDF inklusive Bewertungsdiagramm und Zustiegsskizze
- Scatterplot Ordner: Zwischenspeicher für alle Qualitäts-Schwierigkeits-Diagramme
- maps Ordner: HTML und PNG der Routen
- Icons Ordner: die Icons die für die KMZ Datei zur Visualiserung der Gipfel benutzt werden
- PDF_Ausgabe Ordner: Ausgabe der erstellten PDFs

# Wie bekomme ich ein PDF mit Karte und Kommentaren?
Nachdem alle Pythonpakete installiert sind, Chromedriver vorhanden ist und der API-Token vorhanden ist geht es nun wie folgt weiter:
 
1. In der Datei Routen_Distanz.py müssen die jeweiligen gewünschten Gebiete ausgewählt werden, welche eine Parkplatzroutenberechnung erhalten sollen. Die Startparkplätze stehen in der Exceldatei und müssen dabei folgende Formatierung haben: NAME_PARKPLATZ (Längengrad;Breitengrad), z.B. Neumannmühle (50.9236588, 14.2845213). Der Parkplatz muss, wie beispielsweise im Zschand, für 3 Parkplätze erfolgen. Anschließend wird vom Skript die Route zwischen diesen Parkplätzen und den Gipfeln berechnet und die Ergebnisse als JSON gespeichert. Die JSON-Datei wird dabei auf Länge, Dauer und Streckeninformationen ausgelesen und die Werte in die Exceltabelle zurückgeschrieben. Die Routeninformation wird dann grafisch in eine HTML eingetragen mit OSM als Untergrund. Um aus der HTML ein PNG zu erstellen muss (leider) jedesmal der Browser automatisch geöffnet werden und ein Screenshot erstellt werden. Die gespeicherten Bilder stehen anschließend zur Verfügung.
2. Im Skript Tabelle_To_PDF.py wird nun wieder das gewünschte Gebiet ausgewählt und anschließend das Skript gestartet. Hier werden nun die Kommentare von Teufelsturm mittes Webcrapper ausgelesen und in einem PDF gespeichert. Die einzelnen Schwierigkeitsgrade und Qualitätsbewertungen werden dabei zusätzlich noch als Scatterplot ausgegeben. 

# Welche Vorbereitung brauche ich:
Die Pythonpakete für die Durchführung der beiden Skripte befinden sich in der Datei requirements.txt. 

Zusätzlich dazu wird der aktuelle chromedriver für Chrome benötigt. Mit anderen Browser lässt es sich sicherlich auch durchführen, aber mit Chrome hats immer geklappt. Link: https://chromedriver.chromium.org/downloads.

Für die Berechnung der Routen wird ein API-Token von openrouteservice (https://openrouteservice.org/dev/#/home) benötigt. Dieser kann nach kurzer Anmeldung gratis erworben werden und muss anschließend in der Datei secret_api.py hinterlegt werden.



#Sonstige Anmerkungen
Die Daten stammen hauptsächlich von Teufelsturm.de. 
