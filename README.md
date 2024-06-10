# DB-Matcher-v5
Database (Excel) matching program written in C# with .NET-Framework using nPoi

**Installation**

1. Download INSTALLER.zip
2. Extract the files to the target directory
3. Run the INSTALLER.exe
4. You will be guided through the process of installation by INSTALLER.exe

**Description**

*** English ***

DB-Matcher
Data ---
Programming language: Visual C#
Ide: Visual Studio
Length: ~1600 lines
Interface: CMD / Terminal
Platform: Windows 10/11
Architecture: x64 (64bit)
Primary algorithm: Levenshtein-Distance
Secondary algorithm:	Hamming-Distance
Tertiary algorithm: Jaccard index
Summary ---
Area of application: Merging two databases in Excel format (*.xlsx, *.xls)
Basic function: The DB matcher calculates how similar a value from the primary database is to a value from the secondary database. It follows the following algorithm scheme: 

    start --> LD
    fails --> HD
    fails --> JD

Functions:
- Definition of a separate database (using a wizard in DB-Matcher) in which known differences (e.g. company abbreviations) can be saved. These differences are taken into account during matching. The database created can also be edited externally. 
- Create a backup copy of the Excel file to prevent data loss and increase convenience, as the original file can remain open in Excel without causing data stream conflicts.
- Automatic dynamic determination of databases: The most likely desired table ranges are provided as default values during input. They can either be accepted by pressing ENTER or changed by entering text. The default values determined behave dynamically in relation to the input made and are updated in real time
- Incorrect entries are recognised and reported back to the user.
- The path entered is checked for authenticity
- The DB-Matcher has several setting options, which are saved in a configuration file so that they do not have to be redefined each time it is started
- Possible interrupt sequence when starting DB-Matcher
- The databases do not have to be on the same worksheet. 
- The results can be written from any column to any worksheet.
- For the results, it is possible not only to write the checked values, but also to transfer the associated cells as well
- For security-relevant databases or for registration in other processes, programmes or security software, the checksum (SHA256) is automatically calculated and provided before and after matching. 
- A progress bar is displayed in the bottom line of the console window in combination with a time display. The progress bar is dynamically adjusted to the width of the console window. The time display shows the calculated time that the DB matcher will probably still need for matching.
- The number of array accesses is displayed for diagnostic purposes
- The user is guided through the programme at all times and each entry is labelled

*** German ***

DB-Matcher
Daten ---
Programmiersprache:	Visual C#
Ide:				Visual Studio
Länge:				~1600 Zeilen
Oberfläche:			CMD / Terminal
Plattform:			Windows 10/11
Architektur:			x64 (64bit)
Primärer Algorithmus:	Levenshtein-Distance
Sekundärer Algorithmus:	Hamming-Distance
Tertiärer Algorithmus:	Jaccard-Index
Zusammenfassung ---
Einsatzbereich:	Zusammenführung zweier Datenbanken im Excelformat (*.xlsx, *.xls)
Basisfunktion:		Der DB-Matcher berechnet, wie ähnlich ein Wert aus der primären Datenbank einem Wert aus der sekundären Datenbank ist. Er folgt dabei dem folgenden Algorithmus-Schema: 

    start    --> LD
    fails    --> HD
    fails    --> JD


Funktionen:
-	Definierung einer eigenen Datenbank (mittels Wizard in DB-Matcher), in welcher bekannte Unterschiede (z.B. Firmenkürzel) gespeichert werden können. Diese Unterschiede werden beim Matching berücksichtigt. Die erzeugte Datenbank kann auch extern bearbeitet werden. 
-	Erstellen einer Sicherheitskopie der Excel-Datei um Datenverluste zu verhindern und den Komfort zu erhöhen, da die Ur-Datei weiter in Excel geöffnet bleiben kann, ohne dass es zu Datenstreamkonflikten kommt.
-	Automatische dynamische Ermittlung der Datenbanken: Die wahrscheinlich gewünschten Tabellenbereiche werden als default-Wert bei der Eingabe bereitgestellt. Sie können entweder mit ENTER akzeptiert oder per Texteingabe geändert werden. Die ermittelten default-Werte verhalten sich dynamisch zu der geleisteten Eingabe und werden ich Echtzeit aktualisiert
-	Fehlerhafte Eingaben werden erkannt und an den Nutzer rückgemeldet.
-	Der eingegebene Pfad wird auf Echtheit überprüft
-	Der DB-Matcher verfügt über mehrere Einstellungsmöglichkeiten, welche in einer Konfigurationsdatei  gespeichert werden, sodass sie nicht bei jedem Start neu definiert werden müssen
-	Mögliche Interrupt-Sequenz beim Starten von DB-Matcher
-	Die Datenbanken müssen sich nicht auf dem gleichen Arbeitsblatt befinden. 
-	Die Resultate können ab einer beliebigen Spalte in ein beliebiges Arbeitsblatt geschrieben werden.
-	Bei den Resultaten besteht die Möglichkeit, nicht nur die geprüften Werte zu schreiben, sondern zusätzlich die dazugehörenden Zellen mitzuübertragen
-	Für Sicherheitsrelevante Datenbanken oder zur Registrierung in andere Prozesse, Programme oder Sicherheitssoftware wird automatisch die Prüfsumme (SHA256) vor und nach dem Matching berechnet und bereitgestellt. 
-	Ein Fortschrittsbalken wird in der untersten Zeile des Konsolenfensters in Kombination mit einer Zeitanzeige angezeigt. Der Fortschrittsbalken wird dynamisch an die Breite des Konsolenfensters angepasst. Die Zeitanzeige zeigt die berechnete Zeit, die der DB-Matcher vorraussichtlich noch für das Matching benötigt.
-	Zu Diagnosezwecken wird die Anzahl an Array-zugriffen angezeigt
-	Der Benutzer wird zu jeder Zeit durch das Programm geführt und jede Eingabe ist betitelt

**Donations**

    Monero (XMR): 439avs1Cp5gcwWnxTqsicoJcvX5SiK69TdhumoXgULzVXYp94PJLbobAzKUPA9GkSqBdXP6cgRb4dEpSEGAgdUkTHjcVsaG

    Polygon (MATIC): 0x8AA598780c6529DCB771B9455E2717043BB1a1e1

    Solana (SOL): 21BpX9xhkitMEWx2iBxGXEVsFCMF9huViud24NbrPqNC

    Ethereum (ETH): 0x8AA598780c6529DCB771B9455E2717043BB1a1e1

    Bitcoin (BTC): bc1q2ss7n5gv3tr68dfqtmcvy830pj4clf7wqxztk5


**Hold**


    The DB-Matcher is an easy-to-use console application written in C# and based on the .NET framework. 
    The DB-Matcher can merge two databases in Excel format (*.xlsx, *.xls). 
    It follows the following algorithms in order of importance: Levenshtein distance, Hamming distance, Jaccard index. 
    The DB-Matcher takes you by the hand at all times and guides you through the data matching process    

    Copyright (C) 2024  Carl Öttinger (Carl Oettinger)

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU Affero General Public License as published
    by the Free Software Foundation, either version 3 of the License, or
    any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU Affero General Public License for more details.

    You should have received a copy of the GNU Affero General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.

    You can contact me in the following ways:
        EMail: oettinger.carl@web.de, big-programming@web.de
