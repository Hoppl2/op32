Attribute VB_Name = "modWinRezX"
'5.3.00 230824  AE  Lieferengpass für Muster16
'5.1.42 210712  AE  'DatumObenRezeptVersatz' eingebaut
'5.1.32 201207  AE  CheckAutIdem: Prüfung jetzt auf AusnahmeErsetzungFl=1 wegen neuem Wert 'bedingt'
'5.1.23 200707  AE  Einbau 'Patientenalter' in A+V Taxierung
'5.1.14 190903  AE  Public BtmDerivatMengeIgnorieren% eingebaut; beigetretene Vereinbarungen (jetzt in GetPrivateProfileIni mit 2000 Stellen)
'5.1.1  190516  AE  'WinArtDebug' wieder aktiviert durch Ini-Eintrag
'5.1.0  180914  AE  Einbau 'Stützstrumpf-Problematik': ZusatzkomponentenNr, ZusatzKomponentenBasisNr,ZusatzKomponentenFaktor
'5.0.25 180723  AE  SucheInGruppeMDB: jetzt immer 'PauschaleNr' in SQL-String, denn wenn keine Pauschale gewünscht ist, dürfen auch nur die Sätze ohne PauschaleNr genommen werden!'5.0.21 180423  AE  IVF-Rezepte: Berücksichtigung Festbetrag in RezeptHolenDB, RezeptHolenMDB
'5.0.20 180420  AE  Neue Bedruckung: mit Mehrkosten (werden intern in Zusatz(1) gespeichert, Verwendung in InitRezeptDruck): jetzt nur wenn in verkaufsdatei als Hilfsmittel gesetzt und Mehrkosten>0
'5.0.8 10.04.18     AE  Neue Bedruckung: mit Mehrkosten (werden intern in Zusatz(1) gespeichert, Verwendung in InitRezeptDruck)
'5.0.7  170531      AE  InitRezeptDruck: bei Verschieben Noctu (2567018) Berücksichtigung der Verfügbarkeiten
'5.0.1  14.03.17    AE  Code128: in InitRezeptDruck+WriteRezeptSpeicher an WinRezK angepasst
'5.0.0      170222     Einbau Kunden als SQL; Adaption Ruekkauf.cls für 8-stellige PZN
'4.0.89 07.02.17    AE  Code128: jetzt unabhängig von AVP und FiveRx
'4.0.88 03.02.17    AE  Code128: bisher musste FiveRx vorhanden sein, jetzt nur das erste If mit AvpTeilnahme oder FiveRx
'4.0.82 30.11.16    AE  InitRezeptDruck: wenn Noctu vorhanden, dann ans Ende der RezArtikel
'4.0.81 14.11.16    AE  Einbau Code128: Menü,Ini,Druck
'4.0.75 30.06.16    AE  ApvParse: bei AMPV_Alt jetzt auch korrekt abziehen: erg# = CalcAMPV(ActAep# * ActMenge, ActMwst%) / (1# + (ActMwst%) / 100#) - (ActAep# * ActMenge)      bisher nur '- ActAep#
'4.0.74 21.06.16    AE  Ab 1.7. neue Berechnungsreihenfolge für AbrechnungsKz=2: jetzt zuerst mit Faktor multiplizieren, dann MwSt,Rabatt; dafür .VdpPauschale zweckentfremdet; Adaption InitRezeptDruck; zum Test Aufruf mit 'ABR2'
'4.0.72 20.05.16    AE  .Fortsetzung eingebaut (wegen Übersteuerung Preis,Zuzahlung in OpKasse)
'4.0.71 15.04.16    AE  AplusV überarbeitet (laut Fehlerprotokolle)
'4.0.70 05.04.16    AE  AplusV überarbeitet (APVparse eingebaut wegen neuer Inhalte: RUNDE,MWST,EURO,AMPREISV)
'4.0.66 11.02.16    AE  RezeptHolenMDB: beim Zusammenfassen von Artikeln wird bei HimiVerbrauch die Zuzahlung jetzt nicht mhr addiert
'4.0.65 11.02.16    AE  RezeptDruck: vor Ausdruck von Taxierungen jetzt auch Prüfung des Parameters 'RezepturDruck'
'4.0.62 14.01.16    AE  Handling §302: in VerbandmittelMDB die Prüfung auf 'AbrechnungsverfahrenKz=2' auch vor die Erzeugung/Prüfung auf HmNummer,2.Zeile (Übersteuerung Variable Paragraph302)
'4.0.61 18.12.15    AE  Prüung auf 32000 in RezeptHolenDB entfernt
'4.0.60 04.12.15    AE  Einbau PreisKz '74' auch für "09999146" (indiv. hergest. Schmerzlösung) für HashCode
'4.0.59 19.11.15    AE  Handling §302: jetzt auch Berücksichtigung 'PZN' in AbrechNr in Berechnung (.tkkpzndruck)'4.0.58 16.11.15    AE  Einbau PreisKz '74' auch für "09999169" (indiv. hergest. parent. virustatikahaltige Infusionslösungen) für HashCode
'4.0.57 05.11.15    AE  Handling §302: jetzt mittels AbrechnungsVerfahrenKz=2 in VdbBedingungen; 'ABR302' in VdbVOKG nicht mehr vorhanden
'4.0.55 28.10.15    AE  Einbau PreisKz '74' (Zytostatika) für HashCode
'4.0.54 01.10.15    AE  Wenn IVF-Rezept, dann diese Sonder-PZN immer in 1.Zeile; Einbau AutIdemDB; Einbau §302 Stückpreis ...
'4.0.52 21.07.15    AE  PrivRezept als BtmRezept: in RezeptHolenDB,RezeptHolenMdb setzen BtmRezept auch bei PrivatRezept
'5.1.24 31.07.20    AE  AMPV_NEU: jetzt 8.56 statt 8.51
'4.0.51 05.07.15    AE  UmspeichernSpezialitaet: bei FAM=Substanz (ttyp=mag_anteilig) Taxe-EK nehmen falls > 0
'4.0.50 30.04.15    AE  "PlusMehrKosten inaktiviert in RezeptHolenDB" - wieder aktiviert; SonderPzns für Verfügbarkeit und Armin getauscht
'4.0.48 08.03.15    AE  BTM: Ausdruck 4.Zeile (@BTM2@) ermöglicht, dafür wird der BTM-ApoName abgeschnitten; Wirkstoffverordnung (jetzt Feld 'Wirkstoffverordnung' anstatt bisher 'WVORabattArtikel';
'4.0.46 14.01.15    AE  Prüfung auf kunbezug.mdb adaptiert
'4.0.45 12.01.15    AE  Abfrage nach NEUEM Btm-Rezept-Formular entfernt, jetzt immer NEUES Rezeptformular
'4.0.43 18.10.14    AE  Druckeranpassungen
'4.0.42 10.10.14    AE  PruefeRezkontrDat: TmHeader.ActMenge setzen, ansonsten Preise für Arb, Emb, ... falsch
'4.0.41 09.10.14    AE  HiMiVerbrauch: immer IstWg4 (CheckZuzZeilenWert), RezeptHolenZuz*Faktor in CalcZuzaNeu
'4.0.37 30.07.14    AE  PlusMehrKosten inaktiviert in RezeptHolenDB,RezeptHolenMDB
'4.0.35 13.07.14    AE  Einbau ARMIN WVO....
'4.0.34 24.04.14    AE  Handelt es sich um eine NEUES Btm..', jetzt DefaultButton1
'4.0.33 27.03.14    AE  Einbau VebNr,PauschaleNr auf Artikelebene: in RezeptHolenMDB war irrtümlich 'VerkaufAdoRec'
'4.0.32 25.03.14    AE  NeueTaxierung: EinlesenPassendeArbEmb% jetzt (NAME=A.NAME) statt bisher (ID=A.ID)
'4.0.31 20.03.14    AE  Einbau VebNr,PauschaleNr auf Artikelebene
'4.0.29 19.03.14    AE  Bei Pauschalen: Prüfung auf alle mit gleichen ersten 7 Stellen der HM-Nummer, dann Zuz nur für 1.
'4.0.25 14.03.14    AE  Einbau Druck der mag.Taxierungen
'4.0.24 01.12.13    AE  Umbau Speichern Rezeptspeicher (in Struct Warenzeichen+N0 aktiviert; neue Felder in Tabellen: Rezepte+Artikel);
'4.0.23 18.11.13    AE  Berücksichtigung Wunschartikel, Neues-BTM-Formular (dafür iMsgBox systemmodal)
'4.0.22 18.11.13    AE  Berücksichtigung 'RezeptNrVersatzY' auch bei BTM-Rezepten
'4.0.21 14.11.13    AE  Berücksichtigung 'AutIdemKreuz' bei Ausdruck
'4.0.20 13.11.13    AE  Einbau '.VdbPauschale', damit für Pauschal-Artikel der Faktor keinen Einfluss auf die Ermittlung RezSumme,RezGebSumme, ... nimmt
'                   Uhrzeitformat bei Noctu-Rezepten; Bei Selbsterklärung keine Avp('X')-Zeile oben
'4.0.18 26.09.13    AE  Wegfall 'Handelt es sich ...' bei BTM-Rezepten
'4.0.17 09.09.13    AE  AMPV_NEU: jetzt 8.51 statt 8.35; Noctu bei Prüfung auf Sonder-PZN ausnehmen
'4.0.8  08.05.13    AE  Überarbeitung zus.Hilfsmittel/WertermittlungKz=2; beigetretene Vereinbarungen (jetzt in GetPrivateProfileIni mit 1000 Stellen)
'4.0.1  11.04.13    AE  Berücksichtigung 'Zuzahlung' in Verbandmittel
'4.0.0  09.04.13    AE  Wawi-SQL und Verkauf-SQL
'3.0.17 25.03.13    AE  Einbau neues A+V Handling
'2.0.53 05.09.12    AE  alle PZN-Felder in DBs als long; SqlOp in DLL ausgelagert;
'2.0.52 xx.03.12    AE  Progstart jetzt wieder mit Main gegen Automatisierungsfehler
'2.0.51 06.03.12    AE  RezeptNrPositionAlt eingebaut: default N; damit kann RezeptNr wieder wie früher (unten) gedruckt werden
'2.0.48 18.11.11    AE  Ausdruck adaptiert (gegen irrtümliches nochmal-drucken des letzten Artikels an Pos BenutzerNr)
'2.0.45 29.10.11    AE  RezeptHolenMDB: bei Abholer und Preis=0: wenn FB dann nehmen, analog zu WinRezK
'2.0.44 26.10.11    AE  Einbau RezeptNrVersatzY
'2.0.40 21.10.11    AE  Ausdruck RezeptNr wieder ca. 1.5 mm nach unten
'                       Löschen Rezept aus Rezspeicher, wenn vorhanden; damit immer das gedruckte speichern
'2.0.39 19.10.11    AE  Adaption Rabatte: in RabWerte nur noch Kassenrabatt,neues Feld für GhRabatte
'                       Ausdruck RezeptNr ca. 5 mm nach oben
'2.0.37 04.10.11    AE  Ausdruck Privatrezepte im Hochformat: jetzt auch Ik,Datum,ApoName drucken
'2.0.36 02.10.11    AE  Druck der Rezeptnr jetzt an der Stelle, wo bei BTm-Rezepten der ApoName
'                       Ausdruck Privatrezepte im Hochformat: Aufrufpara ' HOCH'; PrivatRezeptVersatzY
'2.0.32 14.08.11    AE  Preiseingaben ohne Zuzahlung (für beide Kassen)
'2.0.25 09.05.11    AE  Noctu-ShowNichtInTaxe: bei der Anzeige 'nicht in Taxe ...' ausgenommen; LadeSonderPzn: nur wenn befüllte Zeile, dann berücksichtigen
'2.0.24 02.05.11    AE  PlusMehrkosten in allen Varianten vereinheitlicht (FB ...)
'                       Speicherung Privatrezept in RezSpeicher adaptiert
'2.0.23 27.04.11    AE  Überarbeitung 'PlusMehrkosten': jetzt Prüfung, ob der Artikel überhaupt einen FB hat; wenn nicht, Feld rücksetzen
'2.0.22 01.04.11    AE  PZN für Noctu wieder auf 2567018 geändert
'2.0.20 10.03.11    AE  1010-767BA: BTM-Rezepte - Gebühr
'2.0.19 04.02.11    AE  PZN für Noctu auf 2567024
'2.0.15 03.02.11    AE  Einbau 'Noctu' für Ausdruck: wenn PZN 2567018 in Rezept, dann auch Ausdruck der Abgabe-Uhrzeit
'2.0.14 27.01.11    AE  Einbau des direkten Rezeptdrucks aus der Kasse (direkt nach Rezeptende): 'Grüne Rezepte' eingebaut
'2.0.13 22.01.11    AE  Einbau des direkten Rezeptdrucks aus der Kasse (direkt nach Rezeptende): Privatrezepte eingebaut
'2.0.12 13.01.11    AE  Einbau des direkten Rezeptdrucks aus der Kasse (direkt nach Rezeptende)
'2.0.11 18.09.10    AE  RezeptDruck: ab jetzt auch Faktor drucken, wenn =1
'2.0.3  01.08.10    AE  RezeptHolenMDB: wenn Artikel nur im Stamm (nicht in Taxe) und Wg=3, dann als Rezeptur mit Zuzahlung - Erweitert auch für PRIVATREZEPTE
'2.0.2  14.07.10    AE  RezeptHolenMDB: wenn Artikel nur im Stamm (nicht in Taxe) und Wg=3, dann als Rezeptur mit Zuzahlung
'2.0.1  11.05.10    AE  Debuggings eingebaut zum Auffinden der Zeitverzögerung; gefunden, deshalb Prüfung auf Vorhandensein von DB-Feldern nur wenn App.ExeName <> "WINREZDR"
'2.0.0  15.04.10    AE  NBNEU-Win7: neue DLL, SendKeys ersetzt durch Eigenumsetzung
'                       P302 eingebaut für Wiederaufruf von §302-Rezepten
'1.1.32 25.05.09    AE  Anpassung an neues Fehlerhandling
'1.1.31 27.04.09    AE  KundenNr jetzt auch über 32000 ohne Fehler (LONG!!)
'1.1.30 07.04.09    AE  Privatrezept-Besorger-Stückelung: in diesem Fall nicht Aconto+Restzahlung, sondern den Preis aud Verkauf.mdb
'1.1.28 19.03.09    AE  Für HimiVerbrauch die VerkaufRec!ZuzaGes nehmen (neu in verkauf.mdb: wegen Rundungsumgenauigkeiten ansonsten)
'1.1.27 17.03.09    AE  Mitkompiliert mit WinRezK
'1.1.26 16.03.09    AE  Mitkompiliert mit WinRezK
'1.1.25 12.02.09    AE  Mietgebühr: in Struktur neue Komp: MietDauer; wenn gesetzt, dann auf Ausdruck SonderPZN und Dauer als Faktor
'1.1.23 12.01.09    AE  WriteRezeptSpeicher auch für Computer>=10 ohne Fehler!
'1.1.22 17.11.08    AE  Mitkompiliert mit WinRezK
'1.1.21 16.09.08    AE  MachEinzelBerechnung: Rundung des Ergebnisses mittels fnx
'                       MachBerechnung: wenn False, dann Zuz auf 0 setzen
'1.1.20 17.06.08    AE  Wenn Sonderpreisartikel, dann Feld 'VK' aus Verkauf.mdb nehmen (mit Berücksichtigung FB)
'1.1.19 15.05.08    AE  Umbau für Zuzahlungserlass - in RezeptHolen, RezeptHolenMDB, NeuerArtikel, CalcTaxeZuzahlungDOS (damit reduzierte Zuz vor Prüfung auf Rausfall aus Rezept)
'1.1.18 10.04.08    AE  RezeptHolenMDB: Anpassung an Privatrezepte - Übergabe richtig ausgewertet (jetzt kommt Id von Kasse)
'1.1.17 12.03.08    AE  RezeptHolenMDB eingebaut für FlexKasse
'1.1.15 02.03.08    AE  Tabelle Rezepte: neues Feld 'kKassenIk', für AVP-Connect
'1.1.14 17.02.08    AE  MachBerechnung: Klammerung bei FB/2*ST erzeugte FB/(2*ST) - Schwachsinn - deshalb vorab Prüfung on '+' oder '-' überhaupt in Formel vorhanden
'1.1.10 22.01.08    AE  MachBerechnung überarbeitet: auch für Formeln der Art (MENGE*EK+AMPV_ALT)/MENGE*0,95; auch für EK-15%
'1.1.08 10.01.08    AE  MachBerechnung überarbeitet: da es jetzt EK+ST*0,4 gibt (ohne Klammerung!), wird die Klammerung vom Prog gesetzt
'1.1.07 21.12.07    AE  IVF-Rezepte: taxbetrag jetzt Avp/2, Zuz 0. Als erste Zeile PZN 9999643
'1.1.05 06.12.07    AE  Einbau IVF-Rezepte; Rundung bei A+V adaptiert (clng weg!)
'1.1.04 04.11.07    AE Mitkompiliert mit WinRezK
'1.1.03 16.08.07    AE Anpassung HoleKundenInfo an neue Kundenstammdaten (kunden.mdb)
'1.1.02 13.08.07    AE Mitkompiliert mit WinRezK
'1.1.00 25.07.07    AE Mitkompiliert mit WinRezK
'1.0.99 19.07.07    AE Mitkompiliert mit WinRezK
'1.0.98 03.06.07    AE Mitkompiliert mit WinRezK
'1.0.97 01.06.07    AE Mitkompiliert mit WinRezK
'1.0.96 25.04.07    AE Mitkompiliert mit WinRezK
'1.0.96 25.04.07    AE Mitkompiliert mit WinRezK
'1.0.95 20.04.07    AE  ZuzahlungsErlass eingebaut
'1.0.94 05.04.07    AE  OP-AutIdem wieder aktiviert
'1.0.93 03.04.07    AE  Anpassung an DSK: GS übergibt jetzt uU IK-Nummer
'1.0.92 01.04.07    AE  Anpassung an AutIdem-Sonderregeln: analog WinRezK
'1.0.91 26.03.07    AE  Anpassung an AutIdem-Sonderregeln: analog WinRezK
'1.0.88 27.11.06    AE  Druck Avp-RezeptNr: jetzt 12 cpi
'1.0.87 23.11.06    AE  Avp-Rezept analog WinRezK
'1.0.86 13.11.06    AE Mitkompiliert mit WinRezK
'1.0.85 29.10.06    AE  HiMiVerbrauch: bei Zusammenfassen mehrerer untereinander stehender jetzt Zuzahlung ok
'                       Druck Taxierungen: wenn eine Taxierung da mit Preis>0 (also über RezGeb), dann Ausdruck quer
'1.0.84 22.08.06    AE BTM-Rezepte: in 4.Zeile BTM-Rezept-Text (Extras/Optionen); nur mehr max 3 PZN bei BTM-Rezept
'1.0.83 31.07.06    AE Mitkompiliert mit WinRezK
'1.0.82 28.06.06    AE Mitkompiliert mit WinRezK
'1.0.81 15.06.06    AE Mitkompiliert mit WinRezK;  Überarbeitung PlusMehrKosten in RezeptHolen (jetzt Zuz von Festbetrag)
'1.0.80 12.06.06    AE Mitkompiliert mit WinRezK
'1.0.79 05.05.06    AE 'PlusMehrkosten' eingeführt: wenn in DSK Storno+F8, dann zahlt Krankenkasse die Mehrkosten!
'1.0.78 24.04.06    AE Mitkompiliert mit WinRezK
'1.0.77 08.01.06    AE Mitkompiliert mit WinRezK
'1.0.76 22.10.05    AE Überarbeitung HimiZumVerbrauch: Deckelung von DSK nehmen
'1.0.75 28.10.05    AE Mitkompiliert mit WinRezK
'1.0.74 22.08.05    AE Privatrezepte - Besorger - Stückelung (in RezeptHolen; BesorgerPreis nur wenn nicht mit ( beginnt
'1.0.73 07.08.05    AE Mitkompiliert mit WinRezK
'1.0.72 27.07.05    AE Einbau DatumVersatzX
'1.0.71 11.07.05    AE Mitkompiliert mit WinRezK
'1.0.70 27.06.05    AE neue op32.dll
'1.0.69 21.06.05    AE Handling Privatbesorger adaptiert: jetzt Preis von Aconto-Zeile nehmen; uU Multiplikator in Zeile danach beachten
'1.0.67 08.06.05    AE Anpassung an neue A+V Strukturen
'1.0.66 29.05.05    AE SonderBelegRezept eingeführt; auf sF3 jetzt Privatrezept-NEU; Barverkäufe auf Strg+B
'1.0.65 31.03.05 AE analog WAMA
'1.0.64 16.03.05 AE mitkompiliert mit winrezk
'1.0.62 20.01.05    AE Neues Handling für Zuzahlungsberechnung auf Zeilenwert-Basis: .istwg4 jetzt nicht mehr von WG, sondern von (HilfsmittelKz und NICHT VerbandKz) abhängig!
'                           und das auch nur, wenn feld 'ZuzZeilenwertKz' noch nicht vorhanden
'                           CalctaxeZuzahlungDOS: Handling mit Erstattungsfähig, BedErstattungsFähig geändert!
'1.0.61 27.12.04    AE Adaption Ausdruck für Berlin+RVO
'       16.12.04    GS VK.AVP nur übernehmen, wenn "A" (nicht, wenn "a")
'                   Druck Berlin+RVO: Funktion "MengeErmitteln" (6x20 war falsch)
'1.0.60 16.12.04    AE  spez. Ausdruck für Berlin+RVO
'                       BTM-rezepte: Datum wieder an urspr. Stelle
'1 0.58  06.12.04 AE ParseOperand: bisher wurde AMPV... immer fix vom AEP genommen; jetzt Berücksichtigung des etwaiigen Multiplikators MENGE
'1.0.57  03.12.04 AE,GS A+V: Klammern parsen korrigiert, Bei Suche nach "VK" vor MwSt: Instr statt Left(,2)
'1.0.56  01.12.04 AE  Anpassung an WinApo-CD (wg 3 wegtun!)
'1.0.55  25.11.04 AE  Neue op32.dll
'1.0.54  11.10.04 AE  Positionierung Abgabedatum NEU für Windows-Drucker
'1.0.19  17.03.04 AE  Verbandmittel: vor Ermittlung Zuza nur dann mit Menge multiplizieren, wenn IstWg4 gesetzt ist
'1.0.18   3.02.04 GS  AMPV_NEU: 8.1 statt 8.13
'1.0.17  28.01.04 AE  Diätetika (Taxe-WG 3): auch Zeilenwert nehmen; dafür IstWg4 verwendet
'1.0.16  27.01.04 AE  Verbandmittel: Für Berechnung der Zuzahlung den A+V Rabattfaktor berücksichtigen; auf Zeilenwert-Basis umgestellt
'1.0.15  19.01.04 AE  CalcTaxeZuzahlungDOS: bei RezeptHolen den Preis aus VK mitgeben und daraus Zuz berechnen (wegen Preisänderung taxe!); auch bei SonderPznArtikeln im Rezept
'1.0.14  14.01.04 AE  RezeptHolen: Zuzahlung bei Kontroll-Stop, Preiseingaben ROT und BLINK
'1.0.13  13.01.04 AE  RezeptHolen: PZN 999999 mit Preis>0 (Preiseingaben) auch Zuzahlung berechnen
'1.0.12  02.01.04 AE  Verbandmittel: Prüfung ob Artikel aus dem Rezept fällt, adaptiert (jetzt Einzelpreis/Packung statt MultiPreis)
'1.0.11  30.12.03 AE  CalcZuzaNeu: wenn HimiVerbrauch, dann keine Untergrenze Zuzahlung
'1.0.10  21.12.03 AE  Jahr2004 - Adaption an WinrezK
'1.0.9   19.12.03 AE  Jahr2004
'1.0.8   21.11.03 AE  DAO351 durch DAO360 ersetzt; CreateDatabase mit Default-Version
'1.0.7   30.10.03 AE  sicherheitshalber mitkompiliert
'1.0.6   06.10.03 AE  Fehlerprot: mehrere PrivRezepte; wenn wo Abholer dabei, wurde Artikel nicht angezeigt
'1.0.5   21.08.03 AE  Verbandmittel: MachBerechnung: neue Werte für Rabatt !
'1.0.4   19.08.03 AE  Verbandmittel: auch bei Kostenvoranschlag od. Genehmigungspflicht Preis lt. Berechnung erzeugen
'                     ('erg=FALSE' unter Kommentar gesetzt)
'1.0.3   04.08.03 AE  wegen Komibidrucker: Reset nach Ausdruck

Option Explicit


Public ProgrammChar$

Public ProgrammNamen$(1)
Public ProgrammTyp%

Global buf$

'Public kKassenDB As Database
'Public KkassenRec As Recordset
Public kKassenDB As clsKkassenDB
Public KkassenRec As New ADODB.Recordset

Public VerkaufDB As Database
Public VerkaufRec As Recordset
Public VerkaufAdoDB As clsVerkaufDB
Public VerkaufAdoRec As New ADODB.Recordset
Public VerkaufDbOk%

Public ProgrammModus%
Public AutIdemOk%, AutIdemSonderregelOk%, kKassenOk%

Public FabsErrf%
Public FabsRecno&

Public UserSection$

Public KeinRowColChange%

Public AvVereinbarungen$()
Public AvPauschalen$()


Public ast As clsStamm
Public ass As clsStatistik
'Public taxe As clsTaxe
Public kiste As clsKiste
Public para As clsOpPara
Public wpara As clsWinPara
Public vk As clsVerkauf
Public RezTab As clsVerkRtab
Public VmPzn As clsVmPzn
Public VmBed As clsVmBed
Public VmRech As clsVmRech

Public sqlop As clsSqlTools

'Public TaxeDB As Database
'Public TaxeRec As Recordset
Public taxeAdoDB As clsTaxeAdoDB
Public TaxeRec As New ADODB.Recordset
Public TaxeAdoDBok%

Public Artikel As clsArtikelDB
'Public ArtikelConn As New ADODB.Connection
'Public ArtikelComm As New ADODB.Command
Public ArtikelAdoRec As New ADODB.Recordset
Public ArtikelInfoRec As New ADODB.Recordset
Public AbgabenRec As New ADODB.Recordset
Public LieferungenRec As New ADODB.Recordset
'Public MerkzettelRec As New adodb.Recordset
'Public LieferantenzusatzRec As New ADODB.Recordset
'Public AusnahmenRec As New ADODB.Recordset
'Public RabattTabelleRec As New ADODB.Recordset
'Public BmEkTabelleRec As New ADODB.Recordset
Public ArtikelDbOk%
    
Public hTaxe As clsHilfsTaxe
Public Hilfstaxe As clsHilfstaxeDB
Public HilfstaxeRec As New ADODB.Recordset

Public AplusVDB As clsAplusVDB
Public VdbPznRec As New ADODB.Recordset
Public VdbPznRec2 As New ADODB.Recordset
Public VdbBedRec As New ADODB.Recordset
Public VdbBedRec2 As New ADODB.Recordset
Public VdbBerRec As New ADODB.Recordset
Public VdbVebRec As New ADODB.Recordset
Public VdbPauschaleRec As New ADODB.Recordset
Public VdbHinweiseRec As New ADODB.Recordset
Public AplusVOk%

Public AutIdemDB As clsAutIdemDB
Public AutIdemRec As New ADODB.Recordset

Public AktUhrzeit%

Public ActProgram As Object

Public ActBenutzer%

Public INI_DATEI As String

Public RezNr$

Public AllesFlag%, ImpFlag%, EasyImpFlag%, VmFlag%
Public RezeptHuellen%

Public RezApoNr$, OrgRezApoNr$, RezApoName$(1), RezApoDruckName$, BtmRezDruckName$
Public ZuzFeld!(7)

Public ActBundesland%, ActKasse%, ActVerordnung%, ActVebNr&, ActPauschaleNr&, OrgBundesland%
Public BeigetreteneVereinbarungen$


Public DruckSeite%

Public VmRabattFaktor#

Public RezeptDrucker$
Public RezeptDruckerPara$
Public IstDosDrucker%

Public EingabeStr$

Public EasyMatchModus%

Public ImpfstoffeDa%

Public Taetigkeiten(0) As TaetigkeitenStruct
Public AnzTaetigkeiten%

Public SonderBelege(10) As SonderBelegeStruct
Public AnzSonderBelege%

Public RezepturMitFaktor%

Public AnzBlink%
Public AnzRezepte%

Public DSKNurRezeptnummerDrucken%
Public RezeptDetect%, RezeptDruckPause%
Public BtmAlsZeile%

Public RezepturDruck%
Public AvpTeilnahme%
Public DatumObenPrivatRezept%

Public SonderBelegRezept%
    
'Public MagSpeicherIndex%

Public AutIdemIk&, AutIdemKbvNr&

Public FormErg%
Public FormErgTxt$

Public ParEnteralPzn$(5)
Public ParEnteralTxt$(5)
Public ParEnteralPreis#(5)
Public ParenteralPara$, ParenteralHash$
Public ParenteralRezept%, Parenteral_AOK_LosGebiet%, Parenteral_AOK_NordOst%
Public ParEnteralAufschlag#(1)
Public FiveRxFlag%
Public ParEnteralPrimärPackmittel%
Public ParEnteralAI%
Public ParEnteralAnzEinheiten#
Public ParEnteralHerstellerKey%
Public ParEnteralHerstellerKz$(4)
Public pCharge$
Public HashErstellDat As Date
Public PreisKz_62_70%

Public RezeptFarben&(3)
Public MagDarstellung&(6, 1)

Public DruckDebugAktiv%

Public SQLStr$

Public RezeptMitHilfsmittelNr%

Public EditErg%
Public EditModus%
Public EditTxt$
Public EditAnzGefunden%
Public EditGef%(49)

Public AbrechnungsVerfahren_2_Aktiv%

Public BtmDerivatMengeIgnorieren%

Public AlleRezepturenMitHash As Boolean

Public TA1_V37%

Private Const DefErrModul = "WINREZX.BAS"

Sub InitMisc()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitMisc")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

ProgrammNamen$(0) = "Rezeptkontrolle"
ProgrammNamen$(1) = "Auswertung Rezeptspeicher"

Call DefErrPop
End Sub

'Sub Main()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("Main")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim i%, erg%, tmp%, ind%, t%, m%, j%, c%, AnzPrivat%
'Dim h$, s$
'Dim AutIdemDB As Database
'Dim AiTd As TableDef
'
'
'If (App.PrevInstance) Then End
'If (Command$ = "") Then End
'
'MsgBox (Command$)
'
'If (Dir$("fistam.dat") = "") Then ChDir "\user"
'INI_DATEI = CurDir + "\winop.ini"
'
'
'Set ast = New clsStamm
'Set taxe = New clsTaxe
'Set kiste = New clsKiste
'Set para = New clsOpPara
'Set wpara = New clsWinPara
'Set vk = New clsVerkauf
'Set RezTab = New clsVerkRtab
'Set VmPzn = New clsVmPzn
'Set VmBed = New clsVmBed
'Set VmRech = New clsVmRech
'
'
'UserSection$ = "Computer" + Format(Val(para.User))
'Call wpara.HoleWindowsParameter
'
'frmAction.Show
'
'Call para.HoleFirmenStamm
'Call para.AuslesenPdatei
'Call para.HoleZuzahlungen
'Call para.EinlesenPersonal
'
'If (InStr(para.Benutz, "r") <= 0) Then
'    Call iMsgBox("Dieses Programm hat Ihre Apotheke nicht gekauft !", vbCritical)
'    wpara.ExitEndSub
'    Call frmAction.frmActionUnload
'    End
'    Call DefErrPop: Exit Sub
'End If
'
'ActBenutzer% = HoleActBenutzer%
'
'ast.OpenDatei
'
'VerkaufDbOk% = (Dir("Verkauf.mdb") <> "")
''VerkaufDbOk% = 0
'If (VerkaufDbOk) Then
'    Set VerkaufDB = OpenDatabase("Verkauf.mdb", False, True)
'    Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
'    VerkaufRec.Index = "Unique"
'Else
'    vk.OpenDatei
'    RezTab.OpenDatei
'End If
'
'TaxeOk% = False
'
'erg% = 0
'h$ = para.TaxeLw + ":\taxe\taxe.mdb"
''h$ = "m:\taxe\taxe.mdb"
'Set TaxeDB = taxe.OpenDatenbank(h$, False, True)
'
'TaxeOk% = True
'Set TaxeRec = TaxeDB.OpenRecordset("Taxe", dbOpenTable)
'
'AutIdemOk% = False
'AutIdemSonderregelOk% = False
'h$ = para.TaxeLw + ":\taxe\autidem.mdb"
'If (Dir$(h$) <> "") Then
'    AutIdemOk% = True
'    Set AutIdemDB = OpenDatabase(h$, False, True)
'    For Each AiTd In AutIdemDB.TableDefs
'        If (UCase(AiTd.Name) = "AUTIDEMSONDERREGEL") Then
'            AutIdemSonderregelOk% = True
'            Exit For
'        End If
'    Next AiTd
'    AutIdemDB.Close
'End If
'
'kKassenOk% = False
'h$ = para.TaxeLw + ":\taxe\kkassen.mdb"
'If (Dir$(h$) <> "") Then
'    kKassenOk% = True
'    Set kKassenDB = OpenDatabase(h$, False, True)
'End If
'
'
'VmPzn.OpenDatei
'VmBed.OpenDatei
'VmRech.OpenDatei
'
'erg% = ActProgram.RezKontrInit%
'If (erg% = False) Then Call DefErrPop: Exit Sub
'
'h$ = UCase(Command$)
'
'If (Left$(h$, 2) = "T:") Then
'    h$ = Mid$(h, 3)
'    VerkaufDbOk% = (Dir(h$) <> "")
'    If (VerkaufDbOk) Then
'        VerkaufDB.Close
'        Set VerkaufDB = OpenDatabase(h$, False, True)
'        Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
'        VerkaufRec.Index = "Unique"
'
'        AnzPrivat% = 0
'
'        RezNr$ = VerkaufRec!RezeptNr
'        If (Left$(RezNr$, 1) = "P") Then
'            AnzPrivat% = AnzPrivat% + 1
'            s$ = "Privat" + Str$(AnzPrivat%)
'        Else
'            s$ = RezNr$
'        End If
'        frmAction.Caption = s$
'        h$ = Mid$(h$, ind% + 1)
'        If (ActProgram.RezeptHolen) Then
'        '    Call ActProgram.RepaintBtmGebühr
'        '    Call ActProgram.ShowNichtInTaxe
'            Call ActProgram.DruckeRezept
'        End If
'    End If
'Else
'    If (Right$(h$, 1) <> ",") Then
'        h$ = h$ + ","
'    End If
'
'    Do
'        ind% = InStr(h$, ",")
'        If (ind% > 0) Then
'            RezNr$ = Trim(Left$(h$, ind% - 1))
'            If (Left$(RezNr$, 1) = "P") Then
'                AnzPrivat% = AnzPrivat% + 1
'                s$ = "Privat" + Str$(AnzPrivat%)
'            Else
'                s$ = RezNr$
'            End If
'            frmAction.Caption = s$
'            h$ = Mid$(h$, ind% + 1)
'            If (ActProgram.RezeptHolen) Then
'            '    Call ActProgram.RepaintBtmGebühr
'            '    Call ActProgram.ShowNichtInTaxe
'                Call ActProgram.DruckeRezept
'            End If
'        Else
'            Exit Do
'        End If
'    Loop
'End If
'
'
'Call WinArtDebug("vor Programmende")
'Call ProgrammEnde
'
'Call DefErrPop
'End Sub

Sub Main()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Main")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, erg%, tmp%, ind%, t%, m%, j%, c%, AnzPrivat%
Dim h$, s$
'Dim AutIdemDB As Database
Dim AiTd As TableDef
Dim AvTestRec As New ADODB.Recordset

If (App.PrevInstance) Then End
'If (Command$ = "") Then End

AbrechnungsVerfahren_2_Aktiv = (Format(Now, "YYYYMMDD") >= "20160701")
If (UCase(Command) = "ABR2") Then
    AbrechnungsVerfahren_2_Aktiv = True
End If

If (Dir$("fistam.dat") = "") Then ChDir "\user"
INI_DATEI = CurDir + "\winop.ini"

Set para = New clsOpPara
Set wpara = New clsWinPara
Call wpara.HoleWindowsParameter

Set sqlop = New clsSqlTools
If (sqlop.SqlInit = 0) Then
    End
End If

Set ast = New clsStamm
Set ass = New clsStatistik
Set Artikel = New clsArtikelDB
'Set taxe = New clsTaxe
Set taxeAdoDB = New clsTaxeAdoDB
Set kKassenDB = New clsKkassenDB
Set kiste = New clsKiste
Set vk = New clsVerkauf
Set RezTab = New clsVerkRtab
Set VmPzn = New clsVmPzn
Set VmBed = New clsVmBed
Set VmRech = New clsVmRech
Set AplusVDB = New clsAplusVDB
Set AutIdemDB = New clsAutIdemDB
Set VerkaufAdoDB = New clsVerkaufDB
Set hTaxe = New clsHilfsTaxe
Set Hilfstaxe = New clsHilfstaxeDB


UserSection$ = "Computer" + Format(Val(para.User))
'Call wpara.HoleWindowsParameter

AplusVOk% = False
If (AplusVDB.DBvorhanden) Then
    AplusVOk = AplusVDB.OpenDB
End If
If (AplusVOk) Then
    SQLStr = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'VdbPauschalen'"
    AvTestRec.Open SQLStr, AplusVDB.ActiveConn
    If (AvTestRec.EOF) Then
        Call Shell("WinRezX2.exe", vbNormalFocus)
        End
    End If
    AvTestRec.Close
End If

frmAction.Show

Call WinArtDebug("Main")

Call para.HoleFirmenStamm
Call para.AuslesenPdatei
Call para.HoleZuzahlungen
Call para.EinlesenPersonal

Call WinArtDebug("vor Benutz")

If (InStr(para.Benutz, "r") <= 0) Then
    Call iMsgBox("Dieses Programm hat Ihre Apotheke nicht gekauft !", vbCritical)
    wpara.ExitEndSub
    Call frmAction.frmActionUnload
    End
    Call DefErrPop: Exit Sub
End If

ActBenutzer% = HoleActBenutzer%

'ast.OpenDatei
ArtikelDbOk = 0
If (Artikel.DBvorhanden) Then
    ArtikelDbOk = Artikel.OpenDB
    If (ArtikelDbOk) Then
        ArtikelDbOk = Hilfstaxe.OpenDB
    End If
End If
If (ArtikelDbOk% = 0) Then
    End
End If


Call WinArtDebug("vor Verkaufsdatei")

'VerkaufDbOk% = (Dir("Verkauf.mdb") <> "")
'If (VerkaufDbOk) Then
'    Set VerkaufDB = OpenDatabase("Verkauf.mdb", False, True)
'    Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
'    VerkaufRec.Index = "Unique"
'Else
'    vk.OpenDatei
'    RezTab.OpenDatei
'End If

VerkaufDbOk = 0
If (VerkaufAdoDB.DBvorhanden) Then
    If (VerkaufAdoDB.SqlServerDB) Then
        VerkaufDbOk = VerkaufAdoDB.OpenDB
    Else
        VerkaufDbOk% = (Dir("Verkauf.mdb") <> "")
        If (VerkaufDbOk) Then
            Set VerkaufDB = OpenDatabase("Verkauf.mdb", False, True)
            Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
            VerkaufRec.index = "Unique"
        End If
    End If
End If
If (VerkaufDbOk% = 0) Then
    End
End If

Call WinArtDebug("nach Verkaufsdatei")





erg% = 0
'h$ = para.TaxeLw + ":\taxe\taxe.mdb"
'Set TaxeDB = taxe.OpenDatenbank(h$, False, True)
'Set TaxeRec = TaxeDB.OpenRecordset("Taxe", dbOpenTable)
TaxeAdoDBok = 0
If (taxeAdoDB.DBvorhanden) Then
    TaxeAdoDBok = taxeAdoDB.OpenDB
End If
If (TaxeAdoDBok = 0) Then
    End
End If

Call WinArtDebug("nach Taxe")
    
AutIdemOk% = False
AutIdemSonderregelOk% = False
'h$ = para.TaxeLw + ":\taxe\autidem.mdb"
'If (Dir$(h$) <> "") Then
'    AutIdemOk% = True
'    Set AutIdemDB = OpenDatabase(h$, False, True)
'    For Each AiTd In AutIdemDB.TableDefs
'        If (UCase(AiTd.Name) = "AUTIDEMSONDERREGEL") Then
'            AutIdemSonderregelOk% = True
'            Exit For
'        End If
'    Next AiTd
'    AutIdemDB.Close
'End If
If (AutIdemDB.DBvorhanden) Then
    AutIdemOk = AutIdemDB.OpenDB
    AutIdemSonderregelOk% = True
End If
Call WinArtDebug("nach AutIdem")


kKassenOk% = False
'h$ = para.TaxeLw + ":\taxe\kkassen.mdb"
'If (Dir$(h$) <> "") Then
'    kKassenOk% = True
'    Set kKassenDB = OpenDatabase(h$, False, True)
'End If
If (kKassenDB.DBvorhanden) Then
    kKassenOk = kKassenDB.OpenDB
End If



'AplusVOk% = False
'If (AplusVDB.DBvorhanden) Then
'    AplusVOk = AplusVDB.OpenDB
'End If

Call WinArtDebug("vor RezKontrInit")
erg% = ActProgram.RezKontrInit%
If (erg% = False) Then Call DefErrPop: Exit Sub

frmAction.tmrStart.Enabled = True

'MsgBox (Command$)

'Call WinArtDebug("vor Auswertung")
'
'h$ = UCase(Command$)
''Call MsgBox(h)
'
'HochFormatDruck = 0
'ind% = InStr(h$, "HOCH")
'If (ind% > 0) Then
'    HochFormatDruck = True
'    h$ = Trim(Left$(h$, ind% - 1))
'End If
'
'If (Left$(h$, 2) = "T:") Then
'    h$ = UCase(Mid$(h, 3))
'    If (Right$(h$, 4) <> ".MDB") Then
'        h$ = h$ + ".MDB"
'    End If
'    VerkaufDbOk% = (Dir(h$) <> "")
'    If (VerkaufDbOk) Then
'        VerkaufDB.Close
'        Set VerkaufDB = OpenDatabase(h$, False, True)
'        Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
'        VerkaufRec.Index = "Unique"
'
'        AnzPrivat% = 0
'        AnzRezepte% = 0
'
'        RezNr$ = VerkaufRec!RezeptNr
''        If (Left$(RezNr$, 1) = "P") Then
'        If (VerkaufRec!RezeptArt = 5) Or (VerkaufRec!RezeptArt = 6) Then
'            RezNr = "P" + CStr(VerkaufRec!Id)
'            AnzPrivat% = AnzPrivat% + 1
'            s$ = "Privat" + Str$(AnzPrivat%)
'        Else
'            s$ = RezNr$
'        End If
'        frmAction.Caption = s$
'        If (ActProgram.RezeptHolen) Then
'        '    Call ActProgram.RepaintBtmGebühr
'        '    Call ActProgram.ShowNichtInTaxe
'            Call ActProgram.DruckeRezept
'        End If
'    End If
'Else
'    ind% = InStr(h$, "IK")
'    If (ind% > 0) Then
'        h$ = Trim(Left$(h$, ind% - 1))
'    End If
'    If (Right$(h$, 1) <> ",") Then
'        h$ = h$ + ","
'    End If
'
'    AnzPrivat% = 0
'    AnzRezepte% = 0
'    Do
'        ind% = InStr(h$, ",")
'        If (ind% > 0) Then
'            RezNr$ = Trim(Left$(h$, ind% - 1))
'            If (Left$(RezNr$, 1) = "P") Then
'                AnzPrivat% = AnzPrivat% + 1
'                s$ = "Privat" + Str$(AnzPrivat%)
'            Else
'                s$ = RezNr$
'            End If
'            frmAction.Caption = s$
'            h$ = Mid$(h$, ind% + 1)
'            If (ActProgram.RezeptHolen) Then
'            '    Call ActProgram.RepaintBtmGebühr
'            '    Call ActProgram.ShowNichtInTaxe
'                AnzRezepte% = AnzRezepte% + 1
'                Call ActProgram.DruckeRezept
'            End If
'        Else
'            Exit Do
'        End If
'    Loop
'End If
'
'
'Call ProgrammEnde

Call DefErrPop
End Sub

'Sub Main1()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("Main1")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim i%, erg%, tmp%, ind%, t%, m%, j%, c%, AnzPrivat%
'Dim h$, s$
'Dim AutIdemDB As Database
'Dim AiTd As TableDef
'
''If (App.PrevInstance) Then End
''If (Command$ = "") Then End
''
''If (Dir$("fistam.dat") = "") Then ChDir "\user"
''INI_DATEI = CurDir + "\winop.ini"
''
''
''Set ast = New clsStamm
''Set taxe = New clsTaxe
''Set kiste = New clsKiste
''Set para = New clsOpPara
''Set wpara = New clsWinPara
''Set VK = New clsVerkauf
''Set RezTab = New clsVerkRtab
''Set VmPzn = New clsVmPzn
''Set VmBed = New clsVmBed
''Set VmRech = New clsVmRech
''
''
''UserSection$ = "Computer" + Format(Val(para.User))
''Call wpara.HoleWindowsParameter
''
''frmAction.Show
'
'Call WinArtDebug("Main1")
'
'Call para.HoleFirmenStamm
'Call para.AuslesenPdatei
'Call para.HoleZuzahlungen
'Call para.EinlesenPersonal
'
'Call WinArtDebug("vor Benutz")
'
'If (InStr(para.Benutz, "r") <= 0) Then
'    Call iMsgBox("Dieses Programm hat Ihre Apotheke nicht gekauft !", vbCritical)
'    wpara.ExitEndSub
'    Call frmAction.frmActionUnload
'    End
'    Call DefErrPop: Exit Sub
'End If
'
'ActBenutzer% = HoleActBenutzer%
'
'ast.OpenDatei
'
'Call WinArtDebug("vor Verkaufsdatei")
'
'VerkaufDbOk% = (Dir("Verkauf.mdb") <> "")
''VerkaufDbOk% = 0
'If (VerkaufDbOk) Then
'    Set VerkaufDB = OpenDatabase("Verkauf.mdb", False, True)
'    Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
'    VerkaufRec.Index = "Unique"
'Else
'    vk.OpenDatei
'    RezTab.OpenDatei
'End If
'Call WinArtDebug("nach Verkaufsdatei")
'
'erg% = 0
'h$ = para.TaxeLw + ":\taxe\taxe.mdb"
''h$ = "m:\taxe\taxe.mdb"
'Set TaxeDB = taxe.OpenDatenbank(h$, False, True)
'
'Set TaxeRec = TaxeDB.OpenRecordset("Taxe", dbOpenTable)
'Call WinArtDebug("nach Taxe")
'
'AutIdemOk% = False
'AutIdemSonderregelOk% = False
'h$ = para.TaxeLw + ":\taxe\autidem.mdb"
'If (Dir$(h$) <> "") Then
'    AutIdemOk% = True
'    Set AutIdemDB = OpenDatabase(h$, False, True)
'    For Each AiTd In AutIdemDB.TableDefs
'        If (UCase(AiTd.Name) = "AUTIDEMSONDERREGEL") Then
'            AutIdemSonderregelOk% = True
'            Exit For
'        End If
'    Next AiTd
'    AutIdemDB.Close
'End If
'Call WinArtDebug("nach AutIdem")
'
'kKassenOk% = False
'h$ = para.TaxeLw + ":\taxe\kkassen.mdb"
'If (Dir$(h$) <> "") Then
'    kKassenOk% = True
'    Set kKassenDB = OpenDatabase(h$, False, True)
'End If
'
'
'VmPzn.OpenDatei
'VmBed.OpenDatei
'VmRech.OpenDatei
'
'Call WinArtDebug("vor RezKontrInit")
'erg% = ActProgram.RezKontrInit%
'If (erg% = False) Then Call DefErrPop: Exit Sub
'
''MsgBox (Command$)
'
'Call WinArtDebug("vor Auswertung")
'
'h$ = UCase(Command$)
'
'HochFormatDruck = 0
'ind% = InStr(h$, "HOCH")
'If (ind% > 0) Then
'    HochFormatDruck = True
'    h$ = Trim(Left$(h$, ind% - 1))
'End If
'
'If (Left$(h$, 2) = "T:") Then
'    h$ = UCase(Mid$(h, 3))
'    If (Right$(h$, 4) <> ".MDB") Then
'        h$ = h$ + ".MDB"
'    End If
'    VerkaufDbOk% = (Dir(h$) <> "")
'    If (VerkaufDbOk) Then
'        VerkaufDB.Close
'        Set VerkaufDB = OpenDatabase(h$, False, True)
'        Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
'        VerkaufRec.Index = "Unique"
'
'        AnzPrivat% = 0
'        AnzRezepte% = 0
'
'        RezNr$ = VerkaufRec!RezeptNr
''        If (Left$(RezNr$, 1) = "P") Then
'        If (VerkaufRec!RezeptArt = 5) Or (VerkaufRec!RezeptArt = 6) Then
'            RezNr = "P" + CStr(VerkaufRec!Id)
'            AnzPrivat% = AnzPrivat% + 1
'            s$ = "Privat" + Str$(AnzPrivat%)
'        Else
'            s$ = RezNr$
'        End If
'        frmAction.Caption = s$
'        If (ActProgram.RezeptHolen) Then
'        '    Call ActProgram.RepaintBtmGebühr
'        '    Call ActProgram.ShowNichtInTaxe
'            Call ActProgram.DruckeRezept
'        End If
'    End If
'Else
'    ind% = InStr(h$, "IK")
'    If (ind% > 0) Then
'        h$ = Trim(Left$(h$, ind% - 1))
'    End If
'    If (Right$(h$, 1) <> ",") Then
'        h$ = h$ + ","
'    End If
'
'    AnzPrivat% = 0
'    AnzRezepte% = 0
'    Do
'        ind% = InStr(h$, ",")
'        If (ind% > 0) Then
'            RezNr$ = Trim(Left$(h$, ind% - 1))
'            If (Left$(RezNr$, 1) = "P") Then
'                AnzPrivat% = AnzPrivat% + 1
'                s$ = "Privat" + Str$(AnzPrivat%)
'            Else
'                s$ = RezNr$
'            End If
'            frmAction.Caption = s$
'            h$ = Mid$(h$, ind% + 1)
'            If (ActProgram.RezeptHolen) Then
'            '    Call ActProgram.RepaintBtmGebühr
'            '    Call ActProgram.ShowNichtInTaxe
'                AnzRezepte% = AnzRezepte% + 1
'                Call ActProgram.DruckeRezept
'            End If
'        Else
'            Exit Do
'        End If
'    Loop
'End If
'
'
'Call ProgrammEnde
'
'Call DefErrPop
'End Sub

Sub ProgrammEnde()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ProgrammEnde")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim erg%
Dim l&
Dim h$

wpara.ExitEndSub

'TaxeDB.Close
taxeAdoDB.CloseDB

If (kKassenOk%) Then
'    kKassenDB.Close
    kKassenDB.CloseDB
End If

If (ArtikelDbOk%) Then
    Artikel.CloseDB
Else
    ast.CloseDatei
End If

If (AutIdemOk%) Then
'    AutIdemDB.Close
    AutIdemDB.CloseDB
End If


If (VerkaufDbOk) Then
    If (VerkaufAdoDB.SqlServerDB) Then
        VerkaufAdoDB.CloseDB
    Else
        VerkaufDB.Close
    End If
Else
    vk.CloseDatei
    RezTab.CloseDatei
End If


If (AplusVOk) Then
'    AplusVDB.Close
    AplusVDB.CloseDB
'Else
'    VmPzn.CloseDatei
'    VmBed.CloseDatei
'    VmRech.CloseDatei
End If


If (RezSpeicherOK%) Then
    RezSpeicherDB.Close
    KassenDB.Close
End If
    
ast.FreeClass
vk.FreeClass
RezTab.FreeClass
VmPzn.FreeClass
VmBed.FreeClass
VmRech.FreeClass

Set ast = Nothing
'Set taxe = Nothing
Set kiste = Nothing
Set para = Nothing
Set wpara = Nothing
Set vk = Nothing
Set RezTab = Nothing
Set VmPzn = Nothing
Set VmBed = Nothing
Set VmRech = Nothing

Call KillSysTrayIcon(frmAction, 1)

'Call frmAction.frmActionUnload
'End

Call DefErrPop
End Sub

Function HoleActBenutzer%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleActBenutzer%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim mPos1 As clsMpos

Set mPos1 = New clsMpos
mPos1.OpenDatei
mPos1.GetRecord (Val(para.User) + 1)
HoleActBenutzer% = mPos1.pwCode
mPos1.CloseDatei

Call DefErrPop
End Function

Function iMsgBox%(prompt$, Optional buttons% = 0, Optional title$ = "")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("iMsgBox%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim OrgKeinRowColChange%, ret%

OrgKeinRowColChange% = KeinRowColChange%
KeinRowColChange% = True
If (title$ = "") Then title$ = "Rezeptkontrolle"
If (title$ <> "") Then
    ret% = MsgBox(prompt$, vbSystemModal Or buttons%, title$)
Else
    ret% = MsgBox(prompt$, vbSystemModal Or buttons%)
End If
KeinRowColChange% = OrgKeinRowColChange%

iMsgBox% = ret%

Call DefErrPop
End Function

Function FileOpen%(fName$, fAttr$, Optional modus$ = "B", Optional SatzLen% = 100)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FileOpen%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim Handle%

On Error Resume Next
FileOpen% = False
Handle% = FreeFile


If (fAttr$ = "R") Then
    If (modus$ = "B") Then
        Open fName$ For Binary Access Read Shared As #Handle%
    Else
        Open fName$ For Random Access Read Shared As #Handle% Len = SatzLen%
    End If
    If (Err = 0) Then
        If (LOF(Handle%) = 0) Then
            Close #Handle%
            Kill (fName$)
            Err.Raise 53
        Else
            Call iLock(Handle%, 1)
            Call iUnLock(Handle%, 1)
        End If
    End If
ElseIf (fAttr$ = "W") Then
    If (modus$ = "B") Then
        Open fName$ For Binary Access Write As #Handle%
    Else
        Open fName$ For Random Access Write As #Handle% Len = SatzLen%
    End If
ElseIf (fAttr$ = "RW") Then
    If (modus$ = "B") Then
        Open fName$ For Binary Access Read Write Shared As #Handle%
    Else
        Open fName$ For Random Access Read Write Shared As #Handle% Len = SatzLen%
    End If
    Call iLock(Handle%, 1)
    Call iUnLock(Handle%, 1)
ElseIf (fAttr$ = "I") Then
    Open fName$ For Input Access Read Shared As #Handle%
ElseIf (fAttr$ = "O") Then
    Open fName$ For Output Access Write Shared As #Handle%
End If

If (Err = 0) Then
    FileOpen% = Handle%
Else
    Call iMsgBox("Fehler" + Str$(Err) + " beim Öffnen von " + fName$ + vbCr + Err.Description, vbCritical, "FileOpen")
    Call ProgrammEnde
End If

Call DefErrPop
End Function

Sub iLock(file As Integer, SatzNr&)
Dim LockTime As Date
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("iLock")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:

If Err = 70 Or Err = 75 Then
  If LockTime = 0 Then LockTime = DateAdd("s", 20, Now)
  If LockTime > Now Then
    'Sleep (1)
    Resume
  End If
End If

Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

If (para.MehrPlatz) Then Lock #file, SatzNr&

Call DefErrPop
End Sub

Sub iUnLock(file As Integer, SatzNr&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("iUnLock")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

If (para.MehrPlatz) Then Unlock #file, SatzNr&

Call DefErrPop
End Sub


Function MengeErmitteln(TaxMenge As String) As Single

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MengeErmitteln")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim menge As Single
Dim mal As Integer
Dim NextMal As Integer
Dim Füllmenge As Single

menge = Val(TaxMenge)
mal = InStr(TaxMenge, "X")
Füllmenge = menge
While mal > 0 And mal < Len(TaxMenge)
  NextMal = InStr(mal + 1, TaxMenge, "X")
  If NextMal = 0 Then NextMal = Len(TaxMenge) + 1
  Füllmenge = Val(Mid(TaxMenge, mal + 1, NextMal - mal - 1))
  menge = menge * Füllmenge
  mal = NextMal
Wend

MengeErmitteln = menge
Call DefErrPop
End Function


Function CheckZuzZeilenwert%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckZuzZeilenwert%")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ErrNumber%, ret%

On Error Resume Next
ret% = TaxeRec!ZuzZeilenwertKz
ErrNumber% = Err.Number
On Error GoTo DefErr
If (ErrNumber% > 0) Then
    ret% = 0
    If (TaxeRec!VerbandKz) And (TaxeRec!HilfsmittelKz = 0) Then
        ret% = 1
    End If
End If

If (TaxeRec!HimiVerbrauch) Then
    ret = 1
End If

CheckZuzZeilenwert% = ret%

Call DefErrPop
End Function
      
Function CheckNullStr$(s As Variant)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckNullStr$")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ret$

If (IsNull(s)) Then
    ret$ = ""
Else
    ret$ = s
End If

CheckNullStr$ = ret$

Call DefErrPop
End Function

Sub WinArtDebug(sDebug$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WinArtDebug")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim WINARTDEB%, ret%
Dim i%, ind%
Dim s$

If (DruckDebugAktiv%) Then
    s$ = Format(Now, "dd.mm.yyyy") + " " + Format(Now, "hh:nn:ss") + "   " + sDebug$
    
    WINARTDEB% = FreeFile
    Open "WinRezDr.DEB" For Append As #WINARTDEB%
    Print #WINARTDEB%, s
    Close #WINARTDEB%
End If

Call DefErrPop
End Sub

Function HashFaktor$(dFaktor As Double)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HashFaktor$")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

HashFaktor = Format(dFaktor, IIf(TA1_V37, "000000.000000", "00000"))

Call DefErrPop
End Function


Function HashPreis$(dPreis As Double)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HashPreis$")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

If (TA1_V37) Then
    HashPreis = Format(dPreis / 100#, "000000000.00")
Else
    HashPreis = Format(dPreis, "000000000")
End If

Call DefErrPop
End Function


