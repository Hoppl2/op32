Attribute VB_Name = "modWinRezK"
'6.3.05 251129  AE  Einbau Honorierung für Sichtbezug (PZN 18774506): Nichtberücksictigung bei USt, Hashcode
'6.3.04 251126  AE  PZN 06461334 (Stoffe ohne PZN) für den HashCode berücksichtigen: Faktor=1, PreisKz=11
'6.3.03 251016  AE  Selbsterklätung: Abfrage, ob elektronisch oder Ausdruck
'6.3.02 250901  AE  kompiliert mit op1_2025_2.dll
'6.3.01 250701  AE  Einbau Selbsterklärung für elektr.Abrechnung; auch Werte von WinVK direkt in Rezept anzeigen
'6.3.00 250607  AE  Einbau Pflegehilfsmittel-Abrechnung
'6.2.01 250513  AE  Scrollbalken-Spalte für SonderPZNs eingebaut; eRezeptSpeichern: PznStr eingebaut analog zu InitRezeptDruck
'6.2.00 250405  AE  Impfleistungen, ChefModus eingebaut
'6.1.01 250204  AE  'mnuLennatz.enabled=false' bei txtRezeptNr_gotfocus unter kommentar
'6.1.00 250130  AE  Umbau 'BfArm'-Cannabis: fix 100% dazu
'6.0.04 241202  AE  Einbau 'BfArM'-Cannabis: dafür ParenteralPzn erweitert und überall die Prüfung auf ParenteralRezept angepasst
'6.0.03 241201  AE  KeyPress: Call UmspeichernSpezialitaet(IIf(ParenteralRezept > 23, MAG_SPEZIALITAET, MAG_ANTEILIG)) wegen Cannabisblüten
'6.0.02 241116  AE  Einbau CheckFreiText
'6.0.01 241105  AE  Einbau Teilmengenabgabe
'6.0.00 240930  AE  Komplett neue TI_Back.dll
'5.6.03 240923  AE  Leistungserbringergruppenschluessel; mehrere ChargenNr;
'5.6.02 240808  AE  Einbau für Verkleinerung DB: in Vorabprüfung 'Update TI_eRezepte ...' wegen eDispensierung; und in WebService: die 'eRezeptData' nur mahr als Länge in TI_FiveRx speichern
'5.6.01 240623  AE  Einbau PKV_Ausdruck
'5.6.00 240617  AE  Einbau Pharm.DL, xslt für grünes Rezept, Einbau Storno bei ARZ
'                   AMPV_NEU: ab 19.7.2023; jetzt inkl 0.20 EUR für 'Finanzierung zusätzlicher pharmazeutischer Dienstleistungen nach § 129 Absatz 5e des Fünften Buches SGB'
'5.5.00 231221  AE  Modus eRezepte: bei Befüllen flx 'Tabulator' (vbtab) rausnehmen, ansonsten Problem bei der Anzeige

'5.4.00 231221  AE  Einbau des Handlings Preisermittlung ab 1.1.24
'5.3.01 230926  AE  Einbau Anzeige RezeptNr, ChargenNr bei der Auswahl eRezepte
'5.3.00 230815  AE  Lieferengpass für Muster16 + eRezept
'5.2.02 230424  AE  Abhhilfe gegen das fehlende Fenster 'Substitutionsrezept'
'5.2.01 230416  AE  InitRezeptDruck: bei HashFlag und ParenteralRezept den MagSpeicherIndex% unabhängig vom Preis setzen
'5.2.00 230415  AE  HoleMagSpeicher: Berücksichtigung 09999117 (Import RX), 09999206 (Import Substanz RX)
'                   Qualitätszuschlag PZN 04443869 immer mit Menge 1 (v.a. bei Taxmustern)
'                   RezeptHolenDB: Mehrkostenverzicht wenn keine gültigen Rab.Artikel verfügbar (PlusMehrkosten wird nicht gesetzt, erkennbar an Mehrkosten=0)
'                   UmspeichernSpezialitaet: Abhilfe gegen den F3704 (lag an 'TaxeRec.Close' am Ende, jetzt in 'on error resume next' eimgebettet
'5.1.87 221106  AE  Für eRezept: in TI_BACK Packungsgröße brachte F, wenn mit Nachkomma; WriteRezeptSpeicher: die RezNr war bisher die eRezept-Task, jetzt wird in RezeptHolenDB die RezeptNr vom Rezept reingeschrieben
'5.1.86 221019  AE  UmspeichernSpezialitaet: Übernahme sEinheit aus GPMenge
'5.1.85 221007  AE  HA - Hilfstaxe: jetzt Typ ist MAG_ANTEILIG
'5.1.84 221007  AE  WriteRezeptSpeicher: Speicherung 'zusatz2' richtiggestellt, wenn länge nicht > 36
'5.1.83 221007  AE  Auseinzelung TA1_V37: Neue Felder 'Zusatz2a' und 'Zusatz3a': weil 'ALTER TABLE .. TEXT(50)' bringt oft 'Zu wenig Arbeitsspeicher'
'5.1.82 221004  AE  Auseinzelung TA1_V37: Vergrößerung ZusatzX auf 50 Stellen (Struktur und DB), Adaptionen in InitRezeptDruck
'5.1.81 220930  AE  Einbau HA für Hilfstaxe, Gefäße; auch bei Einlesen Taxmuster
'5.1.80 220922  AE  Adaptionen für eRezept: Botendienst, Beschaffungskosten, Komfortsignatur, mehrere auf einmal ....
'5.1.79 220804  AE  Automat. F7 nach F6, TI_BACK aktualisiert für AnzahlPackungen>1
'5.1.78 220804  AE  Neue Kartei 'XML' bei eRezepten, Durchgriff zu 'e', Anzeige XML-Struktur, Anpassung an neues TI_BACK (mit processing instruction in MedicationDispense)
'5.1.77 220731  AE  Anpassungen für F5: eRezept vor Abrechnen neu initialisieren, Prüfung auf Ergebnis des _Abrechnen
'5.1.76 220727  AE  Rückmeldungen ARZ: jetzt sowohl STATUS als auch VSTATUS suchen; Problematik mit Toolbar behoben in op3.dll (leer.gif, bisher bis 16,jetzt für alle 50)
'5.1.75 220727  AE  Methadon: Einzeldosis bis 140: war irrtümlich 245 statt 2.45
'5.1.74 220725  AE  Anpassung eRezept, Zusatzattibut 'Zuzahlungsstaus ....'
'5.1.73 220720  AE  Hashcode-Erstellung: bei den Sätzen mit FaktorKz 55 (zB Levomethadon) fehlte Anpassung an neue TA1_V37
'5.1.72 220708  AE  eRezept: kein Ausdruck bei F6
'5.1.71 220707  AE  eRezept: Farbe links  für Status, Eingabemöglichkeit 'e1' bis 'e5' für Status
'5.1.70 220706  AE  Adaption für eRezept-Anzeige .. Eingabe 'e', Toolbar
'5.1.69 220630  AE  Einbau TA1_V37
'5.1.68 220629  AE  Einbau der Lennartz-Rezepturen
'5.1.67 220405  AE  EinlesenPassendeArbEmb: 'AktKurz' eingeführt, weil es Gefässe gibt, deren Name in den ersten 20 Stellen gleich sind
'                   frmTaxieren - cmdOk: Einbau 'If (ParenteralRezept >= 0) Then MagSpeicherIndex = 1....' wegen Problem Imgrund (F6)
'5.1.66 220323  AE  'kill magtax..' jetzt auch bei Form_Load in frmTaxieren (wegen F6 in IndividuelleRezeptur)
'5.1.65 220314  AE  Einbau 'FiveRxPzn': FiveRxPzn.ini, Optionen, Combo beim Taxieren, Berücksichtigung mittels SumPreisZuz ... für Btmgeb: auch .kp setzen für Faktor
'5.1.64 220302  AE  Einbau 'frm_eRezept';
'                   für Hashcode: jetzt auch 'MAG_Anteilig' berücksichtigen
'5.1.63 220124  AE  Neue ParenteralPzn für Methadon,Levomethadon: jeweils '(Wirkstoff)', die bisherigen werden mit '(FAM)' ergänzt
'5.1.62 220117  AE  Zusätzliche ParenteralPzn "Subutex-Einzeldosen (Take Home)"
'5.1.61 220103  AE  HoleMagSpeicher: BTmGebühr (PZN 02567001) jetzt immer im Hash-Code dabei
'5.1.60 211222  AE  Einbau Aufruf 'eRezepte'
'5.1.59 211201  AE  Btm-Gebühr bei allen Parenteral-Rezepten bei Einlesen eines Btm-Artikels; allg.Rezepturen: Erstellungsdatum 1 min nach Herstellungsdatum; Einlesen Parenteral-TM bei Rezepten aud verkauf-dbb
'5.1.58 211122  AE  Einbau 'AlleRezepturenMitHash': Anpassungen in InitRezeptDruck, WriteRezeptspeicher, HoleMagSpeicher
'                   Adaption 'CheckTaxierungArbEmb' laut GS
'5.1.57 211014  AE  Auseinzelung: Einbau AuseinzelungBtm, Adaption Hashcode (Einbau Btm-Gebühr)
'5.1.56 210930  AE  Zusätzliche ParenteralPzn "Cannabis-Rezeptur BG", damit wird auch Hashcode gedruckt
'5.1.55 210926  AE  Cannabiszuschlag: bei Blüten jetzt kein Preis bei Hilfstaxe, Spezialität (weil ansonsten doppelt berücksichtigt)
'5.1.54 210903  AE  InitRezeptDruck:  If (ParenteralRezept < 21) Then HashErstellDat = Left(h2, 2) + "." + Mid(h2, 3, 2) + ".20" + Mid(h2, 5, 2) + " 00:02" (früher 00:01), wegen Fiverx-Meldung
'                   Der Hertsellungszeitpunkt muß vor dem Erstellungszeitpunkt des Datensatzes liegen'
'5.1.53 210901  AE  Substitutionsrezepte: jetzt pro Abgabe max. 30 ED, entsprechender Hinweis. Multiplikator für Actmenge jetzt 30 (weil jetzt bis zu 30 pro Abgabe möglich, früher 10)
'5.1.52 210818  AE  Cannabisblüten: bei Lagerartikeln wurde irrtümlich der AEP verwendet, behoben
'                   Gefäße OHNE Pzn werden mit der SonderPzn '06461328' und Format=1 für die Hashcode-Eemittlung genommen
'                   wenn bei ParenteralRezept>15 und Spezialität das Verhältnis .ActMenge/.Gstufe>=100 ist, erfolgt Hinweis - wegen Problemen bei der Hashcode-Ermittlung
'5.1.51 210817  AE  CannabisFixAufschlag nur beim 1.Mal dazurechnen für Hashcode (ParenteralPara)
'5.1.50 210811  AE  Substitutionsrezepte: neuen Arbeitstyp 'Subutex-Einzeldosen (KEIN Aut-Idem)': neue Berechnungs-Tabelle dafür integriert, interne PZN '2567114' vor Druck und Hashcode durch '2567113' ersetzen
'                   Bei Subst-Rezepte 18-20: bei Anzahl Einzeldosen>7 jetzt Hinweis und Möglichkeit, trotzdem zu speichern
'5.1.49 210808  AE  Cannabisblüten - jetzt auch den Festpreis von 9.52/g mitberechnen
'5.1.48 210804  AE  Substitutionsrezepte: Uhrzeit bei Zeitpunkt jetzt "00:01"
'5.1.47 210802  AE  Fertigstellen Substitutionsrezepte
'5.1.46 210728  AE  Einbau PreisKz_62_70
'5.1.45 210714  AE  Cannabis - ParenteralPara: für die PZN '06460518' jetzt immer Preis<Kz 74 wie bei den Parenteralia (bisher 62, 70)
'5.1.44 210714  AE  Cannabis: jetzt immer IK als HerstellerKZ, (index 3)
'5.1.43 210714  AE  Cannabis: Herstellungs-Uhrzeit immer '0000'
'5.1.42 210712  AE  'DatumObenRezeptVersatz' eingebaut
'5.1.41 210710  AE  Cannabis: Rezeptur auslesen auch wenn 'RezeptDruck' nicht gesetzt ist
'5.1.40 210624  AE  Cannabis: überarbeitet für Hash-Code entspr. ABDA-Schreiben 'Hinweise zur Techn. Anlage 1, Version 35, Stand 31.5.21
'5.1.39 210528  AE  ImporteOrIdente: bei Importen jetzt 'Rabwert_130a_2_SGB' berücksichtigen (Abschlag Impfstoffe)
'5.1.38 210314  AE  Parenteral/Auseinzelung/Hash: bei Holen bei Rezeptspeicher nicht neu berechnen, kein Speichern in Rezeptspeicher (dafür 'BereitsGedruckt' auf true setzen)
'5.1.37 210213  AE  'Schutzmasken': Einbau Coupon 2, ALG 2 mit neuen SonderPZNs ...
'                   DatumVersatzY jetzt auch im Menü
'5.1.36 210204  AE  Einbau 'DatumVersatzY'
'5.1.35 210129  AE  wegen 'Schutzmasken': Schriftgrößenanpassung für Faktor, Preis weil seeehr groß!
'5.1.34 210107  AE  HoleMagSpeicher: bei Spezialitäten für ParenteralPara jetzt .ActPreis (bisher .kp)
'5.1.33 201218  AE  Einbau 'Schutzmasken'
'5.1.32 201207  AE  CheckAutIdem: Prüfung jetzt auf AusnahmeErsetzungFl=1 wegen neuem Wert 'bedingt'
'5.1.31 201129  AE  wegen Rundungsungenauigkeit bei 'Unverarb.Abgabe': neue Spalte 13 in flxTsxieren für Originalpreis, der bei NeuMalfaktor/MalFaktor verwendet wird.
'5.1.30 201127  AE  Auseinzelung: zusatz2 in Tabelle Artikel zu kurz für 2 AuseinzelungPZNs, deshalb jetzt zusatz3 neu angelegt und dafür verwendet
'5.1.29 201113  AE  Auseinzelung: jetzt auch für 2 Wirkstärken ausgelegt
'5.1.28 200912  AE  ImporteOrAutIdem: für Prüfung BTM jetzt StdMenge statt bisher Menge
'5.1.27 200910  AE  Einbau Ausdruck 'LEGS' zum Druck des HmAbrechnungsKz: jetzt auch für mehrere LEGS, neues Feld in RezSpeicher, RezArtikel
'5.1.26 200908  AE  Einbau Ausdruck 'LEGS' zum Druck des HmAbrechnungsKz
'5.1.25 200807  AE  Einbau MKdurchKK
'5.1.24 200731  AE  AMPV_NEU: jetzt 8.56 statt 8.51
'5.1.23 200707  AE  Einbau 'Patientenalter' in A+V Taxierung
'5.1.22 200115  AE  t-Rezept Aufschlag fix nach Edit (PZN 06460688; bisher Preis 2.91, jetzt 4.26); coNewlineHinweis wurde bisher nicht initialisiert
'5.1.21 191127  AE  Generischer Markt: außer Verkehr-Artikel werden bei der Ermittlung der vier günstigsten AutIdems ignoriert (bei artL neues Feld 'ArtStatus; Anz4billigst nur wenn ArtStatus<>'S'
'5.1.20 191118  AE  Icon für F4 Inhaltsstoffe dazugetan
'5.1.18 191024  AE  Adaption ImporteOrAutIdem analog Kasse; Einbau OrgMenge für BTM; Einbau
'5.1.18 191024  AE  Adaption ImporteOrAutIdem analog Kasse; Einbau OrgMenge für BTM; Einbau ausländ. GTIN in CheckSecurPharm - Codeteil aus OpKasse übernommen
'5.1.17 191017  AE  Einbau 'Rabwert_130a_2_SGB'
'5.1.16 190913  AE  Anzeige "Aufruf Artikelstamm+X mit STRG+X, Nichtverfügbarkeitsabfrage mit Strg+N" in neuer Zeile
'5.1.15 190904  AE  Durchgriffe adpatiert, neu ist WinMsv3
'5.1.14 190903  AE  Public BtmDerivatMengeIgnorieren% eingebaut; beigetretene Vereinbarungen (jetzt in GetPrivateProfileIni mit 2000 Stellen)
'5.1.13 190827  AE  ImporteOrAutIdem: Einbau OrgNGr
'5.1.12 190806  AE  ImporteOrAutIdem: Importe: ist der Festbetrag eines Imports/Originals kleiner als der VK-Herstellerrabatt, wird in der Importanzeige, der Sortierung und der Prüfung der Günstigkeit nun der Festbetrag herangezogen.
'5.1.11 190709  AE  ImporteOrAutIdem: Prüfung auf ImportMarkt erweitert: If (ImportGr > 0) And (TaxeTmpRec!ImportGruppeNr <> ImportGr) Then nurImporte = False
'5.1.10 190628  AE  ImporteOrAutIdem: irrtümlich war da an eienr Stelle 'sCheckNull', jetzt anstelle davon CheckNullStr
'5.1.9  190625  AE  frmImportOrAutIdem: PZN-Spalte sichtbar (1.Spalte)
'5.1.8  190621  AE  Übernahme gewählter Artikel in frmImportOrAutIdem
'5.1.7  190621  AE  Nichtverfügbarkeit abhängig vor/nach 1.7.19, Farbe Ausgangsartikel ...
'5.1.6  190621  AE  Nichtverfügbarkeit erweitert auf 9 Zeilen (bisher 7)
'5.1.5  190620  AE  Einbau neuer Rahmenvertrag
'5.1.4  190527  AE  Einbau Durchgriff auf Kundendaten, PlusX
'5.1.3  190420  AE  TaxSummeEditSatz: bei Parenteral-Rezepten jetzt kein automatischer Fix-Aufschlag mehr
'5.1.2  181204  AE  EditSatz: max. Übergabe von 19 Zeichen wegen ansonsten 'CheckSecurPharm' in dll
'5.1.1  181025  AE  Einbau zwingender Aufruf 'ClassicLine' durch Ini-Eintrag!
'5.1.0  180914  AE  Einbau 'Stützstrumpf-Problematik': ZusatzkomponentenNr, ZusatzKomponentenBasisNr,ZusatzKomponentenFaktor
'                   Privatrezepte: Ausdruck Hochformat: Ausdruck von mehr als 3 Positionen
'5.0.5  180912  AE  Einbau Securpharm: Auslesen DataMatrixCodes (in DLL und in WÜ); Umbau Handling Eingabestr (in picRezept_KeyPress)
'5.0.3  180820  AE  Einbau Auslesen QR-Codes (in DLL und in WÜ)
'5.0.25 180723  AE  SucheInGruppeMDB: jetzt immer 'PauschaleNr' in SQL-String, denn wenn keine Pauschale gewünscht ist, dürfen auch nur die Sätze ohne PauschaleNr genommen werden!
'5.0.24 180714  AE  Einbau 'HilfsmittelAbrechnungsKz' in A+V Berechnung: wenn nur EIN Hilfmittel...Kz, dann wird automatisch dieses genommen ohne Aufblenden des Auswahl-Fensters
'5.0.23 180620  AE  Anzeige der gefundenen Taxe wenn Rezept aus Vormonat
'5.0.22 180508  AE  Einbau 'HilfsmittelAbrechnungsKz' in A+V Berechnung
'5.0.21 180423  AE  IVF-Rezepte: Berücksichtigung Festbetrag in RezeptHolenDB
'5.0.20 180420  AE  Neue Bedruckung: mit Mehrkosten (werden intern in Zusatz(1) gespeichert, Verwendung in InitRezeptDruck): jetzt nur wenn in verkaufsdatei als Hilfsmittel gesetzt und Mehrkosten>0
'5.0.19 180410  AE  Neue Bedruckung: mit Mehrkosten (werden intern in Zusatz(1) gespeichert, Verwendung in InitRezeptDruck)
'5.0.18 180330  AE  InitRezeptDruck: neu für Auseinzelung: If (ParEnteralHerstellerKey < 1) Then ParEnteralHerstellerKey = 3 End If; Befüllen von .Zusatz(1)
'5.0.17 180322  AE  InitRezeptDruck: neu: If (ParEnteralHerstellerKey < 1) Then ParEnteralHerstellerKey = 1 End If
'5.0.16 180322  AE  InitRezeptDruck: 'If (Trim(ParEnteralHerstellerKz(ParEnteralHerstellerKey)) = "") Then ...' pCharge auf "" setzen, "Auseinzelung" wieder weg
'5.0.15 180320  AE  InitRezeptDruck: nun verhindert, daß pCharge zu kurz ist und dadurch zu Problemen bei FiveRx führt!
'5.0.14 180319  AE  WriteRezeptSpeicher: auch bei Auseinzelung Hash-Code speichern, spez. pCharge 'Auseinzelung'
'5.0.13 180313  AE  frmAuseinzelung: PZN-Eingabe auf 8-stellig angepasst
'5.0.12 171205  AE  Abfrage Hoch-/Querformat mittels Ini-Eintrag untrerdrückbar
'5.0.11 171201  AE  Einbau PrivatRezeptVersatzY%, Auslesen aus Ini fehlte
'5.0.10 171120  AE  DruckeRezept: Abfrage nach Hoch-/Querformat bei Privatrezepten eingebaut, analog OpKasse (abh. von PrivRezDruckHoch)
'5.0.9  170608  AE  Beim Holen von Taxmustern 'Fix-Aufschlag' nicht multiplizieren, unabhängig von der Taxmustermenge; bei Rezepturen mit 'Unverarb. Abgabe' jetzt Sonder-PZN 06460702 drucken (anstatt 09999011)
'5.0.8  170602  AE  'AlteTaxeAktiv' eingebaut: für Rezepte mit Abgabedatum in einem früheren Taxe-Zeitraum!
'5.0.7  170531  AE  InitRezeptDruck: bei Verschieben Noctu (2567018) Berücksichtigung der Verfügbarkeiten
'5.0.6  170518  AE  'Fix-Aufschlag'-Zeilen in Rezepturen nicht veränderbar
'5.0.5  170517  AE  Prüfung auf 'Fix-Aufschlag' bei Einlesen Taxmuster; 'Fix-Aufschlag' immer direkt nach Arbeits-Zeile; keine Abfrage auf BTM-Rezepttyp mehr; t-Rezept Aufschlag fix nach Edit (PZN 06460688; Preis 2.91)
'5.0.4  170516  AE  Einbau 'Fix-Aufschlag' von 8.35: in TaxSumme, neu auch row=flxtaxieren.row nach Aufruf TaxSumme in txtTaxieren.KeyPress für TxtCol=1
'5.0.3  170403  AE  DruckKundenlisteKopf: jetzt fix 'Sehr geehrte(r)'
'5.0.2  170331  AE  Sonderpreis im Arbeits-/Infobereich
'5.0.1  170314  AE  MARS Rezeptkontrolle: Privatrezepte (bei Rezeptdruck wird ja die ursprüngliche Nummer eingelesen, deshalb RezNr2$ eingeführt zum Prüfen des Rezeptes (RezeptHolenDB, HoleausRezeptSpeicher)
'5.0.0      170222     Einbau Kunden als SQL; Adaption Ruekkauf.cls für 8-stellige PZN
'4.0.89 170207  AE  Code128: jetzt unabhängig von AVP und FiveRx
'4.0.88 170203  AE  Code128: bisher musste FiveRx vorhanden sein, jetzt nur das erste If mit AvpTeilnahme oder FiveRx
'4.0.87 161228  AE  MARS Rezeptkontrolle: 'EntferneTask' anstatt 'TerminateProcess'
'4.0.86 161223  AE  MARS Rezeptkontrolle: Einbau 'MarsRezeptZurückgestellt', 'StandardDrucker'
'4.0.85 161221  AE  MARS Rezeptkontrolle: Anzeige Freigabedatum,-Personal; Prüfung auf PrivatRezept adaptiert;  Liste DrS vom Mail 16.12.
'4.0.84 161214  AE  RezeptHolenDB: Abbruch bei AnzRezeptArtikel=0;  Überarbeitung Breite linke Box (lblMarsModus, txtRezeptNr, flxKKassen, ...)
'4.0.83 161212  AE  MARS Rezeptkontrolle eingebaut (2 Modi, ....)
'4.0.82 161130  AE  InitRezeptDruck: wenn Noctu vorhanden, dann ans Ende der RezArtikel
'4.0.81 161114  AE  Einbau Code128: Menü,Ini,Druck
'4.0.80 161021  AE  Einbau AbholerSql%
'4.0.79 161005  AE  Suche nach PrivatRezept jetzt auch ohne 'P' bei Eingabe möglich
'4.0.78 160829  AE  Für MARS: bei Privat-Rezepten Ausdruck einer Art Serienbrief
'4.0.77 160823  AE  CheckAutIdem,ImporteOrAutIdem,CheckRabattAutIdem: Einbau Prüfung auf 'HYDROMORPHON', damit zB AutIdem der PZNs 01909161 und 10084268
'4.0.76 160719  AE  Für MARS: Passwortabfrage bei Programmstart; NICHT-Chefs dürfen nichts ändern; bei Besorgern prüfen, ob bereits abgeholt, ansonsten Hinweis und Druck nicht möglich; bei Privat-Rezepten Ausdruck von 'Privat Rezept Beilage.pdf'
'4.0.75 160630  AE  ApvParse: bei AMPV_Alt jetzt auch korrekt abziehen: erg# = CalcAMPV(ActAep# * ActMenge, ActMwst%) / (1# + (ActMwst%) / 100#) - (ActAep# * ActMenge)      bisher nur '- ActAep#
'4.0.74 160621  AE  Ab 1.7. neue Berechnungsreihenfolge für AbrechnungsKz=2: jetzt zuerst mit Faktor multiplizieren, dann MwSt,Rabatt; dafür .VdpPauschale zweckentfremdet; Adaption InitRezeptDruck; zum Test Aufruf mit 'ABR2'
'4.0.73 160602  AE  sF4: Anzeige VebNr/PauschalNr als Tooltip
'4.0.72 160522  AE  .Fortsetzung eingebaut (wegen Übersteuerung Preis,Zuzahlung in OpKasse)
'                   Anzeige 'LastOVP' bei sF4
'4.0.71 160415  AE  AplusV überarbeitet (laut Fehlerprotokolle)
'4.0.70 160405  AE  AplusV überarbeitet (APVparse eingebaut wegen neuer Inhalte: RUNDE,MWST,EURO,AMPREISV)
'4.0.66 160302  AE  AuswahlArbEmb: pos jetzt long, bisher int (wegen Id in Hilfsmittels über 32xxx)
'4.0.65 160210  AE  CalcZuzaNeu für HimiVerbrauch adaptiert: siehe Kommentarzeilen dort
'4.0.64 160125  AE  Abfrage "Stimmt das Abgabedatum des BTM ...." ins PaintRezept, nur Ok, kein Abbruch
'4.0.63 160125  AE  Bei Einlesen BTM-Rezept Abfrage "Stimmt das Abgabedatum des BTM ...."; wenn nein, Abbruch des Einlesens
'4.0.62 160114  AE  Handling §302: in VerbandmittelMDB die Prüfung auf 'AbrechnungsverfahrenKz=2' auch vor die Erzeugung/Prüfung auf HmNummer,2.Zeile (Übersteuerung Variable Paragraph302)
'4.0.61 151215  AE  Prüung auf 32000 in RezeptHolenDB entfernt
'4.0.60 151204  AE  Einbau PreisKz '74' auch für "09999146" (indiv. hergest. Schmerzlösung) für HashCode
'4.0.59 151119  AE  Handling §302: jetzt auch Berücksichtigung 'PZN' in AbrechNr in Berechnung (.tkkpzndruck)
'4.0.58 151116  AE  Einbau PreisKz '74' auch für "09999169" (indiv. hergest. parent. virustatikahaltige Infusionslösungen) für HashCode
'4.0.57 151105  AE  Handling §302: jetzt mittels AbrechnungsVerfahrenKz=2 in VdbBedingungen; 'ABR302' in VdbVOKG nicht mehr vorhanden
'4.0.56 151029  AE  Einbau PreisKz '74' (Zytostatika UND Calcium.. UND Lösungen mit monokl...) für HashCode
'4.0.55 151028  AE  Einbau PreisKz '74' (Zytostatika) für HashCode
'4.0.54 151001  AE  Wenn IVF-Rezept, dann diese Sonder-PZN immer in 1.Zeile
'4.0.53 150922  AE  WinRezDebug eingebaut
'4.0.52 150721  AE  PrivRezept als BtmRezept: in RezeptHolenDB früher PrivatRezept=0, unter Kommentar; in RezeptHolenDB setzen BtrmRezept auch bei PrivatRezept; NeuerArtikel: setzen BtmRezept auch bei PrivatRezept
'4.0.51 150705  AE  UmspeichernSpezialitaet: bei FAM=Substanz (ttyp=mag_anteilig) Taxe-EK nehmen falls > 0
'4.0.50 150430  AE  Abhilfe bei Rezepturen: wenn man sich Rezepturen aus der Kassa auf dem Rezept mit sF8 angeschaut hat und dann mit ESC oder Close geschlossen hat, war beim nächsten Rezept mit Taxierungen die alte Taxierung zusätzlich dabei
'4.0.49 150414  AE  "PlusMehrKosten inaktiviert in RezeptHolenDB" - wieder aktiviert; SonderPzns für Verfügbarkeit und Armin getauscht
'4.0.48 150308  AE  BTM: Ausdruck 4.Zeile (@BTM2@) ermöglicht, dafür wird der BTM-ApoName abgeschnitten; Wirkstoffverordnung (jetzt Feld 'Wirkstoffverordnung' anstatt bisher 'WVORabattArtikel'; F6 nur noch möglich wenn picRezept.visible und tmrF6Speree.enabled=false (500 ms nach letztem Druck)
'4.0.47 150115  AE  Importe: Festbetrag ignorieren; bei 'Abrechnnungsdaten eintragen' "(Extras -> Optionen -> Abrechnungsdaten)" dazu; wenn ARMIN-Rabattartikel von Kasse, dann (W) hinter Namen
'4.0.46 150114  AE  Prüfung auf kunbezug.mdb adaptiert
'4.0.45 150112  AE  Abfrage nach NEUEM Btm-Rezept-Formular entfernt, jetzt immer NEUES Rezeptformular
'4.0.44 141130  AE  Strg+G: Prüfung auf Identgruppe>0; neue Tätigkeit: Importkontrolle (Handling global umgebaut)
'4.0.43 141018  AE  Druckeranpassungen
'4.0.42 141010  AE  PruefeRezkontrDat: TmHeader.ActMenge setzen, ansonsten Preise für Arb, Emb, ... falsch
'4.0.41 141009  AE  IK-Eingabe jetzt auch 9-stellig möglich (führende '10'); HiMiVerbrauch: immer IstWg4 (CheckZuzZeilenWert), RezeptHolenZuz*Faktor in CalcZuzaNeu
'4.0.40 140924  AE  Möglichkeit zum Rücksetzen der IK-Nummer eingebaut: F5 wenn Fokus auf IK-Nr, Sicherheitsabfrage
'4.0.39 140915  AE  CheckAutIdem,ImporteOrIdente: OpAutIdem weg; .AutIdem=3 wenn Rabattartikel ohne AutIdem (grünes Hakerl);
'4.0.38 140826  AE  Aufruf Rezeptspeicher: Prüfung auf Chef-Passwort jetzt erst bei Löschen F5
'4.0.37 140730  AE  PlusMehrKosten inaktiviert in RezeptHolenDB
'4.0.36 140715  AE  Aufruf Rezeptspeicher nur mit Chef-Passwort; Löschen Rezeptspeicher '3112' + aktuellesJahr-2
'4.0.35 140713  AE  Einbau ARMIN WVO....
'4.0.34 140424  AE  Handelt es sich um eine NEUES Btm..', jetzt DefaultButton1
'4.0.33 140404  AE  Prüfung auf Fistam '!' raus (immer aktiv!); ImporteOrAutidem: auch bei Fensteraufblenden ggf. 1.Zeile grün darstellen (durch Aufruf flxedit_rowcolcange); Umlaute in Taxmustern/-Zeilen automatisch richtigstellen (durch CheckUmlaute)
'4.0.32 140325  AE  NeueTaxierung: EinlesenPassendeArbEmb% jetzt (NAME=A.NAME) statt bisher (ID=A.ID)
'4.0.31 140320  AE  Einbau VebNr,PauschaleNr auf Artikelebene; wenn bei NeuemArtikel VerbandMittel keinen Preis liefert, nochmals Aufblender der A+V - Auswahl
'4.0.30 140320  AE  Überarbeitung ImporteOrAutIdem (bei AutIdem keine Berücksichtigung von festbetrag,HerstRabatt)
'4.0.29 140319  AE  Berücksichtigung AMverfügbarkeit bei Shift+Up und Shift+Down; Bei Pauschalen: Prüfung auf alle mit gleichen ersten 7 Stellen der HM-Nummer, dann Zuz nur für 1.; Bundeswehr gebfrei
'                   NeueTaxierung (Feld Emballage in Hilfstaxe, Tabelle Arbeitspreise ...):: Handling für GelatineKapsel
'4.0.28 140317  AE  Berücksichtigung neues Feld 'AusnahmeErsetzungFl'
'4.0.27 140310  AE  NeueTaxierung (Feld Emballage in Hilfstaxe, Tabelle Arbeitspreise ...)
'4.0.26 140127  AE  VdbBedRec!GenehmigungspflichtFl wird von Lauer in der 'neuen' URAL-version nicht mehr geliefert, deshalb Prüfung abgeändert auf (VdbBedRec!GenehmigungsPflicht > 0)
'                   Notdienstgebühr (PZN 02567018): in NeuerArtikel - jetzt Prüfung auf Lagerartikel, wegen Preis
'4.0.25 140109  AE  No-FABSP; Umbau Anzeige Rezeptspeicher (wegen unterschiedlicher Werte in den untergeordneten Fenstern)
'4.0.24 131222  AE  ImporteOrIdente: Berücksichtigung Festbetrag
'4.0.23 131201  AE  Neues Handling für Rezeptspeicher: Speichern+Holen (damit aufgerufene Rezepte gleich wie Rezepte aus der Verkaufsdatei)
'4.0.22 131118  AE  Berücksichtigung 'RezeptNrVersatzY' auch bei BTM-Rezepten
'4.0.21 131115  AE  Abhilfe gegen den F6 bei Ausdruck von mehr als 3 Rezepten mit Rezepturen (Mag_Speicher wird bei Ausdruck nicht geschlossen, wurde deshalb immer größer und MagSpeicherIndex zu groß für RezArtikel.MagIndex, da Wert nur Byte ist
'4.0.20 131113  AE  Einbau '.VdbPauschale', damit für Pauschal-Artikel der Faktor keinen Einfluss auf die Ermittlung RezSumme,RezGebSumme, ... nimmt
'                   Uhrzeitformat bei Noctu-Rezepten; Bei Selbsterklärung keine Avp('X')-Zeile oben
'4.0.19 131004  AE  Anzeige Selbsterklärung überarbeitet: jetzt 2.Zeile 'bitte warten' .. dafür jetzt lblwinvk anstatt auf picrezept zu schreiben
'4.0.18 130921  AE  Selbsterklärung überarbeitet - ua Aufruf WinVK; CheckVerordnung: jetzt Prüfung auf BtmSonderPzn, sollte jetzt endlich passen; Newline: ImporteAutIdem: Selection nur bis zu rotem Preis; Preise bei Privatrezepten; Behebung F91 bei Ausdruck Rezepturen
'4.0.17 130909  AE  AMPV_NEU: jetzt 8.51 statt 8.35; Noctu bei Prüfung auf Sonder-PZN ausnehmen
'4.0.16 130906  AE  Selbsterklärung: als Datum jetzt Vormonat; Taxieren: jetzt OK und ESC
'4.0.15 130822  AE  RezeptHolenDB: Überarbeitung Suchen Rezeptnummer (wieder ohne Tabelle RezeptNummern,Berücksichtigung StornoKunde) nach Rücksprache mit GS
'4.0.14 130812  AE  Hashcode: irrtümlich wurde 'h' verwendet für Datum, deshalb war Hashcode falsch; jetzt 'h2'
'4.0.13 130812  AE  Wieder rückgängig gemacht (aus 3.0.9): "BTM-Rezepte: Adaption CheckVerordnungMdb - jetzt auch Prüfung auf SonderPzn (weil sonst Faktor bei 2567001=0)" -> dies unter Kommentar gesetzt
'4.0.12 130807  AE  Abgabedatum als Herstellungsdatum für Hashcode; Strg+F3 zum Einlesen der Privatrezepte vom VK
'4.0.11 130717  AE  Einbau Selbsterklärung; Knallrot
'4.0.10 130703  AE  Barverkäufe überarbeitet
'4.0.9  130527  AE  Neues FaktorKz 99 (Verwurf); dafür einiges neu
'4.0.8  130508  AE  Überarbeitung zus.Hilfsmittel/WertermittlungKz=2; beigetretene Vereinbarungen (jetzt in GetPrivateProfileIni mit 1000 Stellen)
'4.0.7  130425  AE  Wegfall HmDruck.ini - Ersatz in InitRezeptDruck
'4.0.6  130422  AE  HerstRabatt im AI-Kastl jetzt wieder ohne zus.Berechnung Mwst
'4.0.5  130419  AE  Korrekte Darstellung AutIdemKreuz auch bei ClassicLine
'4.0.4  130418  AE  PaintArtikel - lange Artikelnamen
'4.0.3  130418  AE  Überarbeitung Speichern der Beitritte (jetzt gleich mit SPACE)
'4.0.2  130415  AE  Berücksichtigung 'Berufsgenossenschaften': ZuzFrei, IkSpezialName; Sortierung Vereinbarungen
'4.0.1  130411  AE  Berücksichtigung 'Zuzahlung' in Verbandmittel
'4.0.0  130409  AE  Wawi-SQL und Verkauf-SQL
'3.0.18 130327  AE  CheckVerordnungMdb: Neues Handling für ABR302; dafür auch InitRezeptDruck adaptiert
'3.0.17 130325  AE  Einbau neues A+V Handling
'3.0.16 130131  AE  Einbau PreisDiff-Prüfung
'3.0.15 130123  AE  HoleAusRezeptSpeicher: CheckNullDouble für HerstRabattPrivat130Brutto (wegen schon länger gespeicherten Rezepten)
'3.0.14 130123  AE  AmpvNeu: 8.35 EUR statt 8.10
'3.0.12 121108  AE  Anzeige Taxmuster - Auswahl: ArbEmb mit richtiger Menge
'3.0.11 121106  AE  Aufruf Matchcode bei Artikel im Rezept; Taxmuster: auch selbstangelegte ohne PZN (<> "0000000")
'3.0.10 121103  AE  Überarbeitung Umlaute bei Taxmustern
'3.0.9  120927  AE  BTM-Rezepte: Adaption CheckVerordnungMdb - jetzt auch Prüfung auf SonderPzn (weil sonst Faktor bei 2567001=0)
'3.0.8  120905  AE  alle PZN-Felder in DBs als long; SqlOp in DLL ausgelagert; neue DBs: Abholer,Taxmuster
'3.0.7  120807  AE  Adaption Erstellung Hashcode: damit dies (siehe vorher) geht, neues Feld 'HashErstellungsDatum' in Tabelle Rezepte (damit FiveRx korrektes Datum/Uhrzeit senden kann!)
'3.0.6  120806  AE  Adaption Erstellung Hashcode: jetzt HerstellungsDatum Now-1Stunde und IstDatum jetzt auch mit Stunde,Minuten befüllen (bisher 0000); wegen Rückmeldung von Abrechnungsstelle "Der Herstellungszeitpunkt muss vor dem Zeitpunkt der Hashwerterstellung liegen"
'3.0.5  120801  AE  neues Feld 'pCharge', damit die neuen HerstellerKz's im FiveRx dann auch übertragen werden können
'3.0.4  120731  AE  Adaption Erstellung Hashcode: nach Taxierung jetzt Abfrage nach HerstellerKey, eigene Form
'3.0.3  120730  AE  Adaption Erstellung Hashcode: jetzt HerstellerKz's dazu ...
'3.0.2  120711  AE  ADO 2.7 wegen Microsoft Problem bei 'älteren' Betriebssystemen .. Konvert2Dos in DLL verbessert
'3.0.1  120628  AE  Ini-Handling für DBs überarbeitet (Databases); OpenLieferantenDB in sqlop.bas ausgelagert
'3.0.0          AE  Lieferanten-Daten als Datenbank (Access bzw. SQL-Server)
'2.0.52 01.06.12    AE  Auseinzelung adaptiert (StdMenge,Preisvorschlag,-Berechnung); BtmSonderPzn eingeführt
'2.0.51 06.03.12    AE  RezeptNrPositionAlt eingebaut: default N; damit kann RezeptNr wieder wie früher (unten) gedruckt werden
'2.0.50 24.02.12    AE  Gesamt-Brutto: Beschaff.Kosten jetzt korrekt "9999637", bisher "8150006"
'2.0.49 23.11.11    AE  Taxmuster-Auswahl: bei Losgebiet nur die für Losgebiet (wie bisher), ansonsten die Nicht-Losgebiet-Taxmuster
'2.0.48 18.11.11    AE  Taxmuster für Losgebiet; RezNr bei mehr als 3 Artikeln; Auseinzelung
'2.0.47 29.10.11    AE  Parenteral: Einheitspreis selbst ausmultiplizieren (Faktor*Stoffpreis)
'2.0.46 28.10.11    AE  Einbau AOK Losgebiet
'2.0.45 27.10.11    AE  Einbau AOK Nordost
'2.0.44 26.10.11    AE  Einbau RezeptNrVersatzY
'2.0.43 25.10.11    AE  Rezeptspeicher - Rezeptauswertung F7 (frmRezEinzeln): Abhilfe gegen F wenn kein Rez gefunden (if .rows=1 ...)
'2.0.42 24.10.11    AE  ImporteOrAutIdem: bei Rabattartikeln jetzt gesamte Zeile in der entsprechenden Hintergrundfarbe; Hinweise für AutIdem unten am Rezept (dafür auch AltG eingebaut)
'2.0.41 22.10.11    AE  Strg+G jetzt für Idente (analog Kasse)
'2.0.40 21.10.11    AE  Ausdruck RezeptNr wieder ca. 1.5 mm nach unten
'2.0.39 19.10.11    AE  Adaption Rabatte: in RabWerte nur noch Kassenrabatt,neues Feld für GhRabatte; Anpassung der RezSpeicher-Masken und Ausdrucke
'                       Ausdruck RezeptNr ca. 5 mm nach oben
'2.0.38 17.10.11    AE  Ausdruckmöglichkeit bei Rezeptspeicher-Tages, Einzeln (dafür bei Einzeln auch Datum,KuNr dabei)
'                       Berücksichtigung des Auswertungszeitraumes (bisher fix 1.-31.)
'2.0.37 12.10.11    AE  F91 beim Einlesen BTM-Rezepte: in CheckAutIdem bereinigt
'2.0.36 02.10.11    AE  Druck der Rezeptnr jetzt an der Stelle, wo bei BTm-Rezepten der ApoName
'2.0.35 15.09.11    AE  ImporteOrAutIdem: Weitere Anpassung an Kasse (Mail BA), GhVerfügbarkeit: Wunschartikel jetzt wieder bei Click mit Abfrage; Einbau 'abweichende DAR ...'
'2.0.34 01.09.11    AE  PaintAnzuzeigendeKassen: kleinste Schriftgr. jetzt 8, keine SmallFont; zwingend Linksbündig
'2.0.33 24.08.11    AE  ImporteOrAutiIdem: F13 - lag an cDbl, jetzt xVal
'2.0.32 14.08.11    AE  Auch Privaterezepturen können jetzt gedruckt werden (PruefeRezKontrDat)
'                       neue PZN für Parenteral - werden im Programm in der Ini ersetzt
'                       Preiseingaben ohne Zuzahlung (für beide Kassen)
'2.0.31 02.08.11    AE  Newline: Farbanpassung; auch ImportFenster, private BTM-Rezepte; GhVerfügbarkeit: Wunschartikel jetzt bei LostFocus, bisher in Click
'2.0.30 10.07.11    AE  Gegen F480: picStammdatenBack - wenn nicht der aktive Karteireiter, dann klein machen
'2.0.29 06.07.11    AE  Einbau 'NurHashCodeDruck': ab dem 2.Rezept war HashCode falsch, lag an fehlendem 'close #mag_speicher, Unload frmTaxieren'
'2.0.28 05.07.11    AE  Einbau 'NurHashCodeDruck'
'2.0.27 20.06.11    AE  Alle Parenteral-Arbeitspreise jetzt auch als Privat (zus. Spalte bei Optionen)
'2.0.26 31.05.11    AE  Preiskz jetzt 14 Preiskz jetzt 14 statt bisher 11
'2.0.25 09.05.11    AE  F-Prot 1103-845KO: Handling bei Auswahl SonderPzn
'                       Noctu-ShowNichtInTaxe: bei der Anzeige 'nicht in Taxe ...' ausgenommen; LadeSonderPzn: nur wenn befüllte Zeile, dann berücksichtigen
'2.0.24 02.05.11    AE  PlusMehrkosten in allen Varianten vereinheitlicht (FB ...)
'                       Einlesen Privatrezepte jetzt auch für die neuen mit Nummer eingebaut
'2.0.23 21.04.11    AE  Newline: Farbanpassungen bei AutIdem
'2.0.22 01.04.11    AE  PZN für Noctu wieder auf 2567018 geändert
'2.0.21 31.03.11    AE  Verfügbarkeit - neu dazu: 7 - Wunscharzneimittel; Tabelle Gesamt-Brutto wird auch bei Newline gespeichert
'                       Ausdruck jetzt nur noch wenn PZN 2567024 nicht durch AmVerfügbarkeit eingefügt wurde
'2.0.20 10.03.11    AE  1006-720JG: Parenteral-PZN wird korrekt übernommen bei Wiederaufruf
'2.0.19 04.02.11    AE  PZN für Noctu auf 2567024
'2.0.18 03.02.11    AE  Einbau 'Noctu' für Ausdruck: wenn PZN 2567018 in Rezept, dann auch Ausdruck der Abgabe-Uhrzeit
'                       Verfügbarkeit - neu dazu: 5 - Notwendigkeit unverzüglicher Abgabe, 6 - Pharmazeutische Bedenken
'2.0.17 25.01.11    AE  HolenRezeptMDB - HimiVerbrauch - bei gleichen aufeinanderfolgenden Zeilen wird Zuz jetzt nicht mehr addiert
'2.0.16 15.01.11    AE  Zus. SonderPzn für Parenteral: 'Zytostatika - Zubereitungen privat'
'2.0.15 13.01.11    AE  Sortierung der Rezeptzeilen mittels shift+Up, shift+down
'                       Artikelbezeichnung aus Taxe für Anzeige; FB bei Plusmehrkosten nur wenn FB unter VK
'2.0.14 16.12.10    AE  wenn N0, dann gelbes ! (anstatt bisher autidem=2)
'2.0.13 25.10.10    AE  BtmDerivatMenge berücksichtigen in CheckAutIdem
'2.0.12 22.10.10    AE  die selbsterzeugten idente (GibtsOpAutimdem) ignorieren
'2.0.11 18.09.10    AE  RezeptDruck: ab jetzt auch Faktor drucken, wenn =1
'2.0.10 01.08.10    AE  RezeptHolenMDB: wenn Artikel nur im Stamm (nicht in Taxe) und Wg=3, dann als Rezeptur mit Zuzahlung - Erweitert auch für PRIVATREZEPTE
'2.0.9  14.07.10    AE  RezeptHolenMDB: wenn Artikel nur im Stamm (nicht in Taxe) und Wg=3, dann als Rezeptur mit Zuzahlung
'2.0.8  01.07.10    AE  Parenteral-TM löschen aktiviert
'                       bei Hm-Drucker auchj bei Einzelfaktor wieder 'mal Faktor'
'2.0.7  30.06.10    AE  Überarbeitung Handling HmDruck.ini, Adaption für Ausdruck HmNummer
'                       F-Prots: 1004-679CM, 1005-700HA, 1006-708HA, 1006-714HA, 1006-715KO
'2.0.6  22.06.10    AE  Überarbeitung 'PlusMehrkosten': jetzt Prüfung, ob der Artikel überhaupt einen FB hat; wenn nicht, Feld rücksetzen
'2.0.5  18.05.10    AE  Überarbeitung HmNummer-Druck: jetzt in HmDruck.ini, ob Ausdruck der Hm-Nummer
'2.0.4  11.05.10    AE  HmPositionsNr2 ab jetzt ignorieren
'2.0.3  10.05.10    AE  Parenterale: für Spezialitäten bei Taxmustern: Preis auch wenn Menge > Packungsmenge (CheckTaxierungSpezialitaet)
'2.0.2  04.05.10    AE  Parenterale: für die Berechnung des Hash-Codes bei Spezialitäten jetzt den EK nehmen, nicht mehr den ermittelten Preis
'2.0.1  16.04.10    AE  Speicherung HmNummer,HmFaktor und HmStückPreis für FiveRx
'                       InstitutsKzPraefix eingeführt für Test-Apo-IKs mit '20' am Anfang
'2.0.0  15.04.10    AE  NBNEU-Win7: neue DLL, SendKeys ersetzt durch Eigenumsetzung
'                       P302 eingebaut für Wiederaufruf von §302-Rezepten
'1.1.60 25.03.10    AE  Ausdruck HilfsmittelNr auf Rezept (in CheckVerordnungMDB eingebaut): jetzt keine Abfrage mehr, sondern 'ABR302'in VOKG; Stückpreis bei Ausdruck
'                       BtmRezepte: wenn die BTM-Zeile editiert wird, dann ab diesem Zeitpunkt keine automat. Berechnung des Faktors mehr
'1.1.59 22.03.10    AE  Ausdruck HilfsmittelNr auf Rezept (in CheckVerordnungMDB eingebaut)
'1.1.58 21.03.10    AE  Einbau Parenteral-Taxmuster
'1.1.57 11.03.10    AE  Signaturen: Ini-Eintrag; Abfrage bei Ausdruck, Meldung wenn kein Druck; gewählte BenutzerNr am Rezept
'1.1.56 09.03.10    AE  Überarbeitung BTM-Rezepte: ab jetzt ist der Faktor in der SonderPZN-Zeile die Anzahl der Artikel MIT BTM-KZ
'1.1.55 05.03.10    AE  Eingabe Wirkstoffmenge: jetzt xVal statt Val, damit sowohl '.' als auch ',' eingebbar sind
'1.1.54 03.03.10    AE  Einbau 'picGDI': bei Win95,Win98 ... pro gedruckten Rezept 15 picGDIs erzeugen
'1.1.53 25.02.10    AE  FrmEasyMatch,FrmTaxieren: im Flex jeweils Berücksichtigung der TopRow beim Edit
'1.1.52 21.02.10    AE  Handling der FiveRx-Felder überarbeitet: beim Drucken jetzt 0 bei SendeStatus und AbrechnungsStatus; rzLieferId leer oder ursprüngliche bei Rezept aus Rezeptspeicher
'1.1.51 18.02.10    AE  Überarbeitung der Anzeige der Rückmeldungen der Abrechnungsstelle
'1.1.50 15.02.10    AE  FiveRxFlag: jetzt sowohl bei 'z' als auch 'Z'; Einbau BtmDerivatMenge für AutIdem
'1.1.49 10.02.10    AE  Neuer Karteireiter 'Gesamt-Brutto': damit neue Zeilen anlegbar und Preise editierbar
'1.1.48 09.02.10    AE  Überarbeitung BTM-Rezepte: ab jetzt ist der Faktor in der SonderPZN-Zeile die Anzahl der Artikel
'1.1.47 05.02.10    AE  Überarbeitung Ini-Handling ParenteralPzn; Firmenstamm 'Z' für Parenteral; 'Md5Code.exe' statt 'GhCheck2.exe'
'1.1.46 03.02.10    AE  FiveRx eingearbeitet
'1.1.45 01.02.10    AE  Überarbeitung Parenteral
'1.1.44 29.01.10    AE  Einbau neues Handling Parenteral: 'Anlage3'
'1.1.43 07.01.10    AE  Parameter 'TmCheck' für Prüfung auf rote Zeilen bei Taxmustern: Default=N
'1.1.42 07.01.10    AE  Parameter 'TmCheck' für Prüfung auf rote Zeilen bei Taxmustern (wegen extremer Zeitverzögerung; Default=J)
'1.1.41 02.01.10    AE  Adaption 'PARENTERAL': Default-Belegung ParentalRezept auch für Ind.Rezepturen; Erweiterung SondrPzn.dat
'1.1.40 29.12.09    AE  Adaption 'PARENTERAL': Neue Tabelle 'ParenteralTaxierungen', damit alle Rezeptdaten gespeichert, jetzt auch wieder Ausdruck Rezepturen
'1.1.39 28.12.09    AE  Adaption 'PARENTERAL': bei Auswahl für Aktive jetzt alle mit gleicher M2-Nr (nicht nur Lagernde), neue Berechnung TransaktionsId ..
'1.1.38 13.12.09    AE  Einbau 'PARENTERAL'
'1.1.37 03.12.09    AE  TaxMusterBefüllen: bei 'UNVERARBEITETE ABGABE' nicht Rot; TmmengenFaktor mit 1 initialisieren
'1.1.36 18.11.09    AE  LadeAvDateiMDB: nur wenn Id>0
'1.1.35 03.11.09    AE  Taxmuster aus Kasse: in PruefeRezkontrDat das Holen des Taxmusters eingebaut, auch Ausdruck auf Rezept
'1.1.34 02.11.09    AE  Taxmuster: wenn ein Bestandteil nicht in den Apodaten oder Preis=0, dann rot darstellen
'1.1.33 27.10.09    AE  Neues A+V-Handling (aus MDB), CheckVerordnungMDB
'1.1.32 25.05.09    AE  Anpassung an neues Fehlerhandling
'1.1.31 27.04.09    AE  KundenNr jetzt auch über 32000 ohne Fehler (LONG!!); Datentyp Knr in Tabelle Rezepte automatisch konvertieren mittels SQL (am Nachmittag nochmals überarbeitet)
'1.1.30 07.04.09    AE  Privatrezept-Besorger-Stückelung: in diesem Fall nicht Aconto+Restzahlung, sondern den Preis aud Verkauf.mdb
'1.1.29 31.03.09    AE  Editierung IK-Nummer wieder ermöglicht;
'1.1.28 19.03.09    AE  Für HimiVerbrauch die VerkaufRec!ZuzaGes nehmen (neu in verkauf.mdb: wegen Rundungsumgenauigkeiten ansonsten)
'1.1.27 17.03.09    AE  erweitern für MAX_KKTYP von 50 auf 100
'1.1.26 24.02.09    AE  Abhilfe gegen F3420 im Matchcode beim Taxieren: durch Setzen von SonderPrOk=0 zu Beginn der Func Matchcode (in op32.dll)
'      +16.03.09        bei BTM-Rezepten u.U. auch 4 Zeilen drucken (wenn z.B. 'Nichtverfügbarkeit eines Rabattartikels')
'1.1.25 12.02.09    AE  Mietgebühr: in Struktur neue Komp: MietDauer; wenn gesetzt, dann auf Ausdruck SonderPZN und Dauer als Faktor
'1.1.24 22.01.09    AE  In 1.Zeile auf Bildschirm-Rezept: rechts neben der AutIdem-KK der KKName laut Kundendaten; dafür neue Box 60, gleiche Grösse wie 34, aber rechtsbündig ohne Anzeige des Rahmens, damit sowohl AutIdem-KK linksbündig als auch die KundenRec-KKasse rechtsbündig darstellbar
'1.1.23 12.01.09    AE  WriteRezeptSpeicher auch für Computer>=10 ohne Fehler!
'1.1.22 19.11.08    AE  Aufruf WinStat eingeführt, Hinweis wenn Merkzettel nicht bereit, keine Neuanlage
'1.1.21 17.09.08    AE  MachEinzelBerechnung: Rundung des Ergebnisses mittels fnx
'                       MachBerechnung: wenn False, dann Zuz auf 0 setzen
'1.1.20 17.06.08    AE  Wenn Sonderpreisartikel, dann Feld 'VK' aus Verkauf.mdb nehmen (mit Berücksichtigung FB)
'                       für FlexKasse: barverkäufe jetzt von gesamten Kunden berücksichtigen
'1.1.19 15.05.08    AE  Umbau für Zuzahlungserlass - in RezeptHolen, RezeptHolenMDB, NeuerArtikel, CalcTaxeZuzahlungDOS (damit reduzierte Zuz vor Prüfung auf Rausfall aus Rezept)
'1.1.17 12.03.08    AE  ActkKaase$ in RezeptHolenMDB: jetzt mit führenden Nullen übernehmen, wegen Anzeige im Rezeptspeicher
'                       RezeptHolenMDB: auch Grüne Rezepte bei 'p' berücksichtigen
'1.1.16 03.03.08    AE  'Abfrage zu komplex' bei Aufbereiten der Idente - deshalb jetzt auf 250 PZNs beschränkt
'1.1.15 02.03.08    AE  Tabelle Rezepte: neues Feld 'kKassenIk', für AVP-Connect
'                       RezeptHolenMdb: jetzt auch bei 'u n t' auslesen der AbholNr - wegen Übernahme Taxierungen aus WinAbhol
'1.1.14 17.02.08    AE  MachBerechnung: Klammerung bei FB/2*ST erzeugte FB/(2*ST) - Schwachsinn - deshalb vorab Prüfung on '+' oder '-' überhaupt in Formel vorhanden
'1.1.13 13.02.08    AE  RezeptHolenMDB: Überarbeitung .Appli; Faktor bei untereinander gleichen Zeilen korrekt aufsummieren; Zuz bei HimiVerbrauch; wenn Zuz>9.98 und <10, dann =10
'1.1.12 28.01.08    AE  RezeptHolenMDB: Sprechstundenbedarf berücksichtigen - bzgl. GebFrei
'1.1.11 23.01.08    AE  RezeptHolenMDB: wenn Preis=0 und AbholNr>0, dann Preis=VK
'1.1.10 22.01.08    AE  MachBerechnung überarbeitet: auch für Formeln der Art (MENGE*EK+AMPV_ALT)/MENGE*0,95; auch für EK-15%
'1.1.09 17.01.08    AE  RezeptHolen: PZN wieder mit führenden Nullen; auch wieder Anzeige von Menge+Einheit
'1.1.08 10.01.08    AE  MachBerechnung überarbeitet: da es jetzt EK+ST*0,4 gibt (ohne Klammerung!), wird die Klammerung vom Prog gesetzt
'1.1.07 21.12.07    AE  IVF-Rezepte: taxbetrag jetzt Avp/2, Zuz 0. Als erste Zeile PZN 9999643
'1.1.06 11.12.07    AE  Einbau Verkauf.MDB: RezeptHolen, HoleVkPrivat ... Ivf, ..
'1.1.05 06.12.07    AE  Einbau IVF-Rezepte; Rundung bei A+V adaptiert (clng weg!)
'1.1.04 04.11.07    AE  Generika (!) - rot wenn wegen Rabattvertrag, Gelb wenn wegen 'nicht unter den Günstigsten ...'
'1.1.03 16.08.07    AE  Anpassung HoleKundenInfo an neue Kundenstammdaten (kunden.mdb)
'1.1.02 13.08.07    AE  AM-Verfügbarkeit: auch für BTM-Rezepte
'1.1.01 26.07.07    AE  AM-Verfügbarkeit: cboVerfügbarkeit Schriftgrösse berechnen; RezArtikel mit '1' initialisieren
'1.1.00 25.07.07    AE  AM-Verfügbarkeit: Einbau in Rezeptspeicher, Handling mittels Combos, Auch bei Ansicht in Rezeptspeicher ..
'1.0.99 19.07.07    AE  Rezeptspeicher/Tabelle Artikel: wenn nicht vorhanden, wird der Index 'Unique' bei Programmstart angelegt
'1.0.98 03.06.07    AE  Übernahme der Taxierungen aus WinAbhol (mittels Tabelle Taxierungen)
'1.0.97 01.06.07    AE  Wenn VerkComp>10, dann bisher F bei WriteRezeptSpeicher, behoben
'                       Taxmuster: auch selbstangelegte Substanzen werden jetzt bei Holen von TM korrekt berechnet (durch Einbau MatchHilfstaxe - MakeMatchcode)
'1.0.96 25.04.07    AE  ZuzahlungsErlass:wenn schon durch DSK IK-Nummer da, dann auch für Prüfung Mehrkosten/Zuzahlungserlass verwenden
'1.0.95 20.04.07    AE  ZuzahlungsErlass eingebaut
'1.0.94 05.04.07    AE  OP-AutIdem wieder aktiviert
'1.0.93 03.04.07    AE  Anpassung an AutIdem-Sonderregeln: Überarbeitung CheckAutIdem, ImporteOrAutIdem: KK kann mehrere Sonderregeln haben!
'1.0.92 02.04.07    AE  Anpassung an AutIdem-Sonderregeln: Ausdruck Ik-Auswertung in WinInfo umgelagert (wegen uU kleinen Rezeptdrucker); keine Op-AutIdems mehr!
'1.0.91 26.03.07    AE  Anpassung an AutIdem-Sonderregeln: Auftrennung von AutIdemIkNr und ActKKasse - damit funkt RezSpeicher und ImportKontrolle wie bisher
'1.0.90 22.03.07    AE  Anpassung an AutIdem-Sonderregeln; Überarbeitung
'1.0.89 31.01.07    AE  Anpassung auf 80 Benutzer (Taetigkeiten, ..)
'1.0.88 27.11.06    AE  Druck Avp-RezeptNr: jetzt 12 cpi
'1.0.87 23.11.06    AE  Einbau indiv. HM-Nummer; Avp-Rezept Nr (DB, LaufNr, Parameter, Druck, Eingabe/Aufruf)
'                       Handling Avp nochmals umgebaut nach Gespräch mit AVP; Aufruf AvpSend abhängig von AvpTeilnahme
'                       PruefeRezkontrDat: jetzt AbholerNr 4stellig auslesen aus Rezeptsatz
'1.0.86 13.11.06    AE  Druck Taxierungen: jetzt immer 1.Rezeptur drucken; Parameter ob Ausdruck sein soll (default=nein)
'1.0.85 29.10.06    AE  HiMiVerbrauch: bei Zusammenfassen mehrerer untereinander stehender jetzt Zuzahlung ok
'                       Druck Taxierungen: wenn eine Taxierung da mit Preis>0 (also über RezGeb), dann Ausdruck quer
'1.0.84 22.08.06    AE  BTM-Rezepte: in 4.Zeile BTM-Rezept-Text (Extras/Optionen); nur mehr max 3 PZN bei BTM-Rezept
'1.0.83 31.07.06    AE  BTM-Gebühr hat jetzt bundesweit die Sonder-PZN 2567001 (bisher im Programm: 8150012)
'1.0.82 28.06.06    AE  Besorgerverwaltung: wenn AbholerMdb%, dann Aufruf winabhol
'1.0.81 15.06.06    AE  Mehrkostenverzicht, Überarbeitung PlusMehrKosten in RezeptHolen (jetzt Zuz von Festbetrag)
'1.0.80 12.06.06    AE  Abholer-Verwaltung eingeführt: Aufruf winabhol
'1.0.79 05.05.06    AE 'PlusMehrkosten' eingeführt: wenn in DSK Storno+F8, dann zahlt Krankenkasse die Mehrkosten!
'1.0.78 24.04.06    AE Alles mit abholerDB entfernt wegen WinABhol
'1.0.77 08.01.06    AE Überarbeitung Original, Import;  bei AutIdem Auswahl: F2 - Auswahl Idente
'1.0.76 22.12.05    AE Überarbeitung HimiZumVerbrauch: Deckelung von DSK nehmen; bei EditSatz MsgBox
'1.0.75 28.10.05    AE Überarbeitung AutIdem
'1.0.74 22.08.05    AE Privatrezepte - Besorger - Stückelung (in RezeptHolen; BesorgerPreis nur wenn nicht mit ( beginnt
'1.0.73 07.08.05    AE Einbau Edit SonderPzn; Berücksichtigung Umlaute; auch neg.Werte für DatumVersatzX
'1.0.72 27.07.05    AE Einbau DatumVersatzX
'1.0.71 11.07.05    AE Handling Privatrezept: bei sF3 war CalcTaxeZuzahlungDOS Schwachsinn!
'1.0.70 27.06.05    AE neue op32.dll
'1.0.69 21.06.05    AE Handling Privatbesorger adaptiert: jetzt Preis von Aconto-Zeile nehmen; uU Multiplikator in Zeile danach beachten
'1.0.68 19.06.05    AE Richtigstellung Enabled für MnuBearbeitenZusatz; erweitern für MAX_KKTYP von 25 auf 50
'1.0.67 08.06.05    AE Anpassung an neue A+V Strukturen
'1.0.66 29.05.05    AE SonderBelegRezept eingeführt; auf sF3 jetzt Privatrezept-NEU; Barverkäufe auf Strg+B
'1.0.65 31.03.05 AE analog WAMA
'1.0.64 16.03.05 AE Merkzettel abhängig von Fistam (wenn 't' - MDB, sonst stammlos.dat)
'1.0.63 01.02.05    AE Merkzettel jetzt als MDB
'1.0.62 20.01.05    AE Neues Handling für Zuzahlungsberechnung auf Zeilenwert-Basis: .istwg4 jetzt nicht mehr von WG, sondern von (HilfsmittelKz und NICHT VerbandKz) abhängig!
'                           und das auch nur, wenn feld 'ZuzZeilenwertKz' noch nicht vorhanden
'                           CalctaxeZuzahlungDOS: Handling mit Erstattungsfähig, BedErstattungsFähig geändert!
'1.0.61 27.12.04    AE Adaption Ausdruck für Berlin+RVO
'       16.12.04    GS VK.AVP nur übernehmen, wenn "A" (nicht, wenn "a")
'                   Druck Berlin+RVO: Funktion "MengeErmitteln" (6x20 war falsch)
'1.0.60 16.12.04    AE  spez. Ausdruck für Berlin+RVO
'                       BTM-rezepte: Datum wieder an urspr. Stelle
'1.0.59 06.12.04    GS Krankenkasse löschen
'1 0.58 06.12.04    AE ParseOperand: bisher wurde AMPV... immer fix vom AEP genommen; jetzt Berücksichtigung des etwaiigen Multiplikators MENGE
'1 0.57 03.12.04    AE,GS A+V: Klammern parsen korrigiert, Abfrage nach "VK" vor MwSt-Berechnung: Instr statt Left(,2)
'1.0.56 01.12.04    AE  Anpassung an WinApo-CD (wg 3 wegtun!)
'1.0.55 25.11.04    AE  Neue op32.dll
'1.0.54 11.10.04    AE  Positionierung Abgabedatum NEU für Windows-Drucker
'1.0.53 09.09.04    GS,AE Importquoten neu, Abrechnung quartalsweise
'                   Importalternativen eingeschränkt auf günstige Importe
'1.0.52 08.07.04    GS KK-Stammdaten: Kassentypen aus VM-Kassen-Array
'                   Auswahl KK-Typ bei Stammdaten korrigiert
'1.0.51 17.03.04    AE  Verbandmittel: vor Ermittlung Zuza nur dann mit Menge multiplizieren, wenn IstWg4 gesetzt ist
'1.0.50  3.02.04    GS AMPV_NEU: 8.1 statt 8.13
'1.0.49 28.01.04    AE  Diätetika (Taxe-WG 3): auch Zeilenwert nehmen; dafür IstWg4 verwendet
'1.0.48 27.01.04    AE  Verbandmittel: Für Berechnung der Zuzahlung den A+V Rabattfaktor berücksichtigen; auf Zeilenwert-Basis umgestellt
'1.0.47 19.01.04    AE  CalcTaxeZuzahlungDOS: bei RezeptHolen den Preis aus VK mitgeben und daraus Zuz berechnen (wegen Preisänderung taxe!); auch bei SonderPznArtikeln im Rezept
'1.0.46 14.01.04    AE  RezeptHolen: Zuzahlung bei Kontroll-Stop, Preiseingaben ROT und BLINK
'1.0.45 13.01.04    AE  RezeptHolen: PZN 999999 mit Preis>0 (Preiseingaben) auch Zuzahlung berechnen
'1.0.44 02.01.04    AE  Verbandmittel: Prüfung, ob es aus dem Rezept fällt, adaptiert
'1.0.43 30.12.03    AE  CalcZuzaNeu: wenn HimiVerbrauch, dann keine Untergrenze Zuzahlung
'1.0.42 18.12.03    AE  Adaptionen für 2004
'1.0.41 16.12.03    AE  Änderungen für 2004
'1.0.40 02.12.03    AE,GS  AbrechMonatErmitteln: Abfrage auf EOF an den nötigen Stellen eingebaut
'1.0.39 28.11.03    AE  neue Abholerteile unter Kommentar gesetzt
'1.0.38 21.11.03    AE  DAO351 durch DAO360 ersetzt; CreateDatabase jetzt mit Default-Version
'1.0.37 28.10.03    AE  WindowsDruck: Grund für falsche Einzelzeilen auf Ausdruck gefunden (von Fr.Hasselberg):
'                       wenn man direkt nach Druck nochmals ausdruckt (kein rezept am Schirm), dann falsche Werte -> behoben
'1.0.36 06.10.03    AE  WindowsDruck: da manchmal bei Kunden auf Rezept was falsches gedruckt wird, jetzt Summenprüfung in InitRezeptDruck
'                       eingebaut, mit MsgBox wenn SOllPreis <> IstPreis
'1.0.35 30.09.03    AE  Fehlerprot: mehrere PrivRezepte; wenn wo Abholer dabei, wurde Artikel nicht angezeigt
'1.0.34 28.08.03    AE  F6 wurde manchmal irrtümlich aktiviert, in PaintRezept unter Kommentar gesetzt
'1.0.33 21.08.03    AE  Verbandmittel: MachBerechnung: neue Werte für Rabatt !
'1.0.32 19.08.03    AE  Verbandmittel: auch bei Kostenvoranschlag od. Genehmigungspflicht Preis lt. Berechnung erzeugen
'                       ('erg=FALSE' unter Kommentar gesetzt)
'1.0.31 04.08.03    AE  wegen Kombidrucker 5000: DruckerReset nach Asudruck
'1.0.30 23.06.03    AE  nur kompiliert wegen neuer DLL
'1.0.29 14.04.03    AE  WriteRezeptSpeicher: Prüfung pzn="@" vor Abfrage taxe.mdb eingebaut
'                       IstWg4: nur setzen, wenn kein 20%er
'                       keine Prüfung Fistam 'r' bis 30.6.03
'1.0.28 07.04.03    AE  HolenAusRezeptSpeicher: Richtigstellung Struktur für Berechnung GebSumme (.flag, .istwg4); Druck der Laufnummer
'                       Anzeige Rezeptspeicher: in allen Modi richtige Anzeige des RabattWerts und der Auswertung
'1.0.27 10.03.03    AE  BeitragsSatzSicherungsGesetz, BtmAlsZeile%, AnzBlink
'1.0.26 27.01.03    GS  Importquoten f. 2003
'1.0.25 02.01.03    GS  Rezeptspeicher, Importkontrolle: statt "2002 ges." stand "2012 ges."
'1.0.24 02.01.03    AE  AbrechMonatErmitteln: Funktion umgebaut wegen F3021
'1.0.20 27.10.02    AE  nur neu kompiliert wegen DLL
'1.0.19 09.09.02    AE  HoleAusRezeptSpeicher: Berücksichtigung ScreenPzn wegen F3075
'                       TaxMuster: Speichern ermöglicht wenn noch keine TM da; dafür auch Datei initialisiert
'                       Fehlerprot 12.9.02: bei ImportOrIdent EditErg% mit False initialisieren, damit mit Close oder AltF4 nicht alter ErgWert übernommen wird!
'1.0.18 06.09.02    AE  Aktivieren Ausbesserung Priv VK
'1.0.17 23.08.02    AE  Ausbesserungen Priv VK, Richtigstellung 'AD'
'1.0.16 14.08.02    AE  Richtigstellung Multipli von Zuz: jetzt nur noch für WG=4 auf faktor=1 setzen
'                       dafür .IstWg4 eingebaut
'                       Hinweis bei Neuem Artikel, wenn Zuz = "nicht betroffen"
'                       Mag.Taxieren: Einbau zus. Typ FAM=Substanz
'1.0.15 12.08.02    AE  beim Holen von Taxmustern: wenn für eine Substanz kein Preis, dann Meldung anzeigen
'                       RezeptDruck: wenn Faktor mind. 2stellig, dann fehlte manchmal 7.Stelle der PZN; behoben durch trim
'                       nach HoleDruckZeile für Faktor
'                       Einbau für DOS-Drucker: ev. auch parallet drucken (wenn Parameter LPT)
'1.0.14 31.07.02    AE  PaintRezeptDaten: bei Uhrzeit war falsche Division, dadurch eine Stunde zuviel, wenn Min>50
'                       wenn SonderPZN eingescannt, mit F2 oder von Kasse, Zuazhlung richtig und Namen übernehmen
'1.0.13 16.07.02    AE  Taxieren: auch Preiseingaben < 1 zulassen (wegen lWert&, jetzt lWert#)
'                       ShowNichtInTaxe: nur anzeigen, wenn keine SonderPZN
'       02.07.02    AE  ImporteOrAutIdem: zus. Spalte mit S wenn AußerVerkehr
'                       Bei alter AutIdem-Regelung: schauen ob mindestens 2 mit dieser Packungsgröße
Option Explicit

Const GDI_OBJEKTE_PRO_REZEPT% = 15

Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 ' Service Pack
End Type



Public ProgrammChar$

Public ProgrammNamen$(1)
Public ProgrammTyp%

Global buf$

'Public AutIdemDB As Database
'Public AutIdemRec As Recordset
Public AutIdemDB As clsAutIdemDB
Public AutIdemRec As New ADODB.Recordset

'Public MerkzettelDB As Database
'Public MerkzettelRec As Recordset
'Public MerkzettelParaRec As Recordset

'Public kKassenDB As Database
'Public KkassenRec As Recordset
Public kKassenDB As clsKkassenDB
Public KkassenRec As New ADODB.Recordset

Public VerkaufDB As Database
Public VerkaufRec As Recordset
Public VerkaufAdoDB As clsVerkaufDB
Public VerkaufAdoRec As New ADODB.Recordset
Public VerkaufDbOk%


'Public AplusVDB As Database
'Public VdbPznRec As Recordset
'Public VdbBedRec As Recordset
'Public VdbBerRec As Recordset
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

Public ProgrammModus%
Public AutIdemOk%, AutIdemSonderregelOk%, kKassenOk%

Public FabsErrf%
Public FabsRecno&

Public UserSection$

Public KeinRowColChange%

Public AvKassen$()
Public WochenTag$(6)

Public AvVereinbarungen$()
Public AvPauschalen$()

Public ast As clsStamm
Public ass As clsStatistik
'Public taxe As clsTaxe
Public ww As clsWawiDat
Public BESORGT As clsBesorgt
Public Merkzettel As clsMerkzettel
Public Kiste As clsKiste
Public arttext As clsArttext
Public para As clsOpPara
Public wpara As clsWinPara
Public vk As clsVerkauf
Public RezTab As clsVerkRtab
Public VmPzn As clsVmPzn
Public VmBed As clsVmBed
Public VmRech As clsVmRech
Public nnek As clsNNEK
Public hTaxe As clsHilfsTaxe

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
    
Public Hilfstaxe As clsHilfstaxeDB
Public HilfstaxeRec As New ADODB.Recordset

Public lif As clsLieferanten
Public lifzus As clsLiefZusatz
Public lieftext As clsLieftext
Public Lieferanten As clsLieferantenDB
'Public LieferantenConn As New ADODB.Connection
'Public LieferantenComm As New ADODB.Command
Public LieferantenRec As New ADODB.Recordset
Public LieferantenzusatzRec As New ADODB.Recordset
Public AusnahmenRec As New ADODB.Recordset
Public RabattTabelleRec As New ADODB.Recordset
Public BmEkTabelleRec As New ADODB.Recordset
Public LieferantenDbOk%
    
Public ABDA_Komplett_SQL%
Public ABDA_Komplett_Conn As New ADODB.Connection
Public ABDA_Komplett_Rec As New ADODB.Recordset


Public EditErg%
Public EditModus%
Public EditTxt$
Public EditAnzGefunden%
Public EditGef%(49)

Public ArtikelStatistik%

Public AktUhrzeit%

Public ActProgram As Object

Public ErstAuslesen%

Public ActBenutzer%

Public FarbeGray&
Public INI_DATEI As String

Public RezNr$

Public AllesFlag%, ImpFlag%, EasyImpFlag%, VmFlag%
Public RezeptHuellen%, Knallrot%, PrivatPreisTyp%
'Public MarsAktiv%

Public RezApoNr$, OrgRezApoNr$, RezApoNrPraefix$, RezApoName$(1), RezApoDruckName$, BtmRezDruckName$
Public ZuzFeld!(7)

Public ActBundesland%, ActKasse%, ActVerordnung%, ActVebNr&, ActPauschaleNr&, OrgBundesland%
Public BeigetreteneVereinbarungen$

Public PznAuswahlWert$
Public PznAuswahlName$

Public DruckSeite%

Public RezeptFarben&(3)
Public MagDarstellung&(6, 1)

Public FormErg%
Public FormErgTxt$

Public VmRabattFaktor#

Public OptionenNeu%

Public RezeptDrucker$, StandardDrucker$
Public RezeptDruckerPara$
Public IstDosDrucker%

Public EingabeStr$

Public EasyMatchModus%

Public ImpfstoffeDa%

Public Taetigkeiten(1) As TaetigkeitenStruct
Public AnzTaetigkeiten%

Public SonderBelege(10) As SonderBelegeStruct
Public AnzSonderBelege%

Public RezepturMitFaktor%
Public BtmAlsZeile%
Public BtmFaktorManuell%
Public RezepturDruck%
Public AvpTeilnahme%
Public FormatPrivatRezept%
Public DatumObenPrivatRezept%
    
Public SonderBelegRezept%

Public IdentPzn$
    
Public AbholerDB As Database
Public AbholerNummerRec As Recordset
Public AbholerDetailRec As Recordset
Public AbholerInfoRec  As Recordset
Public AbholerMdb%

Public AbholerSQL%
Public AbholerSqlDatabase$
Public AbholerConn As New ADODB.Connection
Public AbholerNummerAdoRec As New ADODB.Recordset
Public AbholerDetailAdoRec As New ADODB.Recordset
Public AbholerInfoAdoRec  As New ADODB.Recordset
Public AnfMagAdoRec  As New ADODB.Recordset
Public ParameterAdoRec As New ADODB.Recordset

Public AutIdemIk&, AutIdemKbvNr&

Public InhaltsstoffeModus%

Public SQLStr$

Public ParenteralPzn$(31)
Public ParenteralTxt$(31)
Public ParenteralPreis#(31)
Public ParenteralRezept%, Parenteral_AOK_LosGebiet%, Parenteral_AOK_NordOst%
Public ParenteralPara$, ParenteralHash$
Public ParEnteralAufschlag#(1)
Public ParEnteralPrimärPackmittel%
Public ParEnteralAI%
Public ParEnteralAnzEinheiten#
Public ParEnteralHerstellerKey%
Public ParEnteralHerstellerKz$(4)
Public pCharge$
Public HashErstellDat As Date
Public PreisKz_62_70%

Public FiveRxFlag%

Public RezeptMitHilfsmittelNr%

Public AuseinzelungPzn$(1)
Public AuseinzelungFaktor&(1)
Public AuseinzelungPreis#(1), AuseinzelungPreisGesamt#
Public AuseinzelungBtm%

Public PreisDiffAktiv%
Public PreisDiffStr$
Public KundenBoxX&, KundenBoxY&

Public WinVkId&, WinVkX&, WinVkY&, WinVkAnzahl&
Public WinVkDatum As Date
Public WinVkText$

Public WinRezDebugAktiv%

Public AbrechnungsVerfahren_2_Aktiv%

Public Chef As Boolean

Public AlteTaxeAktiv As Boolean

Public HmAbrechnungsKz$, Druck_HmAbrechnungsKz$, LEGS$

Public BtmDerivatMengeIgnorieren%

Public AutIdemKreuz0%

Public AlleRezepturenMitHash As Boolean

Public ChefModus As Boolean

Public LennartzPfad$
Public TA1_V37%

Dim AnzPicGdi%

Public FD_OP As TI_Back.Fachdienst_OP
Public eRezeptTaskId$, eRezeptBundleKBV$, eRezeptStatusWerte$, eRezeptPzn$, eRezeptAMverfügbarkeit$, eRezeptOpStatus%, eRezeptGespeichert%, eRezeptSonderPZN%, eRezeptVerordnungsTyp%
Public eRezeptNoctu#, eRezeptBotendienst#, eRezeptBeschaffungsKosten#
Public FDok%

Public RezeptUnique$

Public eRezeptListe$

Public PharmDienstleistungenPzn$(29)
Public PharmDienstleistungenTxt$(29)

Public ImpfLeistungenPzn$(1)
Public ImpfLeistungenTxt$(1)
Public ImpfLeistungModus As String

Private Const DefErrModul = "WINREZK.BAS"

Function InitProgramm%(Optional Visible% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitProgramm%")
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
Dim i%, auf%, sMax%
Dim h$

Select Case (ProgrammChar$)
    Case "R"
        
        Call StopAnimation(frmAction)
        Call frmAction.WechselModus(0)
        
End Select

AnzPicGdi = 1
AlteTaxeAktiv = False

InitProgramm% = True
Call DefErrPop
End Function

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
Dim i%

ProgrammNamen$(0) = "Rezeptkontrolle"
ProgrammNamen$(1) = "Auswertung Rezeptspeicher"

WochenTag$(0) = "Montag"
WochenTag$(1) = "Dienstag"
WochenTag$(2) = "Mittwoch"
WochenTag$(3) = "Donnerstag"
WochenTag$(4) = "Freitag"
WochenTag$(5) = "Samstag"
WochenTag$(6) = "Sonntag"
    
ParenteralPzn$(0) = "9999092"
ParenteralTxt$(0) = "Zytostatika - Zubereitungen"
ParenteralPreis#(0) = 53
ParenteralPzn$(1) = "9999100"
ParenteralTxt$(1) = "indiv. hergest. Mittel zur parenteralen Ernährung"
ParenteralPreis#(1) = 50
ParenteralPzn$(2) = "9999123"
ParenteralTxt$(2) = "indiv. hergest. parent. antibiotikahaltige Infusionslösungen"
ParenteralPreis#(2) = 30
ParenteralPzn$(3) = "9999169"
ParenteralTxt$(3) = "indiv. hergest. parent. virustatikahaltige Infusionslösungen"
ParenteralPreis#(3) = 30
ParenteralPzn$(4) = "9999146"
ParenteralTxt$(4) = "indiv. hergest. Schmerzlösungen"
ParenteralPreis#(4) = 30
ParenteralPzn$(5) = "9999152"
ParenteralTxt$(5) = "sonstige indiv. hergest. parenterale Lösungen"
ParenteralPreis#(5) = 40
ParenteralPzn$(6) = "2567478"   '"9999092"
ParenteralTxt$(6) = "Lösungen mit monoklonalen Antikörpern"
ParenteralPreis#(6) = 67
ParenteralPzn$(7) = "2567461"  '"9999152"
ParenteralTxt$(7) = "Calciumfolinatlösungen"
ParenteralPreis#(7) = 39
'ParenteralPzn$(8) = "9999092"
'ParenteralTxt$(8) = "Zytostatika - Zubereitungen privat"
For i = 8 To 15
    ParenteralPzn$(i) = ParenteralPzn$(i - 8)
    ParenteralTxt$(i) = ParenteralTxt$(i - 8) + "  (PRIVAT)"
    ParenteralPreis#(i) = ParenteralPreis#(i - 8)
Next i

ParenteralPzn$(16) = "9999086"
ParenteralTxt$(16) = "Methadon-Lösungen (FAM)"
ParenteralPzn$(17) = "06461506"
ParenteralTxt$(17) = "Methadon-Lösungen (Wirkstoff)"
ParenteralPzn$(18) = "2567107"
ParenteralTxt$(18) = "Levomethadon-Zubereitungen (FAM)"
ParenteralPzn$(19) = "06461512"
ParenteralTxt$(19) = "Levomethadon-Zubereitungen (Wirkstoff)"
ParenteralPzn$(20) = "2567113"
ParenteralTxt$(20) = "Buprenorphin-Einzeldosen (Take Home)"
ParenteralPzn$(21) = "2567136"
ParenteralTxt$(21) = "Suboxone-Einzeldosen (Take Home)"
ParenteralPzn$(22) = "2567114"
ParenteralTxt$(22) = "Subutex-Einzeldosen (KEIN Aut-Idem)"
ParenteralPzn$(23) = "2567115"
ParenteralTxt$(23) = "Subutex-Einzeldosen (Take Home)"

ParenteralPzn$(24) = "06460665"
ParenteralTxt$(24) = "Cannabisblüten in Zubereitungen"
ParenteralPzn$(25) = "6460694"
ParenteralTxt$(25) = "Cannabisblüten unverändert"

ParenteralPzn$(26) = "06461446"
ParenteralTxt$(26) = "Cannabisblüten in Zubereitungen (BfArM)"
ParenteralPzn$(27) = "06461423"
ParenteralTxt$(27) = "Cannabisblüten unverändert (BfArM)"

ParenteralPzn$(28) = "06460748"
ParenteralTxt$(28) = "Cannabisextrakt in Zubereitungen"
ParenteralPzn$(29) = "06460754"
ParenteralTxt$(29) = "Cannabisextrakt unverändert"
ParenteralPzn$(30) = "06460749"
ParenteralTxt$(30) = "Dronabinol in Zubereitungen"

ParenteralPzn$(31) = "09999011"
ParenteralTxt$(31) = "Cannabis-Rezeptur BG"


ParEnteralAufschlag#(0) = 3
ParEnteralAufschlag#(1) = 15

PharmDienstleistungenPzn(0) = "17716808"
PharmDienstleistungenTxt(0) = "Erweiterte Medikationsberatung von Patienten mit Polymedikation"
PharmDienstleistungenPzn(1) = "17716814"
PharmDienstleistungenTxt(1) = "Erweiterte Medikationsberatung von Patienten mit Polymedikation (Umstellung vor 12-Monats-Frist)"
PharmDienstleistungenPzn(2) = "17716843"
PharmDienstleistungenTxt(2) = "Pharmazeutische Betreuung nach Organtransplantation"
PharmDienstleistungenPzn(3) = "17716866"
PharmDienstleistungenTxt(3) = "Pharmazeutische Betreuung nach Organtransplantation (Follow-up-Gespräch)"
PharmDienstleistungenPzn(4) = "17716820"
PharmDienstleistungenTxt(4) = "Pharmazeutische Betreuung unter oraler Antitumortherapie"
PharmDienstleistungenPzn(5) = "17716837"
PharmDienstleistungenTxt(5) = "Pharmazeutische Betreuung unter oraler Antitumortherapie (Follow-up-Gespräch)"
PharmDienstleistungenPzn(6) = "17716872"
PharmDienstleistungenTxt(6) = "Standardisierte Risikoerfassung bei Bluthochdruck-Patienten"
PharmDienstleistungenPzn(7) = "17716783"
PharmDienstleistungenTxt(7) = "Erweiterte Einweisung in die korrekte Arzneimittelanwendung und Üben der Inhalationstechnik"

ImpfLeistungenPzn(0) = "17716926"
ImpfLeistungenTxt(0) = "Impfleistung: Grippe"
ImpfLeistungenPzn(1) = "17717400"
ImpfLeistungenTxt(1) = "Impfleistung: Corona"

Call DefErrPop
End Sub

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
Dim i%, erg%, tmp%, ind%, t%, m%, j%, c%
Dim l&, hWnd&
Dim h$, SQLStr$, DirRet$, FuDate$, CdInfoVer$, s$
Dim WinText As String * 255
Dim a!
Dim dat1 As Date
Dim AiTd As TableDef
Dim AvTestRec As New ADODB.Recordset

ChDrive "Z"
'ChDrive "C"
Call InitMisc
If (App.PrevInstance) Then
    hWnd& = GetForegroundWindow()
    hWnd& = GetWindow(hWnd&, GW_HWNDFIRST)
    Do Until (hWnd& = 0)
        l& = GetWindowText(hWnd&, WinText, 255)
        h$ = Left$(WinText, l&)
        For i% = 0 To 1
            j% = Len(h$)
            m% = Len(ProgrammNamen$(i%))
            If (j% > m%) Then
                s$ = Left$(h$, m%)
                If (s$ = ProgrammNamen$(i%)) Then
                    AppActivate h$
                    End
                End If
            End If
        Next i%
        hWnd& = GetWindow(hWnd&, GW_HWNDNEXT)
    Loop
    End
End If

FDok = 0

AbrechnungsVerfahren_2_Aktiv = (Format(Now, "YYYYMMDD") >= "20160701")
If (UCase(Command) = "ABR2") Then
    AbrechnungsVerfahren_2_Aktiv = True
End If

If (Dir$("fistam.dat") = "") Then ChDir "\user"
INI_DATEI = CurDir + "\winop.ini"

Set para = New clsOpPara
Set wpara = New clsWinPara
Call wpara.HoleWindowsParameter


h$ = Space$(100)
l& = GetPrivateProfileString("Rezeptkontrolle", "ClassicLine", h$, h$, 101, INI_DATEI)
h$ = Trim(Left$(h$, l&))
If (h$ <> "") Then
    h$ = "," + h$ + ","
    para.Newline = Not (InStr(h, "," + para.User + ",") > 0)
End If
    

MarsModus = 0
'If (para.MARS) Then
'    MarsModus = MARS_REZEPT_DRUCK
'    If (Command = "KONTROLLE") Then
'        MarsModus = MARS_REZEPT_KONTROLLE
'    End If
'End If

'If (SqlInit = 0) Then
'    End
'End If
Set sqlop = New clsSqlTools
If (sqlop.SqlInit = 0) Then
    End
End If

Set ast = New clsStamm
Set ass = New clsStatistik
Set arttext = New clsArttext
Set hTaxe = New clsHilfsTaxe
Set nnek = New clsNNEK
Set Artikel = New clsArtikelDB
Set Hilfstaxe = New clsHilfstaxeDB
'Set taxe = New clsTaxe
Set taxeAdoDB = New clsTaxeAdoDB
Set kKassenDB = New clsKkassenDB
Set AplusVDB = New clsAplusVDB
Set AutIdemDB = New clsAutIdemDB
Set lif = New clsLieferanten
Set lifzus = New clsLiefZusatz
Set Lieferanten = New clsLieferantenDB
Set ww = New clsWawiDat
Set BESORGT = New clsBesorgt
Set Merkzettel = New clsMerkzettel
Set Kiste = New clsKiste
Set vk = New clsVerkauf
Set VerkaufAdoDB = New clsVerkaufDB
Set RezTab = New clsVerkRtab
Set VmPzn = New clsVmPzn
Set VmBed = New clsVmBed
Set VmRech = New clsVmRech


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
        Call Shell("WinRezK2.exe", vbNormalFocus)
        End
    End If
    AvTestRec.Close
End If


Call para.HoleFirmenStamm

frmAction.Show

Call StartAnimation(frmAction, "Parameter werden eingelesen ...")

'Call para.HoleFirmenStamm
Call para.AuslesenPdatei
Call para.HoleZuzahlungen
Call para.EinlesenPersonal

If (InStr(para.Benutz, "r") <= 0) Then
    ind% = DateDiff("d", Now, "30.06.2003")
    If (ind% < 0) Then
        Call iMsgBox("Dieses Programm hat Ihre Apotheke nicht gekauft !", vbCritical)
        wpara.ExitEndSub
        Call frmAction.frmActionUnload
        End
        Call DefErrPop: Exit Sub
    End If
End If

ActBenutzer% = HoleActBenutzer%

'h$ = "0"
'l& = GetPrivateProfileString("Mars", "Mars", "0", h$, 2, "\user\vkpara.ini")
'h$ = Left$(h$, l&)
'MarsAktiv = (Val(h$) = 1)

Chef = True
If (para.MARS) Then
    Chef = False
    If ActBenutzer% = 1 Then Chef = True
    If Not Chef Then
        If (para.Newline) Then
            ActBenutzer = HoleBenutzerSignatur
        Else
            ActBenutzer = 0
        End If
    End If
    If ActBenutzer <= 0 Then
        Call iMsgBox("Kein gültiges Passwort eingegeben !", vbCritical)
        wpara.ExitEndSub
        Call frmAction.frmActionUnload
        End
        Call DefErrPop: Exit Sub
    End If
    If ActBenutzer% = 1 Then Chef = True
    
    MarsModus = MARS_REZEPT_DRUCK
'    If (Chef) Then
        If (MessageBox("Programm im Modus 'Rezept-KONTROLLE' starten?" + vbCrLf + vbCrLf + "(Bei NEIN wird das Programm im Modus 'Rezept-DRUCK' ausgeführt)", vbQuestion Or vbYesNo, "Rezeptkontrolle") = vbYes) Then
            MarsModus = MARS_REZEPT_KONTROLLE
        End If
'    End If
    With frmAction.lblMarsModus
        If (MarsModus = MARS_REZEPT_KONTROLLE) Then
            .Caption = "Rezept-KONTROLLE"
        Else
            .Caption = "Rezept-DRUCK"
 '           Chef = False
        End If
    End With
    If (para.Newline) And (MarsModus = MARS_REZEPT_KONTROLLE) Then
        Call frmAction.MarsRezeptKontrolleIcon
    End If
        
        
'    If Not Chef Then
'      Call ProgrammEnde
'      End
'    End If
Else
    frmAction.mnuBearbeitenInd(MENU_F7).Caption = ""
    frmAction.mnuBearbeitenInd(MENU_F8).Caption = ""
End If

'ast.OpenDatei
'ass.OpenDatei
'arttext.OpenDatei
'hTaxe.OpenDatei
'nnek.OpenDatei
''BESORGT.OpenDatei
'erg% = OpenCreateMerkzettelDB%
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

'lif.OpenDatei
'lifzus.OpenDatei
'LieferantenDbOk% = OpenLieferantenDB%
'If (LieferantenDbOk% = 0) Then
'    End
'End If
LieferantenDbOk = 0
If (Lieferanten.DBvorhanden) Then
    LieferantenDbOk = Lieferanten.OpenDB
End If
If (LieferantenDbOk% = 0) Then
    End
End If



'VerkaufDbOk% = (Dir("Verkauf.mdb") <> "")
'''VerkaufDbOk% = 0
'If (VerkaufDbOk) Then
'    Set VerkaufDB = OpenDatabase("Verkauf.mdb", False, True)
'    Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
'    VerkaufRec.index = "Unique"
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
'End If
If (AutIdemDB.DBvorhanden) Then
    AutIdemOk = AutIdemDB.OpenDB
    AutIdemSonderregelOk% = True
End If

kKassenOk% = False
'h$ = para.TaxeLw + ":\taxe\kkassen.mdb"
'If (Dir$(h$) <> "") Then
'    kKassenOk% = True
'    Set kKassenDB = OpenDatabase(h$, False, True)
'End If
If (kKassenDB.DBvorhanden) Then
    kKassenOk = kKassenDB.OpenDB
End If

Dim sSqlDatabase$
sSqlDatabase = "ABDA_Komplett"
ABDA_Komplett_SQL = sqlop.SqlCheckDatabase(sSqlDatabase)
If (ABDA_Komplett_SQL = 0) Then
    Call MessageBox("Problem: Datenbank 'ABDA_KOMPLETT' nicht vorhanden !", vbCritical, "Rezeptkontrolle")
    End
Else
    Dim ErrNumber&
    h = sqlop.SqlConnectionString(ABDA_Komplett_SQL)
    ABDA_Komplett_Conn.ConnectionString = Left(h, Len(h) - 1) + sSqlDatabase + "; Data Source=" + sqlop.SqlServer(ABDA_Komplett_SQL)

    ABDA_Komplett_Conn.CursorLocation = adUseClient
    ABDA_Komplett_Conn.CommandTimeout = 300
    On Error Resume Next
    Err.Clear
    ABDA_Komplett_Conn.Open
    ErrNumber = Err.Number
    On Error GoTo DefErr
    erg = (ErrNumber = 0)
    If (erg = 0) Then
        Call MessageBox("Problem: Fehler " + CStr(ErrNumber) + "beim Öffnen der Datenbank 'ABDA_KOMPLETT' !", vbCritical, "Rezeptkontrolle")
        End
    End If
End If



'ww.OpenDatei
'If (ww.DateiLen = 0) Then
'    ww.erstmax = 0
'    ww.erstlief = 0
'    ww.erstcounter = 0
'    ww.erstrest = String(ww.DateiLen, 0)
'    ww.PutRecord (1)
'End If


'AplusVOk% = False
''h$ = para.TaxeLw + ":\taxe\AplusV.mdb"
''If (Dir$(h$) <> "") Then
''    AplusVOk% = True
''    Set AplusVDB = OpenDatabase(h$, False, True)
''Else
''    VmPzn.OpenDatei
''    VmBed.OpenDatei
''    VmRech.OpenDatei
''End If
'If (AplusVDB.DBvorhanden) Then
'    AplusVOk = AplusVDB.OpenDB
'End If

TaxmusterDBok = (Dir(TAXMUSTER_DB) <> "")

erg% = ActProgram.RezKontrInit%
If (erg% = False) Then Call DefErrPop: Exit Sub

erg% = InitProgramm%

Call DefErrPop: Exit Sub
    
ErrorHandler:
    erg% = Err
    If ((erg% > 0) And (erg% <> 3024) And (erg% <> 3044)) Then
        Call iMsgBox("Fehler" + Str$(Err) + " beim Öffnen der Taxe " + h$ + vbCr + Err.Description, vbCritical, "OpenDatabase")
        End
    End If
    Err = 0
    Resume Next
    Return

Call DefErrPop
End Sub

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

'FabsErrf% = clsfabs.Cmnd("K\20", FabsRecno&)
'If (FabsErrf%) Then
'    erg% = MsgBox("Fehler " + Str$(FabsErrf%) + " beim Schliessen der Statistik-Daten", vbInformation)
'End If
    
wpara.ExitEndSub

'TaxeDB.Close
taxeAdoDB.CloseDB

If (ArtikelDbOk%) Then
Else
'    If (InStr(para.Benutz, "t") > 0) Then
'        On Error Resume Next
'        MerkzettelDB.Close
'        On Error GoTo DefErr
'    Else
'        BESORGT.CloseDatei
'    End If
End If

If (AutIdemOk%) Then
'    AutIdemDB.Close
    AutIdemDB.CloseDB
End If

If (kKassenOk%) Then
'    kKassenDB.Close
    kKassenDB.CloseDB
End If

If (ArtikelDbOk%) Then
    Artikel.CloseDB
    Hilfstaxe.CloseDB
Else
    ast.CloseDatei
    ass.CloseDatei
    arttext.CloseDatei
    hTaxe.CloseDatei
    nnek.CloseDatei
End If

'taxe.CloseDatei
'ww.CloseDatei
'BESORGT.CloseDatei

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

If (LieferantenDbOk%) Then
'    LieferantenConn.Close
    Lieferanten.CloseDB
Else
    lif.CloseDatei
    lifzus.CloseDatei
''    lieftext.CloseDatei
End If

If (AplusVOk) Then
'    AplusVDB.Close
    AplusVDB.CloseDB
'Else
'    VmPzn.CloseDatei
'    VmBed.CloseDatei
'    VmRech.CloseDatei
End If

If (ABDA_Komplett_SQL) Then
    ABDA_Komplett_Conn.Close
    Set ABDA_Komplett_Conn = Nothing
End If

If (ARBEMB% > 0) Then Call iClose(ARBEMB%)

If (RezSpeicherOK%) Then
    RezSpeicherDB.Close
    KassenDB.Close
End If
If (para.MARS) Then
    MarsRezSpeicherDB.Close
End If
    
ast.FreeClass
ass.FreeClass
'taxe.FreeClass
lif.FreeClass
lifzus.FreeClass
ww.FreeClass
BESORGT.FreeClass
'kiste.FreeClass
arttext.FreeClass
vk.FreeClass
RezTab.FreeClass
VmPzn.FreeClass
VmBed.FreeClass
VmRech.FreeClass
nnek.FreeClass
hTaxe.FreeClass

Set ast = Nothing
Set ass = Nothing
'Set taxe = Nothing
Set lif = Nothing
Set lifzus = Nothing
Set ww = Nothing
Set BESORGT = Nothing
Set Merkzettel = Nothing
Set Kiste = Nothing
Set arttext = Nothing
Set para = Nothing
Set wpara = Nothing
Set vk = Nothing
Set RezTab = Nothing
Set VmPzn = Nothing
Set VmBed = Nothing
Set VmRech = Nothing
Set nnek = Nothing
Set hTaxe = Nothing

Call frmAction.frmActionUnload
    
End
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
    ret% = MessageBox(prompt$, buttons%, title$)
Else
    ret% = MessageBox(prompt$, buttons%)
End If
KeinRowColChange% = OrgKeinRowColChange%

iMsgBox% = ret%

Call DefErrPop
End Function

Function FileOpen%(Fname$, fAttr$, Optional modus$ = "B", Optional SATZLEN% = 100)
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
        Open Fname$ For Binary Access Read Shared As #Handle%
    Else
        Open Fname$ For Random Access Read Shared As #Handle% Len = SATZLEN%
    End If
    If (Err = 0) Then
        If (LOF(Handle%) = 0) Then
            Close #Handle%
            Kill (Fname$)
            Err.Raise 53
        Else
            Call iLock(Handle%, 1)
            Call iUnLock(Handle%, 1)
        End If
    End If
ElseIf (fAttr$ = "W") Then
    If (modus$ = "B") Then
        Open Fname$ For Binary Access Write As #Handle%
    Else
        Open Fname$ For Random Access Write As #Handle% Len = SATZLEN%
    End If
ElseIf (fAttr$ = "RW") Then
    If (modus$ = "B") Then
        Open Fname$ For Binary Access Read Write Shared As #Handle%
    Else
        Open Fname$ For Random Access Read Write Shared As #Handle% Len = SATZLEN%
    End If
    Call iLock(Handle%, 1)
    Call iUnLock(Handle%, 1)
ElseIf (fAttr$ = "I") Then
    Open Fname$ For Input Access Read Shared As #Handle%
ElseIf (fAttr$ = "O") Then
    Open Fname$ For Output Access Write Shared As #Handle%
End If

If (Err = 0) Then
    FileOpen% = Handle%
Else
    Call iMsgBox("Fehler" + Str$(Err) + " beim Öffnen von " + Fname$ + vbCr + Err.Description, vbCritical, "FileOpen")
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
Dim FüllMenge As Single

menge = Val(TaxMenge)
mal = InStr(TaxMenge, "X")
FüllMenge = menge
While mal > 0 And mal < Len(TaxMenge)
  NextMal = InStr(mal + 1, TaxMenge, "X")
  If NextMal = 0 Then NextMal = Len(TaxMenge) + 1
  FüllMenge = Val(Mid(TaxMenge, mal + 1, NextMal - mal - 1))
  menge = menge * FüllMenge
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

'Function OpenCreateMerkzettelDB%()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("OpenCreateMerkzettelDB%")
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
'Dim ret%, DbNeu%, Max%, OpenErg%, ErrNumber%
'Dim i&, DateiMax&
'Dim DBname$, s$
'Dim Td As TableDef
'Dim idx As index
'Dim fld As Field
'Dim ixFld As Field
'
'ret% = True
'DbNeu% = False
'
'If (InStr(para.Benutz, "t") > 0) Then
'    DBname$ = Merkzettel.DateiName
'    On Error Resume Next
'    Err.Clear
'    Set MerkzettelDB = OpenDatabase(DBname$, False, False)
'    ErrNumber% = Err.Number
'    On Error GoTo DefErr
'    If (ErrNumber% > 0) Then
'        s$ = "ACHTUNG:" + vbCrLf + vbCrLf
'        s$ = s$ + "MERKZETTEL-Datenbank NICHT bereit!" + vbCrLf + vbCrLf
'        s$ = s$ + "Kontaktieren Sie bitte UMGEHEND die Hotline der Firma OPTIPHARM"
'        Call MessageBox(s$, vbOKOnly Or vbCritical)
'        Call DefErrPop: Exit Function
'
'        DbNeu% = True
'
'        If Dir(DBname$) <> "" Then Kill DBname$
'        Set MerkzettelDB = CreateDatabase(DBname$, dbLangGeneral)
'
'    'Tabelle Merkzettel
'        Set Td = MerkzettelDB.CreateTableDef("Merkzettel")
'
'        Set fld = Td.CreateField("Pzn", dbText)
'        fld.AllowZeroLength = False
'        fld.Size = 7
'        Td.Fields.Append fld
'
'        Set fld = Td.CreateField("Txt", dbText)
'        fld.AllowZeroLength = False
'        fld.Size = 36
'        Td.Fields.Append fld
'
'        Set fld = Td.CreateField("BestellDatum", dbDate)
'        Td.Fields.Append fld
'
'        Set fld = Td.CreateField("LieferDatum", dbDate)
'        Td.Fields.Append fld
'
'        Set fld = Td.CreateField("Lief", dbInteger)
'        Td.Fields.Append fld
'
'        Set fld = Td.CreateField("Lm", dbInteger)
'        Td.Fields.Append fld
'
'        Set fld = Td.CreateField("Loesch", dbByte)
'        Td.Fields.Append fld
'
'        ' Indizes für Merkzettel
'        Set idx = Td.CreateIndex()
'        idx.Name = "Pzn"
'        idx.Primary = False
'        idx.Unique = False
'        Set ixFld = idx.CreateField("Pzn")
'        idx.Fields.Append ixFld
'        Td.Indexes.Append idx
'
'        Set idx = Td.CreateIndex()
'        idx.Name = "Txt"
'        idx.Primary = False
'        idx.Unique = False
'        Set ixFld = idx.CreateField("Txt")
'        idx.Fields.Append ixFld
'        Set ixFld = idx.CreateField("Pzn")
'        idx.Fields.Append ixFld
'        Set ixFld = idx.CreateField("BestellDatum")
'        idx.Fields.Append ixFld
'        Td.Indexes.Append idx
'
'        MerkzettelDB.TableDefs.Append Td
'
'
'    'Tabelle Parameter
'        Set Td = MerkzettelDB.CreateTableDef("Parameter")
'
'        Set fld = Td.CreateField("Name", dbText)
'        fld.AllowZeroLength = False
'        fld.Size = 30
'        Td.Fields.Append fld
'
'        Set fld = Td.CreateField("Wert", dbInteger)
'        Td.Fields.Append fld
'
'        ' Indizes für Parameter
'        Set idx = Td.CreateIndex()
'        idx.Name = "Name"
'        idx.Primary = True
'        idx.Unique = False
'        Set ixFld = idx.CreateField("Name")
'        idx.Fields.Append ixFld
'        Td.Indexes.Append idx
'
'        MerkzettelDB.TableDefs.Append Td
'
'        MerkzettelDB.Close
'    End If
'    On Error GoTo DefErr
'
'    OpenErg% = Merkzettel.OpenDatenbank("", False, False, MerkzettelDB)
'    Set MerkzettelRec = MerkzettelDB.OpenRecordset("Merkzettel", dbOpenTable)
'    Set MerkzettelParaRec = MerkzettelDB.OpenRecordset("Parameter", dbOpenTable)
'    MerkzettelParaRec.index = "Name"
'
'    If (DbNeu%) Then
'        BESORGT.OpenDatei
'
'        DateiMax& = (BESORGT.DateiLen / BESORGT.RecordLen) - 1
'        For i& = 1 To DateiMax&
'            BESORGT.GetRecord (i& + 1)
'            If (InStr("* ", BESORGT.flag) > 0) And (BESORGT.dat > 0) Then
'                MerkzettelRec.AddNew
'
'                MerkzettelRec!pzn = BESORGT.pzn
'                MerkzettelRec!txt = BESORGT.text
'                MerkzettelRec!BestellDatum = MakeDateStr$(BESORGT.dat)
'
'                If (BESORGT.dt > 0) Then
'                    MerkzettelRec!LieferDatum = MakeDateStr$(BESORGT.dt)
'                Else
'                    MerkzettelRec!LieferDatum = "01.01.1980"
'                End If
'
'                MerkzettelRec!lief = BESORGT.lief
'                MerkzettelRec!lm = BESORGT.lm
'
'                MerkzettelRec!loesch = 0
'
'                MerkzettelRec.Update
'            End If
'        Next i&
'
'        BESORGT.CloseDatei
'    End If
'Else
'    On Error Resume Next
'    Kill "merkzett.mdb"
'    On Error GoTo DefErr
'    BESORGT.OpenDatei
'End If
'
'OpenCreateMerkzettelDB% = ret%
'
'Call DefErrPop
'End Function

Function MakeDateStr$(iDat%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MakeDateStr$")
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
Dim ret$, h$

h$ = sDate(iDat%)

ret$ = Left$(h$, 2) + "." + Mid$(h$, 3, 2) + "."
If (Val(Mid$(h$, 5, 2)) > 50) Then
    ret$ = ret$ + "19"
Else
    ret$ = ret$ + "20"
End If
ret$ = ret$ + Mid$(h$, 5, 2)

MakeDateStr$ = ret$

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

Function CheckNullInt%(s As Variant)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CheckNullInt%")
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
Dim ret%

If (IsNull(s)) Then
    ret% = 0
Else
    ret% = s
End If

CheckNullInt% = ret%

Call DefErrPop
End Function

Sub CreateGdiObjects()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("CreateGdiObjects")
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
Dim i%
Dim retvalue As Integer
Dim osinfo As OSVERSIONINFO

osinfo.dwOSVersionInfoSize = 148
osinfo.szCSDVersion = Space$(128)
retvalue = GetVersionExA(osinfo)

If (osinfo.dwPlatformId = 1) Then
    For i = 1 To GDI_OBJEKTE_PRO_REZEPT
        Load frmAction.picGDI(AnzPicGdi)
        AnzPicGdi = AnzPicGdi + 1
    Next i
End If

Call DefErrPop
End Sub
 
Sub WinRezDebug(Optional sDebug$ = " ")
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WinRezDebug")
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
Dim FAKTJOURNALDEB%
Dim s$

If (WinRezDebugAktiv) Then
    s$ = Format(Now, "dd.mm.yyyy") + " " + Format(Now, "hh:nn:ss") + " " + Format(Timer, "0.00") + "   " + sDebug$

    FAKTJOURNALDEB% = FreeFile
    Open "WINREZK.DEB" For Append As #FAKTJOURNALDEB%
    Print #FAKTJOURNALDEB%, s
    Close #FAKTJOURNALDEB%
End If

Call DefErrPop
End Sub


Public Function AbholerConnOpen%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AbholerConnOpen%")
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
Dim ret%
Dim ErrNumber&
Dim h$
 
ret = 0
If (AbholerSQL) Then
    h = sqlop.SqlConnectionString(AbholerSQL)
    AbholerConn.ConnectionString = Left(h, Len(h) - 1) + AbholerSqlDatabase + "; Data Source=" + sqlop.SqlServer(AbholerSQL)

    AbholerConn.CursorLocation = adUseClient
    AbholerConn.CommandTimeout = 300
    On Error Resume Next
    Err.Clear
    AbholerConn.Open
    ErrNumber = Err.Number
    On Error GoTo DefErr
    ret = (ErrNumber = 0)
End If
AbholerConnOpen = ret
        
Call DefErrPop
End Function

Sub PruefZiffer(s$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PruefZiffer")
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
Dim i%, sum%, zw%

sum% = 0
For i% = 1 To 12
    zw% = Val(Mid$(s$, i%, 1))
    sum% = sum% + (zw% * (1 + 2 * ((i% + 1) Mod 2)))
Next i%

zw% = 10 - (sum% Mod 10)
If (zw% = 10) Then zw% = 0

s$ = Left$(s$, 12) + Format(zw%, "0")

Call DefErrPop
End Sub

Public Function Transformieren(strQuelle As String, strXSLT As String, strZiel As String, Optional strFehler As String) As Long
     Dim objQuelle As MSXML2.DOMDocument60
     Dim objXSLT As MSXML2.DOMDocument60
     Dim objZiel As MSXML2.DOMDocument60
     Set objQuelle = New MSXML2.DOMDocument60
     objQuelle.Load strQuelle
     If objQuelle.parseError = 0 Then
         Set objXSLT = New MSXML2.DOMDocument60
         objXSLT.Load strXSLT
         If objXSLT.parseError = 0 Then
             Set objZiel = New MSXML2.DOMDocument60
             objQuelle.transformNodeToObject objXSLT, objZiel
             objZiel.Save strZiel
         Else
             Transformieren = objXSLT.parseError.errorCode
             strFehler = ".xslt-datei: " & vbCrLf & strXSLT & vbCrLf & objXSLT.parseError.reason
         End If
     Else
         Transformieren = objQuelle.parseError.errorCode
         strFehler = "Quelldatei: " & vbCrLf & strQuelle & vbCrLf & objQuelle.parseError.reason
     End If
End Function

Public Function writeOut(cText As String, file As String) As Integer
    On Error GoTo errHandler
    Dim fsT As Object
    Dim tFilePath As String

    tFilePath = file '+ ".txt"

    'Create Stream object
    Set fsT = CreateObject("ADODB.Stream")

    'Specify stream type - we want To save text/string data.
    fsT.Type = 2

    'Specify charset For the source text data.
    fsT.Charset = "utf-8"

    'Open the stream And write binary data To the object
    fsT.Open
    fsT.WriteText cText

    'Save binary data To disk
    fsT.SaveToFile tFilePath, 2

    GoTo finish

errHandler:
    MsgBox (Err.Description)
    writeOut = 0
    Exit Function

finish:
    writeOut = 1
End Function

Function ReadFileToText(filePath)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ReadFileToText")
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
Dim objStream, strData

Set objStream = CreateObject("ADODB.Stream")
objStream.Charset = "utf-8"
objStream.Open
objStream.LoadFromFile (filePath)
strData = objStream.ReadText()
objStream.Close
Set objStream = Nothing

ReadFileToText = strData

Call DefErrPop
End Function

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

Function FiveRxPreis$(dPreis#)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FiveRxPreis$")
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
Dim ind%
Dim ret$

ret$ = Format(dPreis, "0.00")
ind = InStr(ret, ",")
If (ind > 0) Then
    ret = Left$(ret, ind - 1) + "." + Mid$(ret, ind + 1)
End If

FiveRxPreis = ret

Call DefErrPop
End Function


