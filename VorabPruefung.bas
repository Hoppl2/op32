Attribute VB_Name = "modVorabPruefung"
Option Explicit

Private Const DefErrModul = "VORABPRUEFUNG.BAS"

Public XmlResponse$
Dim FiveRxKuNr$, FiveRxApoIk$, FiveRxPasswort$, FiveRxPath$, FiveRxUrl$, FiveRxAnzeigeStatus$, ApoIk$, SQL$, buf$
Dim FiveRxSndId&
Dim vStatus$
Dim FiveRxRec As New ADODB.Recordset

Public bEinzelErgebnis As Boolean

Public FiveRxRezeptStatus$(8)


Function VorabPruefung(sTaskId$, bVorabPruefung As Boolean) As Boolean
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("VorabPruefung")
'Call DefErrMod(DefErrModul)
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
Dim j%, DRUCK_ID%, löschen%, AnzGef%, iWg%, lm%, KundeOk%, ok%, PARAM_DAT%, AnzRezArtikel%, test%, ErsterDurchlauf%
Dim IMAGE_ID%, PrivRezept%
Dim i&, l&, asatz&, ssatz&, lRecno&, AktKuNr&, lKuNr&, lMax&, ErrNumber&, ind&
Dim Prozent!, StartZeit!, Dauer!, GesamtDauer!, RestDauer!
Dim DruckDatei$, pzn$, h$, txt$, LatKey$, OrgXmlRequest$, KundenDBname$
Dim SQLStr As String
'Dim tseRec As New ADODB.Recordset

VorabPruefung = False

vStatus = ""

test = 0 ' 1 '!!!!!!!

'    XmlAlt$(0) = "&": XmlNeu$(0) = "&amp;"
'    XmlAlt$(1) = "'": XmlNeu$(1) = "&apos;"
'    XmlAlt$(2) = "<": XmlNeu$(2) = "&lt;"
'    XmlAlt$(3) = ">": XmlNeu$(3) = "&gt;"
'    XmlAlt$(4) = Chr$(34): XmlNeu$(4) = "&quot;"
'    XmlAlt$(5) = "Ä": XmlNeu$(5) = "&#196;"
'    XmlAlt$(6) = "Ö": XmlNeu$(6) = "&#214;"
'    XmlAlt$(7) = "Ü": XmlNeu$(7) = "&#220;"
'    XmlAlt$(8) = "ä": XmlNeu$(8) = "&#228;"
'    XmlAlt$(9) = "ö": XmlNeu$(9) = "&#246;"
'    XmlAlt$(10) = "ü": XmlNeu$(10) = "&#252;"
'    XmlAlt$(11) = "ß": XmlNeu$(11) = "&#223;"
'    XmlAlt$(12) = "µ": XmlNeu$(12) = "u"
'    XmlAlt$(13) = "§": XmlNeu$(13) = "&#167;"   '"&sect;"
'    XmlAlt$(14) = "‰": XmlNeu$(14) = "&#8240;"   '"&permil;"
'    XmlAlt$(15) = "€": XmlNeu$(15) = "&#8364;"

Call TI_Back_Protokoll(IIf(bVorabPruefung, "VORAB", "EINREICHUNG") + " " + sTaskId)
SQL = "Select * FROM TI_eRezepte"
SQL = SQL + " WHERE TI_eRezepte.TaskId='" + sTaskId + "'"
'Call SQLSelect(FiveRxRec, VKConn, SQL)
FabsErrf = VerkaufAdoDB.OpenRecordset(FiveRxRec, SQL, 0)
'If (FabsErrf <> 0) Then
'    Call iMsgBox("keine passenden Rezepte gespeichert !")
'    Call DefErrPop: Exit Sub
'End If

If FiveRxRec.EOF Then
    Call DefErrPop: Exit Function
End If

h$ = "00001"
l& = GetPrivateProfileString("Allgemein", "SndId", h$, h$, 6, CurDir() + "\fiverx.ini")
FiveRxSndId = Val(Left$(h$, l&))

h$ = "<?xml version='1.0' encoding='utf-8'?>"
h$ = h$ + vbCrLf + "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:SOAP-ENC='http://schemas.xmlsoap.org/soap/encoding/' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema'>"
h$ = h$ + vbCrLf + "  <soapenv:Body>"

If (bVorabPruefung) Then
    h$ = h$ + vbCrLf + "    <m:pruefeRezept xmlns:m='http://fiverx.de/spec/abrechnungsservice/types'>"
    h$ = h$ + vbCrLf + "      <rzePruefung>"
    
    h$ = h$ + vbCrLf + "<![CDATA["
    h$ = h$ + vbCrLf + "        <rzePruefung xmlns='http://fiverx.de/spec/abrechnungsservice'>"
    
    h$ = h$ + vbCrLf + XmlSendHeader("            ", 0)
Else
    h$ = h$ + vbCrLf + "    <m:sendeRezepte xmlns:m='http://fiverx.de/spec/abrechnungsservice/types'>"
    h$ = h$ + vbCrLf + "      <rzeLeistung>"
    
    h$ = h$ + vbCrLf + "<![CDATA["
    h$ = h$ + vbCrLf + "        <rzeLeistung xmlns='http://fiverx.de/spec/abrechnungsservice'>"
    h$ = h$ + vbCrLf + "          <rzLeistungHeader>"
    
    h$ = h$ + vbCrLf + XmlSendHeader("            ", test)
    
    h$ = h$ + vbCrLf + "            <sndId>" + CStr(FiveRxSndId) + "</sndId>"
    h$ = h$ + vbCrLf + "          </rzLeistungHeader>"
End If


Dim XmlRequest As String
XmlRequest = h$
XmlResponse = ""

Dim TagRezept$, ret$, Unique$
Dim tagTransaktionsId$
TagRezept = "eRezept"
tagTransaktionsId = "eRezeptId"

Unique = "12345"
Dim ind2 As Integer
h = Unique$
Do
    ind2 = InStr(h, " ")
    If (ind2 > 0) Then
        h = Left$(h, ind2 - 1) + "0" + Mid$(h, ind2 + 1)
    Else
        Exit Do
    End If
Loop

ret$ = ""

If (bVorabPruefung) Then
    ret$ = ret$ + vbCrLf + "          <rzPruefungBody>"
    '        ret$ = ret$ + vbCrLf + "            <eLeistungHeader>"
    
    'ret$ = ret$ + vbCrLf + "              <avsId>" + Trim(Unique$) + "</avsId>"
    ret$ = ret$ + vbCrLf + "              <avsId>" + h$ + "</avsId>"
    ret$ = ret$ + vbCrLf + "              <pruefModus>" + "SYNCHRON" + "</pruefModus>"
    ret$ = ret$ + vbCrLf + "              <" + TagRezept + ">"
    ret$ = ret$ + vbCrLf + "              <eRezeptId>" + CheckNullStr(FiveRxRec!PrescriptionID) + "</eRezeptId>"
    ret$ = ret$ + vbCrLf + "              <eRezeptData>" + CheckNullStr(FiveRxRec!eDispensierung) + "</eRezeptData>"
    ret$ = ret$ + vbCrLf + "              </" + TagRezept + ">"
    ret$ = ret$ + vbCrLf + "          </rzPruefungBody>"
    
Else
    ret$ = ret$ + vbCrLf + "          <rzLeistungInhalt>"
    ret$ = ret$ + vbCrLf + "            <eLeistungHeader>"
    ret$ = ret$ + vbCrLf + "              <avsId>" + h$ + "</avsId>"
    ret$ = ret$ + vbCrLf + "            </eLeistungHeader>"
    ret$ = ret$ + vbCrLf + "            <eLeistungBody>"
    ret$ = ret$ + vbCrLf + "              <" + TagRezept + ">"
    ret$ = ret$ + vbCrLf + "              <eRezeptId>" + FiveRxRec!PrescriptionID + "</eRezeptId>"
    ret$ = ret$ + vbCrLf + "              <eRezeptData>" + FiveRxRec!eDispensierung + "</eRezeptData>"
    ret$ = ret$ + vbCrLf + "              </" + TagRezept + ">"
    ret$ = ret$ + vbCrLf + "            </eLeistungBody>"
    ret$ = ret$ + vbCrLf + "          </rzLeistungInhalt>"
End If
    
XmlRequest = XmlRequest + ret

If (bVorabPruefung) Then
    h$ = "        </rzePruefung>"
    h$ = h$ + vbCrLf + "]]>"
    h$ = h$ + vbCrLf + "      </rzePruefung>"
    h$ = h$ + vbCrLf + "      <rzeParamVersion>"
    h$ = h$ + vbCrLf + "<![CDATA["
    h$ = h$ + vbCrLf + "        <rzeParamVersion xmlns='http://fiverx.de/spec/abrechnungsservice'>"
    h$ = h$ + vbCrLf + "          <versionNr>01.10</versionNr>"
    h$ = h$ + vbCrLf + "        </rzeParamVersion>"
    h$ = h$ + vbCrLf + "]]>"
    h$ = h$ + vbCrLf + "      </rzeParamVersion>"
    
    h$ = h$ + vbCrLf + "    </m:pruefeRezept>"
    
    h$ = h$ + vbCrLf + "  </soapenv:Body>"
    h$ = h$ + vbCrLf + "</soapenv:Envelope>"
    
Else
    h$ = "        </rzeLeistung>"
    h$ = h$ + vbCrLf + "]]>"
    h$ = h$ + vbCrLf + "      </rzeLeistung>"
    h$ = h$ + vbCrLf + "      <rzeParamVersion>"
    h$ = h$ + vbCrLf + "<![CDATA["
'    h$ = h$ + vbCrLf + "        <rzeParamVersion xmlns='http://fiverx.de/spec/abrechnungsservice'>"
    h$ = h$ + vbCrLf + "        <rzeParamVersion xmlns='http://fiverx.de/spec/abrechnungsservice'>"
    h$ = h$ + vbCrLf + "          <versionNr>01.10</versionNr>"
    h$ = h$ + vbCrLf + "        </rzeParamVersion>"
    h$ = h$ + vbCrLf + "]]>"
    h$ = h$ + vbCrLf + "      </rzeParamVersion>"
    
    h$ = h$ + vbCrLf + "    </m:sendeRezepte>"
    
    h$ = h$ + vbCrLf + "  </soapenv:Body>"
    h$ = h$ + vbCrLf + "</soapenv:Envelope>"
End If

XmlRequest = XmlRequest + h$

'XmlRequest = My.Computer.FileSystem.ReadAllText(CurDir() + "\test_abr.xml", System.Text.Encoding.Default)

Do
    ind = InStr(XmlRequest, "'")
    If (ind > 0) Then
        XmlRequest = Left$(XmlRequest, ind - 1) + Chr(34) + Mid$(XmlRequest, ind + 1)
    Else
        Exit Do
    End If
Loop

Dim FiveRxModus%
FiveRxModus = IIf(bVorabPruefung, 0, 2)
Call TI_Back_Protokoll("VorabPruefung: " + CStr(Len(XmlRequest)))
Call Logbuch(XmlRequest$)
Call WebService(XmlRequest, FiveRxModus)
Call Logbuch(XmlResponse)
Call TI_Back_Protokoll("VorabPruefung Erg: " + CStr(Len(XmlResponse)))

Dim lRecs&
SQL = "UPDATE TI_eRezepte SET eDispensierung=LEN(eDispensierung) WHERE PrescriptionId='" + CheckNullStr(FiveRxRec!PrescriptionID) + "'"
Call VerkaufAdoDB.ActiveConn.Execute(SQL, lRecs&, adExecuteNoRecords)

If (XmlResponse <> "") Then
    'MsgBox (XmlResponse)

    If (bVorabPruefung) Then
    Else
        FiveRxSndId = FiveRxSndId + 1
        l& = WritePrivateProfileString("Allgemein", "SndId", CStr(FiveRxSndId), CurDir() + "\fiverx.ini")
    End If

    Call WebServiceResponse(FiveRxModus)
End If
Call TI_Back_Protokoll("VorabPruefung Erg: " + vStatus)

'    Call Logbuch(XmlRequest$)
    
FiveRxRec.Close

'If (flxSortierung.Rows >= 1) Then
'    frmFiveRxResult.Show 1
'End If

Dim sOkStr$
sOkStr = "VERBESSERBAR ABRECHENBAR HINWEIS LIEFERID"
VorabPruefung = (InStr(sOkStr, vStatus) > 0)
                    
Call DefErrPop
End Function

Function HoleErgebnisse(sTaskId$, Optional bStorno As Boolean = False) As Integer
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleErgebnisse as integer")
'Call DefErrMod(DefErrModul)
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
Dim j%, DRUCK_ID%, löschen%, AnzGef%, iWg%, lm%, KundeOk%, ok%, PARAM_DAT%, AnzRezArtikel%, test%, ErsterDurchlauf%
Dim IMAGE_ID%, PrivRezept%
Dim i&, l&, asatz&, ssatz&, lRecno&, AktKuNr&, lKuNr&, lMax&, ErrNumber&, ind&
Dim Prozent!, StartZeit!, Dauer!, GesamtDauer!, RestDauer!
Dim DruckDatei$, pzn$, h$, txt$, LatKey$, OrgXmlRequest$, KundenDBname$
Dim SQLStr As String
'Dim tseRec As New ADODB.Recordset

HoleErgebnisse = 0 'False

FiveRxRezeptStatus$(0) = "VOR_PRUEFUNG"
FiveRxRezeptStatus$(1) = "FEHLER"
FiveRxRezeptStatus$(2) = "VERBESSERBAR"
FiveRxRezeptStatus$(3) = "HINWEIS"
FiveRxRezeptStatus$(4) = "STORNIERT"
FiveRxRezeptStatus$(5) = "ABRECHENBAR"
FiveRxRezeptStatus$(6) = "VOR_ABRECHNUNG"
FiveRxRezeptStatus$(7) = "ABGERECHNET"
FiveRxRezeptStatus$(8) = "RUECKWEISUNG"


vStatus = ""

test = 0 ' 1 '!!!!!!!

SQL = "Select * FROM TI_eRezepte"
SQL = SQL + " WHERE TI_eRezepte.TaskId='" + sTaskId + "'"
'Call SQLSelect(FiveRxRec, VKConn, SQL)
FabsErrf = VerkaufAdoDB.OpenRecordset(FiveRxRec, SQL, 0)
'If (FabsErrf <> 0) Then
'    Call iMsgBox("keine passenden Rezepte gespeichert !")
'    Call DefErrPop: Exit Sub
'End If

If FiveRxRec.EOF Then
    Call DefErrPop: Exit Function
End If

Dim XmlRequest As String

h$ = "<?xml version='1.0' encoding='utf-8'?>"
h$ = h$ + vbCrLf + "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:SOAP-ENC='http://schemas.xmlsoap.org/soap/encoding/' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema'>"
h$ = h$ + vbCrLf + "  <soapenv:Body>"

If (bStorno) Then
    h$ = h$ + vbCrLf + "    <m:storniereRezept xmlns:m='http://fiverx.de/spec/abrechnungsservice/types'>"
    h$ = h$ + vbCrLf + "      <rzeParamStorno>"
    
    h$ = h$ + vbCrLf + "<![CDATA["
    h$ = h$ + vbCrLf + "        <rzeParamStorno xmlns='http://fiverx.de/spec/abrechnungsservice'>"
    
    h$ = h$ + vbCrLf + XmlSendHeader("          ", test)
    
'    h$ = h$ + vbCrLf + "          <perRezeptID>"
    h$ = h$ + vbCrLf + "              <eRezeptId>" + CheckNullStr(FiveRxRec!PrescriptionID) + "</eRezeptId>"
'    h$ = h$ + vbCrLf + "          </perRezeptID>"
    h$ = h$ + vbCrLf + "        </rzeParamStorno>"
    
    h$ = h$ + vbCrLf + "]]>"
    h$ = h$ + vbCrLf + "      </rzeParamStorno>"
    h$ = h$ + vbCrLf + "      <rzeParamVersion>"
    h$ = h$ + vbCrLf + "<![CDATA["
    h$ = h$ + vbCrLf + "        <rzeParamVersion xmlns='http://fiverx.de/spec/abrechnungsservice'>"
    h$ = h$ + vbCrLf + "          <versionNr>01.10</versionNr>"
    h$ = h$ + vbCrLf + "        </rzeParamVersion>"
    h$ = h$ + vbCrLf + "]]>"
    h$ = h$ + vbCrLf + "      </rzeParamVersion>"
    
    h$ = h$ + vbCrLf + "    </m:storniereRezept>"
Else
    h$ = h$ + vbCrLf + "    <m:ladeStatusRezept xmlns:m='http://fiverx.de/spec/abrechnungsservice/types'>"
    h$ = h$ + vbCrLf + "      <rzeParamStatus>"
    
    h$ = h$ + vbCrLf + "<![CDATA["
    h$ = h$ + vbCrLf + "        <rzeParamStatus xmlns='http://fiverx.de/spec/abrechnungsservice'>"
    
    h$ = h$ + vbCrLf + XmlSendHeader("          ", test)
    
    h$ = h$ + vbCrLf + "          <perRezeptID>"
    h$ = h$ + vbCrLf + "              <eRezeptId>" + CheckNullStr(FiveRxRec!PrescriptionID) + "</eRezeptId>"
    h$ = h$ + vbCrLf + "          </perRezeptID>"
    h$ = h$ + vbCrLf + "        </rzeParamStatus>"
    
    h$ = h$ + vbCrLf + "]]>"
    h$ = h$ + vbCrLf + "      </rzeParamStatus>"
    h$ = h$ + vbCrLf + "      <rzeParamVersion>"
    h$ = h$ + vbCrLf + "<![CDATA["
    h$ = h$ + vbCrLf + "        <rzeParamVersion xmlns='http://fiverx.de/spec/abrechnungsservice'>"
    h$ = h$ + vbCrLf + "          <versionNr>01.10</versionNr>"
    h$ = h$ + vbCrLf + "        </rzeParamVersion>"
    h$ = h$ + vbCrLf + "]]>"
    h$ = h$ + vbCrLf + "      </rzeParamVersion>"
    
    h$ = h$ + vbCrLf + "    </m:ladeStatusRezept>"
End If

h$ = h$ + vbCrLf + "  </soapenv:Body>"
h$ = h$ + vbCrLf + "</soapenv:Envelope>"
XmlRequest = h$

XmlResponse = ""

Dim TagRezept$, ret$, Unique$
Dim tagTransaktionsId$
TagRezept = "eRezept"
tagTransaktionsId = "eRezeptId"

Do
    ind = InStr(XmlRequest, "'")
    If (ind > 0) Then
        XmlRequest = Left$(XmlRequest, ind - 1) + Chr(34) + Mid$(XmlRequest, ind + 1)
    Else
        Exit Do
    End If
Loop

Dim FiveRxModus%
FiveRxModus = IIf(bStorno, 1, 4)

Call Logbuch(XmlRequest$)
Call WebService(XmlRequest, FiveRxModus)
Call Logbuch(XmlResponse)

If (XmlResponse <> "") Then
'    MsgBox (XmlResponse)

    Call WebServiceResponse(FiveRxModus)
End If

'    Call Logbuch(XmlRequest$)
    
FiveRxRec.Close

'If (flxSortierung.Rows >= 1) Then
'    frmFiveRxResult.Show 1
'End If

'Dim sOkStr$
'sOkStr = "VERBESSERBAR ABRECHENBAR HINWEIS"
'HoleErgebnisse = (InStr(sOkStr, vStatus) > 0)
                    
'MsgBox ("vStatus: " + CStr(vStatus))
For i = 0 To UBound(FiveRxRezeptStatus)
    If (vStatus = FiveRxRezeptStatus(i)) Then
        HoleErgebnisse = i + 1
        Exit For
    End If
Next i

Call DefErrPop
End Function

Function XmlSendHeader$(einzug$, Optional test% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("XmlSendHeader$")
'Call DefErrMod(DefErrModul)
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
Dim i%, iVal%, PARAM_DAT%
Dim l&
Dim ret$, h$, h2$, RezApoIkPraefix$

Dim PROG_INI_DATEI As String
PROG_INI_DATEI = CurDir() + "\fiverx.ini"

h$ = "30"
l& = GetPrivateProfileString("Rezeptkontrolle", "InstitutsKzPraefix", h$, h$, 3, CurDir + "\winop.ini")
h$ = UCase(Left$(h$, l&))
iVal = Val(h)
If (iVal <= 0) Or (iVal >= 100) Then
    iVal = 30
End If
RezApoIkPraefix$ = Format(iVal, "00")

h$ = Space$(50)
l& = GetPrivateProfileString("Rezeptkontrolle", "InstitutsKz", h$, h$, 51, CurDir + "\winop.ini")
h$ = Trim(Left$(h$, l&))
If (h$ = "") Then
    PARAM_DAT% = iFileOpen("param.dat", "I")
    For i% = 0 To 22
        If (EOF(PARAM_DAT%)) Then Exit For
        
        Line Input #PARAM_DAT%, h2$
        
        If (i% = 18) Then
            If (h$ = "") Then
                h$ = Trim(h2$)
            End If
            Exit For
        End If
    Next i%
    Close #PARAM_DAT%
End If
If (h$ <> "") Then
    h$ = RezApoIkPraefix + h$
End If
ApoIk$ = h$
    


h$ = Space$(30)
l& = GetPrivateProfileString("Allgemein", "KundenNr", h$, h$, 31, PROG_INI_DATEI)
'h$ = UCase(Left$(h$, l&))
'If (Val(h$) = 0) Then
'    h$ = "99999"
'End If
h$ = Left$(h$, l&)
FiveRxKuNr$ = h$

'h$ = Space$(10)
'l& = GetPrivateProfileString("Allgemein", "ApoIk", h$, h$, 11, PROG_INI_DATEI)
'h$ = Trim(Left$(h$, l&))
'If (h$ <> "") Then
'    FiveRxApoIk = h$
'Else
'    FiveRxApoIk = ApoIk
'End If
FiveRxApoIk = ApoIk


h$ = Space$(50)
l& = GetPrivateProfileString("Allgemein", "Passwort", h$, h$, 51, PROG_INI_DATEI)
FiveRxPasswort$ = Trim(Left$(h$, l&))

h$ = Space$(100)
l& = GetPrivateProfileString("Allgemein", "Url", h$, h$, 101, PROG_INI_DATEI)
h$ = Trim(Left$(h$, l&))
If (h$ = "") Then
    h$ = "http://ws.fiverx.de/axis2/services/FiverxLinkService"
End If
FiveRxUrl$ = h$



ret = einzug + "<sendHeader>"
ret = ret + vbCrLf + einzug + "  <rzKdNr>" + FiveRxKuNr + "</rzKdNr>"
ret = ret + vbCrLf + einzug + "  <avsSw>"
ret = ret + vbCrLf + einzug + "    <hrst>Optipharm GmbH</hrst>"
ret = ret + vbCrLf + einzug + "    <nm>OP</nm>"
ret = ret + vbCrLf + einzug + "    <vs>1.1</vs>"
ret = ret + vbCrLf + einzug + "  </avsSw>"
ret = ret + vbCrLf + einzug + "  <apoIk>" + FiveRxApoIk + "</apoIk>"
ret = ret + vbCrLf + einzug + "  <test>" + CStr(test) + "</test>"
ret = ret + vbCrLf + einzug + "  <pw>" + FiveRxPasswort + "</pw>"
ret = ret + vbCrLf + einzug + "</sendHeader>"

XmlSendHeader = ret

Call DefErrPop
End Function

Sub WebService(XmlRequest$, iModus%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WebService")
'Call DefErrMod(DefErrModul)
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
Dim CrLf%, einzug%
Dim i&, ind&, ind2&, lRecs&
Dim FaultCode$, FaultStr$, AnzeigeStr$
Dim HttpReq As New MSXML2.ServerXMLHTTP60

XmlResponse = ""

HttpReq.Open "POST", FiveRxUrl, False
        
If Not (Err) Then
    'Set a standard SOAP/ XML header for the content-type
    'HttpReq.setRequestHeader "Content-Type", "text/xml"
    Dim h$
    h = "text/xml"
    'h = "soap+xml; charset=utf-8 "
    HttpReq.setRequestHeader "Content-Type", h
    'HttpReq.setRequestHeader "Content-type", "text/xml; charset=UTF-8"
End If

If Not (Err) Then
    HttpReq.Send XmlRequest
End If

If Not (Err) Then
'    SQL = "INSERT INTO TI_FiveRx (AktionInt,Aktion,TaskId,FiveRxXml) VALUES (" + CStr(iModus) + ", '" + "Vorab" + "','" + CheckNullStr(FiveRxRec!PrescriptionID) + "','" + CheckNullStr(FiveRxRec!eAbgabe) + "')"
    
    Dim s2$
    s2 = XmlRequest
    
    ind = InStr(s2, "<eRezeptData>")
    If (ind > 0) Then
        ind2 = InStr(s2, "</eRezeptData>")
        If (ind2 > 0) Then
            s2 = Left(XmlRequest, ind + 12) + CStr(ind2 - ind) + Mid(XmlRequest, ind2)
        End If
    End If
'    SQL = "INSERT INTO TI_FiveRx (AktionInt,Aktion,TaskId,FiveRxXml) VALUES (" + CStr(iModus) + ", '" + "Vorab" + "','" + CheckNullStr(FiveRxRec!PrescriptionID) + "','" + XmlRequest + "')"
    SQL = "INSERT INTO TI_FiveRx (AktionInt,Aktion,TaskId,FiveRxXml) VALUES (" + CStr(iModus) + ", '" + "Vorab" + "','" + CheckNullStr(FiveRxRec!PrescriptionID) + "','" + s2 + "')"
'    VKComm.CommandText = SQL
'    VKComm.CommandTimeout = 300
'    VKComm.Execute
    Call VerkaufAdoDB.ActiveConn.Execute(SQL, lRecs&, adExecuteNoRecords)

    If (iModus = 2) Then
        SQL = "UPDATE TI_eRezepte SET EinreichDatum=GetDate() WHERE PrescriptionId='" + CheckNullStr(FiveRxRec!PrescriptionID) + "'"
        Call VerkaufAdoDB.ActiveConn.Execute(SQL, lRecs&, adExecuteNoRecords)
    End If
    
    buf = HttpReq.responseText
    Do
        ind = InStr(buf, "&amp;")
        If (ind > 0) Then
            buf = Left(buf, ind - 1) + "&" + Mid(buf, ind + 5)
        Else
            Exit Do
        End If
    Loop
    Do
        ind = InStr(buf, "&apos;")
        If (ind > 0) Then
            buf = Left(buf, ind - 1) + "`" + Mid(buf, ind + 6)
        Else
            Exit Do
        End If
    Loop
    Do
        ind = InStr(buf, "&lt;")
        If (ind > 0) Then
            buf = Left(buf, ind - 1) + "<" + Mid(buf, ind + 4)
        Else
            Exit Do
        End If
    Loop
    Do
        ind = InStr(buf, "&gt;")
        If (ind > 0) Then
            buf = Left(buf, ind - 1) + ">" + Mid(buf, ind + 4)
        Else
            Exit Do
        End If
    Loop
    
    ind = 1
    Do
        If (ind > Len(buf)) Then
            Exit Do
        End If
        
        ind = InStr(ind + 2, buf, vbLf)
        If (ind > 0) Then
            If (Mid$(buf, ind - 1, 1) <> vbCr) Then
                buf = Left(buf, ind - 1) + vbCr + Mid(buf, ind)
            End If
        Else
            Exit Do
        End If
    Loop
    
    ind = 1
    Do
        If (ind > Len(buf)) Then
            Exit Do
        End If
        
        ind = InStr(ind + 3, buf, "<")
        If (ind > 0) Then
            CrLf = True
            If (Mid$(buf, ind + 1, 1) = "/") Then
                If (Mid$(buf, ind - 1, 1) <> ">") Then
                    CrLf = 0
                End If
            ElseIf (Mid$(buf, ind - 2, 2) = vbCrLf) Then
                CrLf = 0
            End If
            If (CrLf) Then
                buf = Left(buf, ind - 1) + vbCrLf + Mid(buf, ind)
            End If
        Else
            Exit Do
        End If
    Loop
    
    einzug = 0
    ind = 1
    Do
        If (ind > Len(buf)) Then
            Exit Do
        End If
        
        ind = InStr(ind, buf, vbCrLf)
        If (ind > 0) Then
            If (Mid(buf, ind + 2, 1) = "<") Then
                If (Mid(buf, ind + 3, 1) = "/") Then
                    einzug = einzug - 2
                Else
                    For i = ind To 1 Step -1
                        If (Mid$(buf, i, 1) = "<") Then
                            If (Mid$(buf, i + 1, 1) = "/") Then
                                einzug = einzug - 2
                            ElseIf (Mid$(buf, i + 1, 1) = "?") Then
                                einzug = einzug - 2
                            End If
                            Exit For
                        End If
                    Next i
                    einzug = einzug + 2
                End If
                If (einzug < 0) Then
                    einzug = 0
                End If
                buf = Left(buf, ind + 1) + Space(einzug) + Mid(buf, ind + 2)
            End If
            ind = ind + 5
        Else
            Exit Do
        End If
    Loop
    'Call Logbuch(buf$)
    
    buf = Replace(buf, "'", "''")
'    MsgBox (buf)
    SQL = "INSERT INTO TI_FiveRx (AktionInt,Aktion,TaskId,FiveRxXml) VALUES (" + CStr(iModus + 1) + ", '" + "Vorab_Ret" + "','" + CheckNullStr(FiveRxRec!PrescriptionID) + "','" + buf + "')"
'    VKComm.CommandText = SQL
'    VKComm.CommandTimeout = 300
'    VKComm.Execute

'            Dim ArtIndexDeb%
'            ArtIndexDeb% = FreeFile
'            Open "VorabRet.xml" For Output Access Write Shared As #ArtIndexDeb%
'            Print #ArtIndexDeb%, , buf
'            Close #ArtIndexDeb%

    Call VerkaufAdoDB.ActiveConn.Execute(SQL, lRecs&, adExecuteNoRecords)
    
    XmlResponse = buf   ' UCase(buf)
    
    Dim sAnzeige$, StatusInfo$, fStatus$, fKurzText$, fCode$, fTCode$
    buf = XmlResponse
    
   
'    vStatus = XmlAbschnitt(XmlResponse, "VSTATUS")
'    sAnzeige = vStatus + vbCrLf + vbCrLf
'    Do
'        StatusInfo = XmlAbschnitt(XmlResponse, "STATUSINFO")
'        If (StatusInfo = "") Then
'            Exit Do
'        End If
'        fCode = XmlAbschnitt(StatusInfo, "FCODE")
'        fStatus = XmlAbschnitt(StatusInfo, "FSTATUS")
'        fTCode = XmlAbschnitt(StatusInfo, "FTCODE")
'        fKurzText = XmlAbschnitt(StatusInfo, "FKURZTEXT")
'
'        sAnzeige = sAnzeige + fStatus + vbTab + vbTab + fKurzText + " (" + fCode + "," + fTCode + ")" + vbCrLf
'    Loop
'    Call MsgBox(sAnzeige, , "Anzeige der FiveRx-Rückmeldungen")
'
'
    ind = InStr(buf, "<FAULTCODE>")
    If (ind > 0) Then
        XmlResponse = ""

        ind = InStr(ind + 2, buf, ">")
        buf = Mid$(buf, ind + 1)
        ind = InStr(buf, "</FAULTCODE>")
    End If
    If (ind > 0) Then
        FaultCode = Left$(buf$, ind - 1)
        ind = InStr(buf, "<FAULTSTRING>")
        If (ind > 0) Then
            ind = InStr(ind + 2, buf, ">")
            buf = Mid$(buf, ind + 1)
            ind = InStr(buf, "</FAULTSTRING>")
        End If
        If (ind > 0) Then
            FaultStr = Left$(buf$, ind - 1)
        End If

        AnzeigeStr = "Problem bei der Übertragung:" + vbCrLf + vbCrLf
        AnzeigeStr = AnzeigeStr + "Fehler " + FaultCode
        If (FaultStr <> "") Then
            AnzeigeStr = AnzeigeStr + "  (" + FaultStr + ")"
        End If
        Call MsgBox(AnzeigeStr)
    End If
End If

Call DefErrPop
End Sub

Sub WebServiceResponse(FiveRxModus%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WebServiceResponse")
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
Dim i%, j%, ind%, AbrechnungsStatus%, OrgAbrechnungsStatus%, fCode%, fTCode%, PosNr%, Abgelehnt%, Storno%, fHauptFehler%
Dim Muster16Id&, lRecs&
Dim fWert#
Dim rzLieferId$, h$, h2$, txt$, AbrechnungsStatusStr$, AbrechnungsStatusLangStr$, SQLStr$, StatusInfo$, RezeptUnique$, RezeptNr$
Dim fStatus$, fKommentar$, fKurzText$, fLangText$, fVerbesserung$, OrgXmlResponse$
Dim tagTransaktionsId$, tagStatus$, tagStatus2$, tagId$
Dim valTransaktionsId$, valStatus$, valStatus2$, valId$, ValVerbesserung$, sAnzeige$
Dim iRec As Recordset

rzLieferId = ""
OrgXmlResponse = XmlResponse
XmlResponse = UCase(XmlResponse)

sAnzeige = ""

'MsgBox ("response: " + CStr(FiveRxModus))
   
If (FiveRxModus = 0) Then
    vStatus = XmlAbschnitt(XmlResponse, "STATUS")
    If (vStatus = "") Then
        vStatus = XmlAbschnitt(XmlResponse, "VSTATUS")
    End If
    sAnzeige = vStatus + vbCrLf + vbCrLf
    Do
        StatusInfo = XmlAbschnitt(XmlResponse, "STATUSINFO")
        If (StatusInfo = "") Then
            Exit Do
        End If
        fCode = Val(XmlAbschnitt(StatusInfo, "FCODE"))
        fStatus = XmlAbschnitt(StatusInfo, "FSTATUS")
        fTCode = Val(XmlAbschnitt(StatusInfo, "FTCODE"))
        fKurzText = XmlAbschnitt(StatusInfo, "FKURZTEXT")
        fLangText = XmlAbschnitt(StatusInfo, "FLANGTEXT")
        
        sAnzeige = sAnzeige + fStatus + vbTab + vbTab + fKurzText + " (" + CStr(fCode) + "," + CStr(fTCode) + ")" + vbCrLf
        If (fLangText <> "") Then
            sAnzeige = sAnzeige + "(" + fLangText + ")" + vbCrLf
        End If
        sAnzeige = sAnzeige + vbCrLf
    Loop
    'Call MsgBox(sAnzeige, , "Anzeige der FiveRx-Rückmeldungen")
    
    XmlResponse = OrgXmlResponse
    If (bEinzelErgebnis) Then
        frmFiveRxRueckmeldung.Show 1
    End If
ElseIf (FiveRxModus = 1) Then
    vStatus = XmlAbschnitt(XmlResponse, "ESTORNO")
ElseIf (FiveRxModus = 2) Then
    rzLieferId = XmlAbschnitt(XmlResponse, "RZLIEFERID")
    If (rzLieferId <> "") Then
        tagTransaktionsId = "EREZEPTID"
        
        sAnzeige = "LIEFERID: " + rzLieferId
        
        Do
            valId = XmlAbschnitt(XmlResponse, "ID")
            If (valId = "") Then
                Exit Do
            End If
            
            If (InStr(valId, "ABLEHNUNG") > 0) Then
                h2$ = XmlAbschnitt(valId, tagTransaktionsId)
                Muster16Id = Val(h2$)
                If (Muster16Id > 0) Then
                    fStatus = XmlAbschnitt(valId, "ISTSTATUS")
                    fKommentar = XmlAbschnitt(valId, "FKOMMENTAR")
                
                    sAnzeige = sAnzeige + "ABLEHNUNG" + vbTab + vbTab + fKommentar + " (" + fStatus + ")" + vbCrLf
                End If
            Else
                h2$ = XmlAbschnitt(valId, tagTransaktionsId)
                Muster16Id = Val(h2$)
                If (Muster16Id > 0) Then
'                    SQLStr = "UPDATE TI_eRezepte SET FiveRx_Status=1, FiveRx_Meldung='' WHERE PrescriptionId='" + h2 + "'"
                    SQL = "UPDATE TI_eRezepte SET OpStatus=4 WHERE PrescriptionId='" + h2 + "'"
'                    MsgBox (SQL)
                    Call VerkaufAdoDB.ActiveConn.Execute(SQL, lRecs&, adExecuteNoRecords)
                    
                    vStatus = "LIEFERID" ' rzLieferId
                End If
            End If
        Loop
    End If
    If (bEinzelErgebnis) Then
        Call MsgBox(sAnzeige, , "Anzeige der FiveRx-Rückmeldungen")
    End If
Else
    Dim eRezeptId$

'    MsgBox (XmlResponse)

    tagStatus = "EREZEPTSTATUS"
    tagStatus2 = "STATUS"
    tagTransaktionsId = "EREZEPTID"

    Do
        valStatus = XmlAbschnitt(XmlResponse, tagStatus)
        If (valStatus = "") Then
            Exit Do
        End If

        h2$ = XmlAbschnitt(valStatus, tagTransaktionsId)
        eRezeptId = h2
        If (eRezeptId <> "") Then
            AbrechnungsStatus = 1
            AbrechnungsStatusStr$ = ""
            AbrechnungsStatusLangStr$ = ""

            valStatus2 = XmlAbschnitt(valStatus, tagStatus2)
            AbrechnungsStatusStr$ = valStatus2
            For i = 0 To UBound(FiveRxRezeptStatus)
                If (AbrechnungsStatusStr = FiveRxRezeptStatus(i)) Then
                    AbrechnungsStatus = i + 1
                    Exit For
                End If
            Next i
            vStatus = valStatus2

            For i = 0 To 10
                StatusInfo = XmlAbschnitt(valStatus, "STATUSINFO")
                If (StatusInfo <> "") Then
                    fCode = Val(XmlAbschnitt(StatusInfo, "FCODE"))
                    fStatus = XmlAbschnitt(StatusInfo, "FSTATUS")
                    fKommentar = XmlAbschnitt(StatusInfo, "FKOMMENTAR")
                    fKommentar = Left(fKommentar, 100)

                    fWert = xVal(XmlAbschnitt(StatusInfo, "FWERT"))
                    fTCode = Val(XmlAbschnitt(StatusInfo, "FTCODE"))
                    PosNr = Val(XmlAbschnitt(StatusInfo, "POSNR"))
                    fKurzText = XmlAbschnitt(StatusInfo, "FKURZTEXT")
                    fLangText = XmlAbschnitt(StatusInfo, "FLANGTEXT")
                    If (fLangText <> "") Then
                        fKurzText = fLangText
                    End If
                    fKurzText = Left(fKurzText, 100)
                    fHauptFehler = Abs(XmlAbschnitt(StatusInfo, "FHAUPTFEHLER") = "TRUE")

                    fVerbesserung = ""
                    ValVerbesserung = XmlAbschnitt(StatusInfo, "FVERBESSERUNG")
                    If (ValVerbesserung <> "") Then
                        txt = XmlAbschnitt(ValVerbesserung, "ZUZAHLUNG")
                        If (txt <> "") Then
                            fVerbesserung = fVerbesserung + txt + ".  "
                        End If
                        txt = XmlAbschnitt(ValVerbesserung, "GESBRUTTO")
                        If (txt <> "") Then
                            fVerbesserung = fVerbesserung + txt + ".  "
                        End If
                        txt = XmlAbschnitt(ValVerbesserung, "FAKTOR")
                        If (txt <> "") Then
                            fVerbesserung = fVerbesserung + txt + ".  "
                        End If
                        txt = XmlAbschnitt(ValVerbesserung, "TAXE")
                        If (txt <> "") Then
                            fVerbesserung = fVerbesserung + txt + ".  "
                        End If
                        fVerbesserung = Left(fVerbesserung, 100)
                    End If

                    If (AbrechnungsStatusLangStr$ = "") Then
                        AbrechnungsStatusLangStr$ = fKommentar
                    End If
                    If (AbrechnungsStatusLangStr$ = "") Then
                        AbrechnungsStatusLangStr$ = fKurzText
                    End If
                Else
                    Exit For
                End If
            Next i
            If (AbrechnungsStatusLangStr$ = "") Then
                AbrechnungsStatusLangStr$ = FiveRxRezeptStatus(AbrechnungsStatus - 1)
            End If

            
            If (AbrechnungsStatus = 1) Then
            Else
                Dim OpStatus%
                OpStatus = IIf(AbrechnungsStatus = 2, 5, 6)
                SQL = "UPDATE TI_eRezepte SET OpStatus=" + CStr(OpStatus) + ", FiveRx_Status=" + CStr(AbrechnungsStatus) + ", FiveRx_Meldung='" + SqlString(AbrechnungsStatusLangStr) + "' WHERE PrescriptionId='" + h2 + "'"
'                MsgBox (SQL)
                Call VerkaufAdoDB.ActiveConn.Execute(SQL, lRecs&, adExecuteNoRecords)

            End If
            
            sAnzeige = h2 + vbTab + AbrechnungsStatusStr + vbTab + AbrechnungsStatusLangStr$
'            If (Mid$(FiveRxAnzeigeStatus, AbrechnungsStatus, 1) = "1") Then
'                txt = eRezeptId + vbTab + h2 + vbTab + AbrechnungsStatusStr + vbTab + AbrechnungsStatusLangStr$
'                flxSortierung.AddItem txt$
'            End If
        End If
    Loop
    'Call MsgBox(sAnzeige, , "Anzeige der FiveRx-Rückmeldungen")
    
    XmlResponse = OrgXmlResponse
    If (bEinzelErgebnis) Then
        frmFiveRxRueckmeldung.Show 1
    End If
End If

Call DefErrPop
End Sub

Function XmlAbschnitt$(XmlStr$, SollTag$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("XmlAbschnitt$")
'Call DefErrMod(DefErrModul)
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
Dim ind&, ind2&, ind3&
Dim ret$, sSearch$, sUcaseSearchTag$, ch$, sNameSpace$, sTagAttribute$, uXmlStr$

uXmlStr = UCase(XmlStr)
SollTag = UCase(SollTag)

ret = ""
sNameSpace = ""
sSearch = "<" + SollTag
ind = 1
Do
    ind = InStr(ind, uXmlStr, sSearch)
    If (ind <= 0) Then
        ind = InStr(uXmlStr, ":" + SollTag)
        If (ind > 0) Then
            ind2 = ind
            Do
                ch = Mid(uXmlStr, ind2, 1)
                If (ch = "<") Then
                    Exit Do
                Else
                    sNameSpace = ch + sNameSpace
                    ind2 = ind2 - 1
                End If
            Loop
        End If
        sSearch = "<" + sNameSpace + SollTag '+ ">"
        ind = InStr(uXmlStr, sSearch)
    End If
    If (ind > 0) Then
        ind = ind + Len(sSearch)
        If (InStr(" >", Mid(XmlStr, ind, 1)) > 0) Then
            ind3 = ind
            ind = InStr(ind, XmlStr, ">")
            sTagAttribute = Trim(Mid(XmlStr, ind3, ind - ind3))
            If (Mid(XmlStr, ind - 1, 1) = "/") Then
                ind2 = ind '- 1
            Else
                ind = ind + 1
                sSearch = "</" + sNameSpace + SollTag + ">"
                ind2 = InStr(ind, uXmlStr, sSearch)  'bis 1.0.4: ind+1
                If (ind2 > 0) Then
                    ret = Trim(Mid(XmlStr, ind, ind2 - ind))
                    ind2 = ind2 + Len(sSearch)
                    XmlStr = Mid(XmlStr, ind2)
                End If
            End If
            Exit Do
        End If
    Else
        Exit Do
    End If
Loop

XmlAbschnitt = Trim(ret)

Call DefErrPop
End Function

'Function GetXmlTag$(XmlStr$, SollTag$, sValue$)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("XmlAbschnitt$")
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
'Dim ind&, ind2&, ind3&
'Dim sSearch$, sUcaseSearchTag$, ch$, sNameSpace$, sTagAttribute$
'
'sValue = ""
'sNameSpace = ""
'sSearch = "<" + SollTag
'ind = InStr(XmlStr, sSearch)
'If (ind <= 0) Then
'    ind = InStr(XmlStr, ":" + SollTag)
'    If (ind > 0) Then
'        ind2 = ind
'        Do
'            ch = Mid(XmlStr, ind2, 1)
'            If (ch = "<") Then
'                Exit Do
'            Else
'                sNameSpace = ch + sNameSpace
'                ind2 = ind2 - 1
'            End If
'        Loop
'    End If
'    sSearch = "<" + sNameSpace + SollTag '+ ">"
'    ind = InStr(XmlStr, sSearch)
'End If
'If (ind > 0) Then
'    ind = ind + Len(sSearch)
'    ind3 = ind
'    ind = InStr(ind, XmlStr, ">")
'    sTagAttribute = Trim(Mid(XmlStr, ind3, ind - ind3))
'    If (Mid(XmlStr, ind - 1, 1) = "/") Then
'        ind2 = ind '- 1
'    Else
'        ind = ind + 1
'        sSearch = "</" + sNameSpace + SollTag + ">"
'        ind2 = InStr(ind, XmlStr, sSearch)  'bis 1.0.4: ind+1
'        If (ind2 > 0) Then
'            sValue = Trim(Mid(XmlStr, ind, ind2 - ind))
'            ind2 = ind2 + Len(sSearch)
'            XmlStr = Mid(XmlStr, ind2)
'        End If
'    End If
'End If
'
'sValue = Trim(ret)
'
'Call DefErrPop
'End Function


Function iFileOpen%(Fname$, fAttr$, Optional modus$ = "B", Optional SATZLEN% = 100)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("iFileOpen%")
'Call DefErrMod(DefErrModul)
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
iFileOpen% = False
Handle% = FreeFile

fAttr$ = UCase(fAttr$)

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
ElseIf (fAttr$ = "RWL") Then
    If (modus$ = "B") Then
        Open Fname$ For Binary Access Read Write Lock Read Write As #Handle%
    Else
        Open Fname$ For Random Access Read Write Lock Read Write As #Handle% Len = SATZLEN%
    End If
ElseIf (fAttr$ = "I") Then
    Open Fname$ For Input Access Read Shared As #Handle%
ElseIf (fAttr$ = "O") Then
    Open Fname$ For Output Access Write Shared As #Handle%
ElseIf (fAttr$ = "A") Then
    Open Fname$ For Append Access Write Shared As #Handle%
End If

If (Err = 0) Then
    iFileOpen% = Handle%
Else
    Call DefErrAnswer(Err.Source, Err.Number, Err.Description + " (" + Fname$ + ")", DefErrModul, 1)
    End
End If

Call DefErrPop
End Function

Private Sub iLock(file As Integer, SatzNr&)
Dim LockTime As Date
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("iLock")
'Call DefErrMod(DefErrModul)
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

Lock #file, SatzNr&

Call DefErrPop
End Sub

Private Sub iUnLock(file As Integer, SatzNr&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("iUnLock")
'Call DefErrMod(DefErrModul)
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

Unlock #file, SatzNr&

Call DefErrPop
End Sub

Sub Logbuch(s$, Optional LeerZeileDavor = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Logbuch")
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
Dim LOGHANDLE%, ErrNumber%
Dim sTime$, s2$

s2 = s
Dim ind&, ind2&
ind = InStr(s2, "<eRezeptData>")
If (ind > 0) Then
    ind2 = InStr(s2, "</eRezeptData>")
    If (ind2 > 0) Then
        s2 = Left(s, ind + 12) + CStr(ind2 - ind) + Mid(s, ind2)
    End If
End If

'If (DatenProtokoll) Then
    On Error Resume Next
    Err.Clear
    LOGHANDLE% = FreeFile
    Open "\user\FiveRxE.LOG" For Append As #LOGHANDLE%
    ErrNumber% = Err.Number
    On Error GoTo DefErr
    If (ErrNumber% > 0) Then
    Else
        If (LeerZeileDavor) Then
            Print #LOGHANDLE%, " "
        End If
        Print #LOGHANDLE%, Format(Now, "dd.mm.yyyy hh:nn:ss  ") + Format$((Timer * 100) Mod 100, "00") + "  " + s2$
        Close #LOGHANDLE%
    End If
'End If

Call DefErrPop
End Sub

Sub TI_Back_Protokoll(s$, Optional LeerZeileDavor = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TI_Back_Protokoll")
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
Dim LOGHANDLE%, ErrNumber%
Dim sTime$

'If (DatenProtokoll) Then
    On Error Resume Next
    Err.Clear
    LOGHANDLE% = FreeFile
    Open "\user\TI_Back.log" For Append As #LOGHANDLE%
    ErrNumber% = Err.Number
    On Error GoTo DefErr
    If (ErrNumber% > 0) Then
    Else
        If (LeerZeileDavor) Then
            Print #LOGHANDLE%, " "
        End If
        Print #LOGHANDLE%, Format(Now, "dd.mm.yyyy hh:nn:ss  ") + Format$((Timer * 100) Mod 100, "00") + "  " + s$
        Close #LOGHANDLE%
    End If
'End If

Call DefErrPop
End Sub

