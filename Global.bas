Attribute VB_Name = "modGlobal"
Option Explicit

Private Type VerkaufStruct
  RezNr       As Integer
  EndeT       As Byte
  pzn         As String * 7
  text        As String * 36
  Preis       As String * 9
  mw          As String * 1
  bs          As Byte
  datum       As String * 2
  user        As Integer
  LaufNr      As Integer
  gebuehren   As String * 1
  wegmark     As String * 1
  AVP         As String * 1
  zeit        As String * 2 'Integer
  bon         As String * 1
  wg          As String * 2
  knr         As Integer
  gebsumme    As String * 8
  pflag       As Byte
  PersCode    As Byte
  angefordert As String * 1
  FremdGeld   As Byte
  FremdBetrag As String * 8
  TxtTyp      As String * 1
  RezEan      As String * 13
  KKNr        As String * 9
  KKTyp       As Byte
  ZuzaFlag    As String * 1
  rest        As String * 9
  multi       As Byte
End Type
Public VkRec As VerkaufStruct

Private Type KZustellType
    knr As String * 5
    Name As String * 30
    Strasse As String * 30
    PlzOrt As String * 30
    telefon As String * 30
    zeit As String * 5              'nicht mehr verwendet auﬂer zum umkopieren in zeitraum
    Geschenke As Integer
    Zahlungsart As String * 1
    KarteNr As String * 19
    gueltig As String * 4
    zHdn As String * 30
    Zeitraum As String * 19
    rest As String * 51
End Type
Public kzu As KZustellType

'Public clsfabs As New clsFabsKlasse
Public clsDat As New clsDateien
Public clsOpTool As New clsOpTools
Public clsDialog As New clsDialoge
Public clsError As New clsError
Public clsKz As New clsKennzeichen
Public clsSI As New clsSendInput

Public Ast1 As clsStamm
Public Ass1 As clsStatistik
Public ArtikelDB1 As clsArtikelDB
Public Taxe1 As clsTaxe
Public TaxeAdoDB1 As clsTaxeAdoDB
Public Lif1 As clsLieferanten
Public LifZus1 As clsLiefZusatz
Public LieferantenDB1 As clsLieferantenDB
Public Ww1 As clsWawiDat
Public WawiMdb1 As clsWawiMdb
Public WawiDB1 As clsWawiDB
Public Bek1 As clsBekart
Public Nb1 As clsNachbearbeitung
Public Wu1 As clsW‹
Public Besorgt1 As clsBesorgt
Public Merkzett1 As clsMerkzettel
Public Absagen1 As clsAbsagen
Public AbsagenDB1 As clsAbsagenDB
Public Kiste1 As clsKiste
Public Zus1 As clsArttext
Public lZus1 As clsLieftext
Public Para1 As clsOpPara
Public wPara1 As clsWinPara
Public Etik1 As clsEtiketten
Public EtikettenDB1 As clsEtikettenDB
Public clsAngebote1 As clsAngebote
Public AngeboteDB1 As clsAngeboteDB
Public clsManuelleAngebote1  As clsManuellAngebote
Public ManuellAngeboteDB1 As clsManuellAngeboteDB
Public Pass1 As clsPasswort
Public ArtStat1 As clsArtStatistik
Public Lb1 As clsLagerBewegung
Public LbDB1 As clsLagerbewegungDB
Public RabTab1 As clsRabattTabellen
Public nnek1 As clsNNEK
Public DM1 As clsDocMorris
Public mPos1 As clsMpos
Public Vk1 As clsVerkauf
Public VerkaufDB1 As clsVerkaufDB
Public RezTab1 As clsVerkRtab
Public VmPzn1 As clsVmPzn
Public VmBed1 As clsVmBed
Public VmRech1 As clsVmRech
Public hTax1 As clsHilfsTaxe
Public HilfstaxeDB1 As clsHilfstaxeDB
Public Kkasse1 As clsKkassen
Public kc1 As clsKredcard
Public Kun1 As clsKunden
Public KunZus1 As clsKundZusatz
Public bez1 As clsKunBezug
Public BezugDB1 As clsBezugDB
'Public kZustell1 As clsKundZustell

Public sqlop1 As clsSqlTools

Public TaxeDB As Database
Public TaxeRec As Recordset

Public TaxeAdoRec As New ADODB.Recordset
Public TaxeAdoDBok%

Public WawiDB As Database
Public WawiRec As Recordset

Public WawiAdoRec As New ADODB.Recordset
Public WawiDBok%

Public MerkzettelDB As Database
Public MerkzettelRec As Recordset

'Public LieferantenDB As Database
'Public LieferantenRec As Recordset
'Public LieferantenConn As New ADODB.Connection
'Public LieferantenComm As New ADODB.Command
Public LieferantenRec As New ADODB.Recordset
Public LieferantenzusatzRec As New ADODB.Recordset
Public AusnahmenRec As New ADODB.Recordset
Public RabattTabelleRec As New ADODB.Recordset
Public BmEkTabelleRec As New ADODB.Recordset
Public Msv3Rec As New ADODB.Recordset
Public LieferantenDBok%

Public HerstellerRec As New ADODB.Recordset

'Public ArtikelConn As New ADODB.Connection
'Public ArtikelComm As New ADODB.Command
Public ArtikelRec As New ADODB.Recordset
Public ArtikelRec2 As New ADODB.Recordset
Public DocMorrisRec As New ADODB.Recordset
Public SonderPrAdoRec As New ADODB.Recordset
'Public LieferantenzusatzRec As New ADODB.Recordset
'Public AusnahmenRec As New ADODB.Recordset
'Public RabattTabelleRec As New ADODB.Recordset
'Public BmEkTabelleRec As New ADODB.Recordset
Public ArtikelDBok%

Public HilfstaxeRec As New ADODB.Recordset
Public HilfstaxeDBok%

'Public AngeboteConn As New ADODB.Connection
'Public AngeboteComm As New ADODB.Command
Public AngeboteRec As New ADODB.Recordset
Public AngeboteDbOk%

'Public ManuelleAngeboteConn As New ADODB.Connection
'Public ManuelleAngeboteComm As New ADODB.Command
Public ManuelleAngeboteRec As New ADODB.Recordset
Public ManuelleAngeboteDbOk%

Public VerkaufDbOk%

Public AnzeigeFensterTyp$

'Public BesorgerAconto$
'Public BesorgerMenge%
Public BesorgerPzn$
Public BesorgerNummer%
Public BesorgerModus%
Public BesorgerBenutzer%

Public WechselZeilen$()
Public WechselStart%
Public WechselErg%

Public ZusatzFensterTyp$
Public ZusatzPzn$
Public ZusatzAnzTxt%
Public ZusatzChanged%

Public IndexBelegt%(20)

Public FabsErrf%


Public InfoLayoutDatei$
Public InfoLayoutBelegung$
Public InfoLayoutBezeichnung$

Public ArtNeu%, AstNeu%, AssNeu%, TaxeNeu%, LiefNeu%, NnekNeu%, hTaxNeu%, DocMorrisNeu%
Public ArtDa%, AstDa%, AssDa%, TaxeDa%, LiefDa%, NnekDa%, hTaxDa%, DocMorrisDa%

Public iMehrPlatz%
Public iUser$

Public DefErrFncStr(50) As String
Public DefErrModStr(50) As String
Public DefErrStk As Integer

Public iDefErrAutoErrors$

Public ProjektForm As Form

Public EkEingabeText$
Public EkEingabePreis#
Public EkEingabeModus%
Public EkEingabeArt%
Public EkEingabeErg%

Public Dp04Ok%
Public SeriellScannerOk%
Public SeriellScannerParam$

Public AktWumsatzLief%
Public AktWumsatzInfo$
Public AktWumsatzTyp$

Public StammdatenPzn$
Public StammdatenStart$
Public StammdatenModus%
Public StammdatenWert As Variant
Public StammdatenClass As Object

Public FremdPznTest%
Public FremdPznOk%
Public FremdPznDB As Database
Public FremdPznRec As Recordset

Public OpPartnerDB As Database
Public OpPartnerRec As Recordset

Public OpPartnerPzn$
Public OpPartnerTxt$
Public OpPartnerBestell%
Public OpPartnerInitLief%

Public SignaturEingabeErg%, SignaturEingabeModus%


Public AbholerDB As Database
Public AbholerNummerRec As Recordset
Public AbholerDetailRec As Recordset
Public AbholerInfoRec  As Recordset
Public AbholerParaRec As Recordset
Public AbholerMdb%

Public AbholerSQL%
Public AbholerConn As New ADODB.Connection
Public AbholerNummerAdoRec As New ADODB.Recordset
Public AbholerDetailAdoRec As New ADODB.Recordset
Public AbholerInfoAdoRec  As New ADODB.Recordset
Public AnfMagAdoRec  As New ADODB.Recordset
Public AbholerParaAdoRec As New ADODB.Recordset



Public iNewLine%
Public iMARS%
Public iFaktorX!, iFaktorY!
Public PixelX&, PixelY&
    
Public nlmsgDefault%(2)
Public nlmsgCancel%(2)
Public nlmsgVisible%(2)
Public nlmsgCaption$(2)

Public nlmsgPicto%
Public nlmsgTitle$, nlmsgPrompt$
Public nlmsgRet%
Public nlmsgRet2$
Public nlmsgEditModus%

Public nlmsgAktiv%

Public ManuelleAngeboteAktiv%

Public SQLStr$

''''''''''''''''''''''''''''
'Public SqlServer$
'Public ConnectionString$, SqlConnectionString$(1)
'Public iSqlServerDataPath$, iNetworkDataPath$
'Public SqlServerDB$(20) ', Database$(1), SqlServerAktiv%(1)
'Public SqlSuffix$(10)
'Public TaxeDatum$
'
'Public SqlError&
'Public SqlErrorDesc$
'
'Public SqlConnectErg%
'
'Public GlobalConn As New ADODB.Connection
'Public GlobalComm As New ADODB.Command
''''''''''''''''''''''''''''
Public AbsagenDBok%
Public LagerbewDBok%
Public EtikettenDBok%

Public ToolbarModus%
Public ToolbarBackR&, ToolbarBackG&, ToolbarBackB&

'Public KundenDB1 As clsKundenDB
Public KundenAdoRec As New ADODB.Recordset
Public KundenDBok%


Private Const DefErrModul = "GLOBAL.BAS"

Sub Main()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Main")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, FabsRecno&
Dim l&
Dim h$, sClientName$, sOpUserIni$

'ChDrive "F"

PixelX = Screen.Width / Screen.TwipsPerPixelX
PixelY = Screen.Height / Screen.TwipsPerPixelY
iFaktorX = PixelX / 1680
iFaktorY = PixelY / 1050

iMehrPlatz% = True
iUser$ = Environ$("USER")

If (Dir$("fistam.dat") = "") Then ChDir "\user"

sOpUserIni = CurDir + "\OpUser.ini"
If (Dir(sOpUserIni) <> "") Then
    sClientName = Trim(Environ$("CLIENTNAME"))
    If (sClientName <> "") Then
        h$ = "01"
        l& = GetPrivateProfileString("ComputerNr", sClientName, "01", h$, 3, sOpUserIni)
        iUser = CStr(Val(h))
    '    MsgBox (iUser)
    End If
End If

'h$ = "0"
'l& = GetPrivateProfileString("Allgemein", "Newline", "0", h$, 2, "newline.ini")
'h$ = Left$(h$, l&)
'iNewLine = (Val(h$) <> 0)
iNewLine = (Dir(CurDir + "\newline.ini") <> "")
nlmsgAktiv = 0
    
h$ = "J"
l& = GetPrivateProfileString("Bestellung", "ManuelleAngebote", "J", h$, 2, CurDir + "\winop.ini")
h$ = Left$(h$, l&)
If (h$ = "J") Then
    ManuelleAngeboteAktiv% = True
Else
    ManuelleAngeboteAktiv% = False
End If
    
h$ = "0"
l& = GetPrivateProfileString("Mars", "Mars", "0", h$, 2, CurDir + "\vkpara.ini")
h$ = Left$(h$, l&)
iMARS = (Val(h$) = 1)


'FabsErrf% = clsfabs.Cmnd("Y\2", FabsRecno&)

For i% = 20 To 0 Step -1
    IndexBelegt%(i%) = False
Next i%

'KundenMdbOk% = 0
KundenDBok% = 0
ArtikelDBok = 0
HilfstaxeDBok = 0

LieferantenDBok = 0
AbsagenDBok = 0
LagerbewDBok = 0
EtikettenDBok = 0
AngeboteDbOk = 0
ManuelleAngeboteDbOk = 0
TaxeAdoDBok = 0
WawiDBok = 0

'If (UCase(Left(CurDir, 1)) = "Z") Then
'    ToolbarModus = 0
'Else
'    ToolbarModus = 2
'End If

Set DM1 = New clsDocMorris

Call clsError.DefErrPop
End Sub

