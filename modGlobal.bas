Attribute VB_Name = "modGlobal"
Option Explicit

Private Type VerkaufStruct
  RezNr       As Integer
  EndeT       As Byte
  pzn         As String * 7
  text        As String * 36
  preis       As String * 9
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
  pFlag       As Byte
  PersCode    As Byte
  Angefordert As String * 1
  FremdGeld   As Byte
  FremdBetrag As String * 8
  TxtTyp      As String * 1
  RezEan      As String * 13
  KKNr        As String * 9
  KKTyp       As Byte
  ZuzaFlag    As String * 1
  rest        As String * 9
  Multi       As Byte
End Type
Public VkRec As VerkaufStruct

Public clsfabs As New clsFabsKlasse
Public clsDat As New clsDateien
Public clsOpTool As New clsOpTools
Public clsDialog As New clsDialoge
Public clsError As New clsError

Public Ast1 As clsStamm
Public Ass1 As clsStatistik
Public Taxe1 As clsTaxe
Public Lif1 As clsLieferanten
Public Bek1 As clsBekart
Public Nb1 As clsNachbearbeitung
Public Wu1 As clsWÜ
Public Besorgt1 As clsBesorgt
Public Absagen1 As clsAbsagen
Public Kiste1 As clsKiste
Public Zus1 As clsArttext
Public lZus1 As clsLieftext
Public Para1 As clsOpPara
Public wPara1 As clsWinPara
Public Etik1 As clsEtiketten
Public clsAngebote1 As clsAngebote
Public Pass1 As clsPasswort

Public TaxeDB As Database
Public TaxeRec As Recordset

Public AnzeigeFensterTyp$

Public BesorgerAconto$
Public BesorgerMenge%
Public BesorgerPzn$
Public BesorgerNummer%

Public WechselZeilen$()
Public WechselStart%
Public WechselErg%

Public ZusatzFensterTyp$
Public ZusatzPzn$

Public IndexBelegt%(20)

Public FabsErrf%


Public InfoLayoutDatei$
Public InfoLayoutBelegung$

Public iMehrPlatz%
Public iUser$


Private Const DefErrModul = "modGlobal.bas"

Sub Main()
Dim i%, FabsRecno&

iMehrPlatz% = True
iUser$ = Environ$("USER")

If (Dir$("fistam.dat") = "") Then ChDir "\user"

FabsErrf% = clsfabs.Cmnd("Y\3", FabsRecno&)

For i% = 20 To 0 Step -1
    IndexBelegt%(i%) = False
Next i%

End Sub

