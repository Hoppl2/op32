VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlbutton.ocx"
Begin VB.Form frmRezSpeicher 
   AutoRedraw      =   -1  'True
   Caption         =   "Importkontrolle"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   8595
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   5520
      Picture         =   "RezSpeicher.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   5280
      Picture         =   "RezSpeicher.frx":00B9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   5040
      Picture         =   "RezSpeicher.frx":016D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin nlCommandButton.nlCommand nlcmdKKNeu 
      Height          =   495
      Left            =   4680
      TabIndex        =   12
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin VB.CommandButton cmdImpAlt 
      Caption         =   "Imp.-Alternativen (F9)"
      Height          =   450
      Index           =   1
      Left            =   3000
      TabIndex        =   9
      Top             =   4440
      Width           =   1200
   End
   Begin VB.CommandButton cmdRezSpeicherLoeschen 
      Caption         =   "Löschen (F5)"
      Height          =   450
      Left            =   3120
      TabIndex        =   5
      Top             =   1920
      Width           =   1200
   End
   Begin VB.CommandButton cmdDruck 
      Caption         =   "Druck (F6)"
      Height          =   450
      Left            =   3120
      TabIndex        =   6
      Top             =   2520
      Width           =   1200
   End
   Begin VB.CommandButton cmdImpAlt 
      Caption         =   "Abg. Originale (F8)"
      Height          =   450
      Index           =   0
      Left            =   3120
      TabIndex        =   8
      Top             =   3720
      Width           =   1200
   End
   Begin VB.PictureBox picFont 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   795
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdAbrech 
      Caption         =   "Abrechnung (F7)"
      Height          =   450
      Left            =   3240
      TabIndex        =   7
      Top             =   3120
      Width           =   1200
   End
   Begin VB.CommandButton cmdKStamm 
      Caption         =   "Stammdaten (F3)"
      Height          =   450
      Left            =   3120
      TabIndex        =   4
      Top             =   1320
      Width           =   1200
   End
   Begin VB.CommandButton cmdKKNeu 
      Caption         =   "neue Kasse (F2)"
      Height          =   450
      Left            =   3120
      TabIndex        =   3
      Top             =   720
      Width           =   1200
   End
   Begin VB.ComboBox cmbDatum 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown-Liste
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   3000
      TabIndex        =   10
      Top             =   5160
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxRezSpeicher 
      Height          =   2700
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   4763
      _Version        =   393216
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483633
      BackColorBkg    =   -2147483633
      FocusRect       =   0
      HighLight       =   2
      GridLines       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin nlCommandButton.nlCommand nlcmdKStamm 
      Height          =   495
      Left            =   4680
      TabIndex        =   13
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdRezSpeicherLoeschen 
      Height          =   495
      Left            =   4680
      TabIndex        =   14
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdDruck 
      Height          =   495
      Left            =   4680
      TabIndex        =   15
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdAbrech 
      Height          =   495
      Left            =   4680
      TabIndex        =   16
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdImpAlt 
      Height          =   495
      Index           =   0
      Left            =   4680
      TabIndex        =   17
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdImpAlt 
      Height          =   495
      Index           =   1
      Left            =   4680
      TabIndex        =   18
      Top             =   4320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   495
      Left            =   4680
      TabIndex        =   19
      Top             =   5040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin VB.Label lblDatum 
      AutoSize        =   -1  'True
      Caption         =   "&Datum"
      Height          =   195
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmRezSpeicher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "REZSPEICHER.FRM"

Private FieldNr%(8)

Dim AuswertungsDatum$

Const W_GESAMTJAHR = 0
Const W_GESAMT = 1
Const W_FAM = 2
Const W_IMPFÄHIG = 3
Const W_IMPIST = 4
Const W_ABZUG = 5
Const W_SALDO = 6
Const W_REZANZ = 7
Const W_IMPERSPART = 8

Dim iEditModus%


Private Sub cmdDruck_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdDruck_Click")
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

If (RezSpeicherModus = 0) Then
    Call AusdruckRezeptSpeicher
Else
    Call AusdruckImportkontrolle
End If

Call DefErrPop
End Sub

Sub cmdImpAlt_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdImpAlt_Click")
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

If (RezSpeicherModus = 0) Then
    Call TesteGH
ElseIf DarfRezSpeicher And (cmdImpAlt(index).Enabled) Then
    EditErg% = True
    If (index = 0) Then
        frmAbgOriginale.Show 1
    End If
    If (EditErg%) Then
        ImpAlternativModus% = index
        frmImpAlternativ.Show vbModal
    End If
End If

Call DefErrPop
End Sub

Private Sub cmbDatum_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmbDatum_Click")
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
Dim AuswertungsMonat$

AuswertungsMonat$ = cmbDatum.List(cmbDatum.ListIndex)
If InStr(AuswertungsMonat$, "gesamt") = 0 Then
  AuswertungsMonat$ = Format(CDate("01. " + AuswertungsMonat$), "YYMM")
Else
  AuswertungsMonat$ = Mid(AuswertungsMonat$, 3, 2)
End If
Call flxRezSpeicherBefuellen(AuswertungsMonat$)

Call DefErrPop
End Sub

Private Sub cmdAbrech_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdAbrech_Click")
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
Dim OldRow&, OldKsatz&
Dim i%

If (RezSpeicherModus% = 0) Then
    frmIkAuswertung.Show 1
ElseIf DarfRezSpeicher And (cmdAbrech.Enabled) Then
    With flxRezSpeicher
        AbrechMo = cmbDatum.List(cmbDatum.ListIndex)
        If InStr(AbrechMo, "gesamt") = 0 And Trim(AbrechMo) > "" Then
          AbrechMo = Format(CDate("01. " + AbrechMo), "YYMM")
          'If Val(.TextMatrix(.row, 1)) > 0 Then
              OldRow& = .row
              KKSatz = .TextMatrix(.row, 3)
              On Error Resume Next    'durch Unload im Form_load kann Fehler auftreten
              frmKKAbrech.Show vbModal
              On Error GoTo DefErr
              Call cmbDatum_Click     'Neuanzeige
              If .Rows >= OldRow& - 1 Then
                  .row = OldRow&
                  If .TextMatrix(OldRow&, 3) <> KKSatz Then
                      For i% = 5 To .Rows - 1
                          If .TextMatrix(i%, 3) = KKSatz Then
                              .row = i%
                              Exit For
                          End If
                      Next i%
                  End If
              End If
          'End If
          flxRezSpeicher.SetFocus
        Else
          cmbDatum.SetFocus
        End If
    End With
End If

Call DefErrPop
End Sub

Private Sub cmdEsc_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdEsc_Click")
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

Unload Me

Call DefErrPop
End Sub

Private Sub cmdKKNeu_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdKKNeu_Click")
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
    
KKSatz = ""
frmKKassen.Show vbModal
Call cmbDatum_Click     'Neuanzeige

Call DefErrPop
End Sub

Private Sub cmdKStamm_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdKStamm_Click")
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

Dim OldRow&
With flxRezSpeicher
    If Val(.TextMatrix(.row, 1)) <> 0 Then
        OldRow& = .row
        KKSatz = .TextMatrix(.row, 3)
        frmKKassen.Show vbModal
        Call cmbDatum_Click     'Neuanzeige
        If .Rows >= OldRow& - 1 Then .row = OldRow&
    Else
        If (para.Newline) Then
            Call nlcmdkkneu_Click
        Else
            Call cmdKKNeu_Click
        End If
    End If
    flxRezSpeicher.SetFocus
End With

Call DefErrPop
End Sub

Private Sub cmdRezSpeicherLoeschen_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdRezSpeicherLoeschen_Click")
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
Dim h$
Dim OldRow&
Dim erg%, iBenutzerNr%
Dim KKDelete As Boolean

iBenutzerNr% = HoleBenutzerSignatur
If (iBenutzerNr = 1) Then
Else
    Call MessageBox("Problem: Löschen des Rezeptspeichers nur mit Chef-Passwort möglich!", vbCritical)
    Call DefErrPop: Exit Sub
End If

KKDelete = True
If (RezSpeicherModus% = 0) Then
    erg% = iMsgBox("Rezeptspeicher löschen?" + vbCr + "(Nein = Krankenkasse löschen)", vbYesNo Or vbInformation, "Rezeptspeicher ")
    If erg% = vbNo Then
        KKDelete = True
    Else
        Do
'            h$ = MyInputBox("Rezepte löschen bis inkl.: ", "Rezeptspeicher löschen", Format(Now, "DDMMYY"))
            h$ = MyInputBox("Rezepte löschen bis inkl.: ", "Rezeptspeicher löschen", "3112" + Format(Year(Now) - 2000 - 2, "00"))
            h$ = UCase(Trim(h$))
            If (h$ = "") Then
                Exit Do
            ElseIf (iDate(h$) <> 0) Then
                Exit Do
            End If
        Loop
        If (h$ <> "") Then Call RezSpeicherLoeschen(h$)
    End If
End If

If KKDelete Then
    With flxRezSpeicher
        If Val(.TextMatrix(.row, 1)) <> 0 Then
            OldRow& = .row
            KKSatz = .TextMatrix(.row, 3)
            KasseRec.index = "Unique"
            KasseRec.Seek "=", KKSatz$
            If Not KasseRec.NoMatch Then
                erg% = iMsgBox(Trim(KasseRec!Name) + " wirklich löschen?", vbYesNo Or vbDefaultButton2 Or vbInformation, "Kassen-Stammdaten")
                If erg% = vbYes Then Call KKLoeschen
            End If
            'frmKKassen.Show vbModal
            Call cmbDatum_Click     'Neuanzeige
            If .Rows >= OldRow& - 1 Then .row = OldRow&
        End If
        .SetFocus
    End With
End If

Call DefErrPop
End Sub

Private Sub RezSpeicherLoeschen(AbLoesch$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RezSpeicherLoeschen")
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
Dim j%, AnzRezArtikel%
Dim h$, Unique$, AbLoeschSort$

MousePointer = vbHourglass

AbLoeschSort$ = Mid$(AbLoesch$, 5, 2) + Mid$(AbLoesch$, 3, 2) + Left$(AbLoesch$, 2)

Set RezepteRec = RezSpeicherDB.OpenRecordset("Rezepte", dbOpenTable)
Set ArtikelRec = RezSpeicherDB.OpenRecordset("Artikel", dbOpenTable)

RezepteRec.index = "KasseDruck"
ArtikelRec.index = "Unique"

RezepteRec.MoveFirst
Do While Not RezepteRec.EOF
    If (RezepteRec!DruckDatum <= AbLoeschSort$) Then
        Unique$ = RezepteRec!Unique
        AnzRezArtikel% = RezepteRec!AnzArtikel

        For j% = 0 To (AnzRezArtikel% - 1)
            h$ = Unique$ + Format(j% + 1, "0")
            If j% >= 9 And j% <= 243 Then
                h$ = Unique$ + Chr$(j% + 12)    'für Privatrezepte
            ElseIf j% > 243 Then Exit For      'kann's nicht geben, nur zur Sicherheit
            End If
            ArtikelRec.Seek "=", h$
            If Not ArtikelRec.NoMatch Then
                ArtikelRec.Delete
            End If
        Next j%
        
        RezepteRec.Delete
    Else
        Beep
    End If
    RezepteRec.MoveNext
Loop

RezSpeicherDB.Close

' Sicherstellen, daß nicht schon eine Datei mit dem Namen der komprimierten Datenbank
' existiert.
h$ = "rezepte2.mdb"
If Dir(h$) <> "" Then
    Kill h$
End If
Name REZ_SPEICHER As h$

DBEngine.CompactDatabase h$, REZ_SPEICHER
    
MousePointer = vbDefault

Set RezSpeicherDB = OpenDatabase(REZ_SPEICHER, False, False)
Call ProgrammEnde

Call DefErrPop
End Sub

Private Sub flxRezSpeicher_DblClick()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxRezSpeicher_DblClick")
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

Call flxRezSpeicher_KeyPress(13)

Call DefErrPop

End Sub

Private Sub flxRezSpeicher_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxRezSpeicher_GotFocus")
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

With flxRezSpeicher
    .col = 0
    .ColSel = .Cols - 1
    .HighLight = flexHighlightAlways
End With

Call DefErrPop
End Sub

Private Sub flxRezSpeicher_KeyPress(KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxRezSpeicher_KeyPress")
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
Dim iPos&, OldRow&
Dim iAnzeige As Boolean
Dim h$, Sort$

If (KeyAscii = vbKeySpace) Then
    With flxRezSpeicher
        Sort$ = .TextMatrix(.row, 0)
        If Left(Sort$, 1) = "4" Or Left(Sort$, 1) = "5" Then
          iPos& = Val(.TextMatrix(.row, 1))
          h$ = Trim(.TextMatrix(.row, 2))
          iAnzeige = (h$ = "")
          
          KasseRec.index = "Unique"
          KasseRec.Seek "=", Trim(.TextMatrix(.row, 3))
          If Not KasseRec.NoMatch Then
              KasseRec.Edit
              KasseRec!Anzeige = iAnzeige
              KasseRec.Update
          End If
          
'          Kkasse.GetRecord (iPos& + 1)
'          Kkasse.Anzeige = Abs(iAnzeige)
'          Kkasse.PutRecord (iPos& + 1)
          
          If (iAnzeige) Then
              h$ = Chr$(214)
              Sort$ = "4" + Mid(Sort$, 2)
          Else
              h$ = ""
              Sort$ = "5" + Mid(Sort$, 2)
          End If
          .TextMatrix(.row, 2) = h$
          .TextMatrix(.row, 0) = Sort$
        End If
    End With
ElseIf KeyAscii = vbKeyReturn Then
    If (RezSpeicherModus% = 0) Then
        With flxRezSpeicher
'            If Val(.TextMatrix(.row, 1)) > 0 Then
                RezHistorieKassenNr$ = .TextMatrix(.row, 3)
                
                RezHistorieIndexSuche% = True
                If (RezHistorieKassenNr$ = "") Or (RezHistorieKassenNr$ = "AlleKK   ") Then
                    RezHistorieIndexSuche% = False
                End If
                RezHistorieKassenName$ = .TextMatrix(.row, 5)
                RezHistorieDatum$ = AuswertungsDatum$
                If (Len(RezHistorieDatum$) = 2) Then
                    RezHistorieTagDirekt% = False
                    frmRezMonate.Show 1
                Else
                    RezHistorieTagDirekt% = True
                    frmRezTage.Show 1
                End If
'                frmRezHistorie.Show 1
    '        Else
    '            Call cmdKKNeu_Click
'            End If
            flxRezSpeicher.SetFocus
        End With
    End If

End If

Call DefErrPop
End Sub

Private Sub flxRezSpeicher_LostFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxRezSpeicher_LostFocus")
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

With flxRezSpeicher
    .HighLight = flexHighlightNever
End With

Call DefErrPop
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_KeyDown")
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
Dim OldRow&

If (para.Newline) Then
    If KeyCode = vbKeyF2 Then
        nlcmdKKNeu.Value = 1
        KeyCode = 0
    ElseIf KeyCode = vbKeyF3 Then
        nlcmdKStamm.Value = 1
        KeyCode = 0
    ElseIf KeyCode = vbKeyF5 Then
        nlcmdRezSpeicherLoeschen.Value = 1
        KeyCode = 0
    ElseIf KeyCode = vbKeyF6 Then
        nlcmdDruck.Value = 1
        KeyCode = 0
    ElseIf KeyCode = vbKeyF7 Then
        nlcmdAbrech.Value = 1
        KeyCode = 0
    ElseIf KeyCode = vbKeyF8 Then
        nlcmdImpAlt(0).Value = 1
        KeyCode = 0
    ElseIf KeyCode = vbKeyF9 Then
        nlcmdImpAlt(1).Value = 1
        KeyCode = 0
    End If
Else
    If KeyCode = vbKeyF2 Then
        Call cmdKKNeu_Click
        KeyCode = 0
    ElseIf KeyCode = vbKeyF3 Then
        Call cmdKStamm_Click
        KeyCode = 0
    ElseIf KeyCode = vbKeyF5 Then
        Call cmdRezSpeicherLoeschen_Click
        KeyCode = 0
    ElseIf KeyCode = vbKeyF6 Then
        Call cmdDruck_Click
        KeyCode = 0
    ElseIf KeyCode = vbKeyF7 Then
        Call cmdAbrech_Click
        KeyCode = 0
    ElseIf KeyCode = vbKeyF8 Then
        Call cmdImpAlt_Click(0)
        KeyCode = 0
    ElseIf KeyCode = vbKeyF9 Then
        Call cmdImpAlt_Click(1)
        KeyCode = 0
    End If
End If

Call DefErrPop
End Sub

Sub ImpKontrolleDruckKopf()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ImpKontrolleDruckKopf")
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
Dim i%, x%, y%
Dim gesBreite&
Dim header$, KopfZeile$, Typ$, h$

'KopfZeile$ = "Import-Kontrolle"
'header$ = "Import-Kontrolle" + " " + cmbDatum.List(cmbDatum.ListIndex)
KopfZeile$ = Me.Caption
header$ = KopfZeile$ + " " + cmbDatum.List(cmbDatum.ListIndex)
Call DruckKopf(header$, Typ$, KopfZeile$, 0)
Printer.CurrentY = Printer.CurrentY - 3 * Printer.TextHeight("A")
    
For i% = 0 To (AnzDruckSpalten% - 1)
    h$ = RTrim(DruckSpalte(i%).Titel)
    If (DruckSpalte(i%).Ausrichtung = "L") Then
        x% = DruckSpalte(i%).StartX
    Else
        x% = DruckSpalte(i%).StartX + DruckSpalte(i%).BreiteX - Printer.TextWidth(h$)
    End If
    Printer.CurrentX = x%
    Printer.Print h$;
Next i%

Printer.Print " "

y% = Printer.CurrentY
gesBreite& = DruckSpalte(AnzDruckSpalten% - 1).StartX + DruckSpalte(AnzDruckSpalten% - 1).BreiteX
Printer.Line (DruckSpalte(0).StartX, y%)-(gesBreite&, y%)

y% = Printer.CurrentY
Printer.CurrentY = y% + 30

Call DefErrPop

End Sub

Sub AusdruckRezeptSpeicher()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AusdruckRezeptSpeicher")
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
Dim ind%, ZeilenHöhe%, i%, j%, OrgAusrichtung%
Dim h$

AnzDruckSpalten% = flxRezSpeicher.Cols - 6
ReDim DruckSpalte(AnzDruckSpalten% - 1)

DruckSpalte(0).TypStr = String$(30, "9")
DruckSpalte(1).TypStr = String$(11, "9")
DruckSpalte(2).TypStr = String$(11, "9")
DruckSpalte(3).TypStr = String$(11, "9")
DruckSpalte(4).TypStr = String$(11, "9")
DruckSpalte(5).TypStr = String$(11, "9")
DruckSpalte(6).TypStr = String$(11, "9")
DruckSpalte(7).TypStr = String$(11, "9")
DruckSpalte(8).TypStr = String$(11, "9")

For i = 0 To (AnzDruckSpalten - 1)
    With DruckSpalte(i)
        .Ausrichtung = "L"
        
        h$ = flxRezSpeicher.TextMatrix(0, i + 5)
        ind = InStr("<^>", Left(h, 1))
        If (ind > 0) Then
            h = Mid(h, 2)
        End If
        .Titel = h$
        
'        If (ind = 2) Then
'            .Ausrichtung = "Z"
'        ElseIf (ind = 3) Then
'            .Ausrichtung = "R"
'        End If
        
        If (flxRezSpeicher.ColAlignment(i + 5) = flexAlignCenterCenter) Then
            .Ausrichtung = "Z"
        ElseIf (flxRezSpeicher.ColAlignment(i + 5) = flexAlignRightCenter) Then
            .Ausrichtung = "R"
        End If
    End With
Next i

OrgAusrichtung% = Printer.Orientation
Printer.Orientation = vbPRORLandscape
Call InitDruckZeile(True)

DruckSeite% = 0
Call ImpKontrolleDruckKopf
ZeilenHöhe% = Printer.TextHeight("A")
'DruckSpalte(0).Attrib = 2
With flxRezSpeicher
    For i% = 1 To .Rows - 1
        h$ = ""
        For j% = 1 To .Cols - 1
            If .ColWidth(j%) > 0 Then
                h$ = h$ + .TextMatrix(i%, j%) + vbTab
            End If
        Next j%
        Call DruckZeile(h$)
        If (Printer.CurrentY > Printer.ScaleHeight - 1000 - ZeilenHöhe%) Then
            Call DruckFuss
            Call ImpKontrolleDruckKopf
        End If
    Next i%
End With
Call DruckFuss(False)
Printer.EndDoc
Printer.Orientation = OrgAusrichtung%

Call DefErrPop
End Sub

Sub AusdruckImportkontrolle()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AusdruckImportkontrolle")
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
Dim ZeilenHöhe%, i%, j%, OrgAusrichtung%
Dim h$

AnzDruckSpalten% = 15
ReDim DruckSpalte(AnzDruckSpalten% - 1)

With DruckSpalte(0)
    .Titel = " "
    .TypStr = String$(1, "X")
    .Ausrichtung = "L"
    .Attrib = 1
End With
With DruckSpalte(1)
    .Titel = " "
    .TypStr = String$(1, "X")
    .Ausrichtung = "L"
End With
With DruckSpalte(2)
    .Titel = "K A S S E"
    .TypStr = String$(24, "X")  '28
    .Ausrichtung = "L"
End With
With DruckSpalte(3)
    .Titel = "ges.Jahr"
    .TypStr = String$(11, "9")
    .Ausrichtung = "R"
End With
With DruckSpalte(4)
    .Titel = "ges.Monat"
    .TypStr = String$(11, "9")
    .Ausrichtung = "R"
End With
With DruckSpalte(5)
    .Titel = "Anzahl"
    .TypStr = String$(7, "9")
    .Ausrichtung = "R"
End With
With DruckSpalte(6)
    .Titel = "FAM"
    .TypStr = String$(10, "9")
    .Ausrichtung = "R"
End With
With DruckSpalte(7)
    .Titel = "ImpFähig"
    .TypStr = String$(10, "9")
    .Ausrichtung = "R"
End With
With DruckSpalte(8)
    .Titel = "ImpIst"
    .TypStr = String$(10, "9")
    .Ausrichtung = "R"
End With
With DruckSpalte(9)
    .Titel = "Qu%"
    .TypStr = "99.99"
    .Ausrichtung = "R"
End With
With DruckSpalte(10)
    .Titel = "Q.EUR"
    .TypStr = String$(9, "9")
    .Ausrichtung = "R"
End With
With DruckSpalte(11)
    .Titel = "Diff."
    .TypStr = String$(8, "9")
    .Ausrichtung = "R"
End With
With DruckSpalte(12)
    .Titel = "Erspart"
    .TypStr = String$(8, "9")
    .Ausrichtung = "R"
End With
With DruckSpalte(13)
    .Titel = "+/-"
    .TypStr = String$(7, "9")
    .Ausrichtung = "R"
End With
With DruckSpalte(14)
    .Titel = "Saldo"
    .TypStr = String$(7, "9")
    .Ausrichtung = "R"
End With
OrgAusrichtung% = Printer.Orientation
Printer.Orientation = vbPRORLandscape
Call InitDruckZeile(True)

DruckSeite% = 0
Call ImpKontrolleDruckKopf
ZeilenHöhe% = Printer.TextHeight("A")
DruckSpalte(0).Attrib = 2
With flxRezSpeicher
    For i% = 1 To .Rows - 1
        h$ = ""
        For j% = 1 To .Cols - 1
            If .ColWidth(j%) > 0 Then
                h$ = h$ + .TextMatrix(i%, j%) + vbTab
            End If
        Next j%
        Call DruckZeile(h$)
        If (Printer.CurrentY > Printer.ScaleHeight - 1000 - ZeilenHöhe%) Then
            Call DruckFuss
            Call ImpKontrolleDruckKopf
        End If
    Next i%
End With
Call DruckFuss(False)
Printer.EndDoc
Printer.Orientation = OrgAusrichtung%

Call DefErrPop
End Sub

Private Sub Form_Load()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_Load")
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
Dim i%, j%, l%, k%, m%, lInd%, wi%, MaxWi%, spBreite%, ind%, lief%, cmdPlusWidth%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%, FeldInd%, FormVersatzY%
Dim iAdd%, iAdd2%
Dim h$, h2$, FormStr$, AuswertungsMonat$, AktDatum$
Dim ButtonW&, ScreenSizeWidth&, ScreenSizeHeight&
Dim Fs!
Dim OriginalBreite As Boolean
Dim rsDatum As Recordset
Dim c As Control

iEditModus = 1

Me.KeyPreview = True

Call ScreenWerte(ScreenSizeHeight&, ScreenSizeWidth&)
'ScreenSizeWidth& = (Screen.Width / Screen.TwipsPerPixelX) * Screen.TwipsPerPixelY - Screen.TwipsPerPixelX
Call wpara.InitFont(Me)

If (para.Newline) Then
    FormVersatzY% = wpara.NlCaptionY
Else
    FormVersatzY% = wpara.FrmCaptionHeight + wpara.FrmBorderHeight
End If
Me.Width = frmAction.Width
Me.Height = frmAction.Height - FormVersatzY%
Me.Left = frmAction.Left
Me.Top = frmAction.Top + FormVersatzY%
If (para.Newline) Then
'    Me.Height = Me.Height - 2 * wpara.NlCaptionY
    Me.Width = Me.Width - 2 * wpara.ButtonX
End If

ScreenSizeWidth& = Me.ScaleWidth



FieldNr%(1) = AuswertungRec.Fields("rez_gesamt").OrdinalPosition
FieldNr%(2) = AuswertungRec.Fields("abr_gesamt").OrdinalPosition
FieldNr%(3) = AuswertungRec.Fields("rez_gesamtFAM").OrdinalPosition
FieldNr%(4) = AuswertungRec.Fields("abr_gesamtFAM").OrdinalPosition
FieldNr%(5) = AuswertungRec.Fields("rez_ImpFähig").OrdinalPosition
FieldNr%(6) = AuswertungRec.Fields("abr_ImpFähig").OrdinalPosition
FieldNr%(7) = AuswertungRec.Fields("rez_ImpIst").OrdinalPosition
FieldNr%(8) = AuswertungRec.Fields("abr_ImpIst").OrdinalPosition

lblDatum.Top = wpara.TitelY
lblDatum.Left = wpara.LinksX
cmbDatum.Top = wpara.TitelY + (lblDatum.Height - cmbDatum.Height) \ 2
cmbDatum.Left = lblDatum.Left + lblDatum.Width + wpara.LinksX
With flxRezSpeicher
    .Rows = 2
    .FixedRows = 1
    If (RezSpeicherModus% = 0) Then
        .Cols = 13
        .FormatString = "|||||<Name|>Anz.Rez.|>Ges.Wert|>Rab.Wert|>Zuzahlungen|>Abrechnung|>Herst.Rab.|>Wert/Rez.|>Artikel/Rez.|>"
        Me.Caption = "Rezeptspeicher"
    Else
        .Cols = 17
        .FormatString = "|||||<Name|>ges.Jahr|>ges.Mon|>Anzahl|>FAM|>ImpFähig|>ImpIst|>Q.%|>Q.EUR|>Erspart|>Diff.|>+/-|>Saldo|>"
        Me.Caption = "Importkontrolle"
    End If
    .Rows = 1
    .Font.Size = .Font.Size + 1
    OriginalBreite = True
    Do
      If .Font.Size < 9 And .Font.Name <> "Small Fonts" Then
          .Font.Name = "Small Fonts"
          .Font.Size = 9
          picFont.FontName = .Font.Name
      End If
      If (.Font.Size - 1) <= 5 Then Exit Do
      .Font.Size = .Font.Size - 1
      picFont.FontSize = .Font.Size
      
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = picFont.TextWidth(String(2, "A"))
      '.ColWidth(3) = Me.TextWidth(String(10, "9"))   'Nummer nicht anzeigen
      .ColWidth(3) = 0
      .ColWidth(4) = picFont.TextWidth(String(2, "A"))
      
      If OriginalBreite Then
        .ColWidth(5) = picFont.TextWidth(String(25, "A"))
        OriginalBreite = False
      Else
        .ColWidth(5) = picFont.TextWidth(String(15, "A"))
      End If
      If (RezSpeicherModus% = 0) Then
          For i% = 0 To 4
            .ColWidth(i%) = 0
          Next i%
          .ColWidth(6) = picFont.TextWidth(String(9, "9"))
          .ColWidth(7) = picFont.TextWidth(String(11, "9"))
          .ColWidth(8) = picFont.TextWidth(String(11, "9"))
          .ColWidth(9) = picFont.TextWidth(String(12, "9"))
          .ColWidth(10) = picFont.TextWidth(String(11, "9"))
          .ColWidth(11) = picFont.TextWidth(String(11, "9"))
          .ColWidth(12) = picFont.TextWidth(String(11, "9"))
          .ColWidth(13) = picFont.TextWidth(String(11, "9"))
          .ColWidth(14) = wpara.FrmScrollHeight
      ElseIf DarfRezSpeicher Then
          .ColWidth(6) = picFont.TextWidth(String(12, "9"))
          .ColWidth(7) = picFont.TextWidth(String(11, "9"))
          .ColWidth(8) = picFont.TextWidth(String(7, "9"))
          .ColWidth(9) = picFont.TextWidth(String(11, "9"))
          .ColWidth(10) = picFont.TextWidth(String(11, "9"))
          .ColWidth(11) = picFont.TextWidth(String(11, "9"))
          .ColWidth(12) = picFont.TextWidth(String(4, "9"))
          .ColWidth(13) = picFont.TextWidth(String(10, "9"))
          .ColWidth(14) = picFont.TextWidth(String(9, "9"))
          .ColWidth(15) = picFont.TextWidth(String(9, "9"))
          '.ColWidth(15) = 0
          .ColWidth(16) = picFont.TextWidth(String(8, "9"))
'      ElseIf (RezSpeicherModus% = 0) Then
'          .ColWidth(6) = picFont.TextWidth(String(7, "9"))
'          .ColWidth(7) = picFont.TextWidth(String(11, "9"))
'          .ColWidth(8) = picFont.TextWidth(String(11, "9"))
'          .ColWidth(9) = picFont.TextWidth(String(11, "9"))
'          .ColWidth(10) = picFont.TextWidth(String(11, "9"))
'          .ColWidth(11) = picFont.TextWidth(String(8, "9"))
'          .ColWidth(12) = picFont.TextWidth(String(5, "9"))
      Else
          .ColWidth(5) = picFont.TextWidth(String(55, "A"))
          For i% = 6 To 15
              .ColWidth(i%) = 0
          Next i%
      End If
      If (RezSpeicherModus% = 1) Then
        .ColWidth(17) = picFont.TextWidth(String(8, "9"))
        .ColWidth(18) = wpara.FrmScrollHeight
      End If
      Breite1% = 0
      For i% = 0 To (.Cols - 1)
          Breite1% = Breite1% + .ColWidth(i%)
      Next i%
      .Width = Breite1% + 90
'      .Height = .RowHeight(0) * 20 + 90
    Loop While (.Width - 100) > ScreenSizeWidth&
    
    .Top = lblDatum.Top + lblDatum.Height + wpara.TitelY
    .Left = wpara.LinksX

    .Height = ((Me.ScaleHeight - .Top - wpara.ButtonY - 300) \ .RowHeight(0)) * .RowHeight(0) + 90

'    If (RezSpeicherModus% = 0) Then
        .Width = ScreenSizeWidth& - 2 * wpara.LinksX
        .ColWidth(5) = 0
        Breite1% = 0
        For i% = 0 To (.Cols - 1)
            Breite1% = Breite1% + .ColWidth(i%)
        Next i%
        .ColWidth(5) = .Width - Breite1% - 90
'    End If
End With

''Kombobox mit Monatswerten befüllen
'Set rsDatum = RezSpeicherDB.OpenRecordset("SELECT DISTINCT Monat FROM Auswertung ORDER BY Monat", dbOpenSnapshot, dbReadOnly)
'If Not rsDatum.EOF Then rsDatum.MoveFirst
'Do While Not rsDatum.EOF
'  'Jahreszeile einfügen
'  If Left(rsDatum!monat, 2) <> Left(h$, 2) Then
'      If h$ <> "" Then
'          cmbDatum.AddItem ("20" + Left(h$, 2) + " gesamt")
'      Else
'      End If
'  End If
'
'  If (Val(Left(h$, 2)) < 2 And Val(Left(rsDatum!monat, 2)) = 2) or (val(left(h$,2) ) Then
'    h$ = "0201"
'    Do While Val(Mid(h$, 3, 2)) < Val(Mid(rsDatum!monat, 3, 2))
'        cmbDatum.AddItem (Format(CDate("01." + Mid(h$, 3, 2) + "." + Left(h$, 2)), "mmmm YYYY"))
'        h$ = "02" + Format(Val(Mid(h$, 3, 2)) + 1, "00")
'    Loop
'  End If
'
'
'
'  h$ = rsDatum!monat
'  cmbDatum.AddItem (Format(CDate("01." + Mid(h$, 3, 2) + "." + Left(h$, 2)), "mmmm YYYY"))
'  If h$ = Format(Now, "YYMM") Then
'      cmbDatum.ListIndex = cmbDatum.ListCount - 1
'  End If
'  rsDatum.MoveNext
'Loop
'If h$ > "" Then
'  cmbDatum.AddItem ("20" + Left(h$, 2) + " gesamt")
'Else
'  cmbDatum.AddItem (Format(Now, "YYYY") + " gesamt")
'  Call cmbDatum_Click
'End If
'rsDatum.Close
'Set rsDatum = Nothing


AktDatum$ = Format(Now, "MM.YY")
If Val(Format(Now, "YYMM")) < Val(AbrechMonat$) Then
    AktDatum$ = Mid(AbrechMonat$, 3, 2) + "." + Left(AbrechMonat$, 2)
End If
j% = 2
m% = 1
Do
    h$ = Format(m%, "00") + "." + Format(j%, "00")
    cmbDatum.AddItem (Format(CDate("01." + h$), "mmmm YYYY"))
    If h$ = AktDatum$ Then
        cmbDatum.ListIndex = cmbDatum.ListCount - 1
        cmbDatum.AddItem (Format(Now, "YYYY") + " gesamt")
        Exit Do
    End If
    m% = m% + 1
    If m% > 12 Then
        cmbDatum.AddItem ("20" + Mid(h$, 4, 2) + " gesamt")
        j% = j% + 1
        m% = 1
    End If
Loop


Font.Bold = False   ' True

With cmdEsc
    .Top = flxRezSpeicher.Top + flxRezSpeicher.Height + 150
    .Width = wpara.ButtonX
    .Height = wpara.ButtonY
    .Left = flxRezSpeicher.Left + flxRezSpeicher.Width - wpara.ButtonX
    'cmdEsc.Left = (Me.Width - cmdEsc.Width) / 2
End With

picFont.FontName = cmdEsc.Font.Name
picFont.FontSize = cmdEsc.Font.Size
ButtonW& = picFont.TextWidth("Import-Alternativen (F8)") + 300
If RezSpeicherModus% = 1 Then
    Do While 5 * (wpara.LinksX + ButtonW&) > (cmdEsc.Left - wpara.LinksX)
      Fs! = picFont.FontSize
      picFont.FontSize = picFont.FontSize - 1
      If Fs! = picFont.FontSize Then Exit Do
      ButtonW& = picFont.TextWidth("Importalternativen (F8)") + 120
    Loop
End If
If (RezSpeicherModus% = 0) Then
    cmdPlusWidth = 150
Else
    cmdPlusWidth = 50
End If
With cmdKKNeu
    .Top = cmdEsc.Top
    .Width = TextWidth(.Caption) + cmdPlusWidth
'    .Width = ButtonW&
    .Height = wpara.ButtonY
    .Left = wpara.LinksX
    .FontSize = picFont.FontSize
End With
With cmdKStamm
    .Top = cmdEsc.Top
    .Width = TextWidth(.Caption) + cmdPlusWidth
'    .Width = ButtonW&
    .Height = wpara.ButtonY
    .Left = cmdKKNeu.Left + cmdKKNeu.Width + wpara.LinksX
    .FontSize = picFont.FontSize
End With
With cmdRezSpeicherLoeschen
    .Top = cmdEsc.Top
    .Width = TextWidth(.Caption) + cmdPlusWidth
'    .Width = ButtonW&
    .Height = wpara.ButtonY
    .Left = cmdKStamm.Left + cmdKStamm.Width + wpara.LinksX
    .FontSize = picFont.FontSize
End With
With cmdDruck
    .Top = cmdEsc.Top
    .Width = TextWidth(.Caption) + cmdPlusWidth
'    .Width = ButtonW&
    .Height = wpara.ButtonY
    .Left = cmdRezSpeicherLoeschen.Left + cmdRezSpeicherLoeschen.Width + wpara.LinksX
    .FontSize = picFont.FontSize
End With
With cmdAbrech
    .Top = cmdEsc.Top
    .Width = TextWidth(.Caption) + cmdPlusWidth
'    .Width = ButtonW&
    .Height = wpara.ButtonY
    .Left = cmdDruck.Left + cmdDruck.Width + wpara.LinksX
    .FontSize = picFont.FontSize
End With
With cmdImpAlt(0)
    .Top = cmdEsc.Top
    .Width = TextWidth(.Caption) + cmdPlusWidth
'    .Width = ButtonW&
    .Height = wpara.ButtonY
    .Left = cmdAbrech.Left + cmdAbrech.Width + wpara.LinksX
    .FontSize = picFont.FontSize
End With
With cmdImpAlt(1)
    .Top = cmdEsc.Top
    .Width = TextWidth(.Caption) + cmdPlusWidth
'    .Width = ButtonW&
    .Height = wpara.ButtonY
    .Left = cmdImpAlt(0).Left + cmdImpAlt(0).Width + wpara.LinksX
    .FontSize = picFont.FontSize
End With

cmdRezSpeicherLoeschen.Visible = True
If DarfRezSpeicher And (RezSpeicherModus% = 1) Then
    cmdAbrech.Enabled = True
    cmdImpAlt(0).Enabled = True
    cmdImpAlt(1).Enabled = True
Else
    cmdAbrech.Enabled = False
    cmdImpAlt(0).Enabled = False
    cmdImpAlt(1).Enabled = False
    If (RezSpeicherModus% = 0) Then
'        cmdDruck.Visible = False
        cmdAbrech.Visible = False
        cmdImpAlt(0).Visible = False
        cmdImpAlt(1).Visible = False
        cmdRezSpeicherLoeschen.Visible = True
        
'        cmdImpAlt(0).Caption = "FiveRx (F8)"
'        cmdImpAlt(0).Enabled = True
    End If
End If

If (RezSpeicherModus% = 0) Then
    With cmdAbrech
        .Visible = True
        .Enabled = True
        .Caption = "IK-Auswertung (F7)"
    End With
End If

'Me.Width = flxRezSpeicher.Left + flxRezSpeicher.Width + 2 * wpara.LinksX

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

If (para.Newline) Then
    iAdd = wpara.NlFlexBackY
    iAdd2 = wpara.NlCaptionY
    
    With nlcmdEsc
        .Init
    End With
    With flxRezSpeicher
'        .ScrollBars = flexScrollBarNone
        .BorderStyle = 0
        .Width = .Width - 90
        .Height = .Height - 90
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridFlat
        .GridColorFixed = .GridColor
        .BackColor = wpara.nlFlexBackColor    'vbWhite
        .BackColorBkg = wpara.nlFlexBackColor    'vbWhite
        .BackColorFixed = wpara.nlFlexBackColorFixed   ' RGB(199, 176, 123)
        .BackColorSel = wpara.nlFlexBackColorSel  ' RGB(232, 217, 172)
        .ForeColorSel = vbBlack
        
        .Left = .Left + iAdd
        .Top = .Top + iAdd
        
        .Height = (Me.ScaleHeight - .Top - (iAdd + 600 + nlcmdEsc.Height + wpara.FrmCaptionHeight + 450))
        .Height = (.Height \ .RowHeight(0)) * .RowHeight(0)
    End With
    
    cmdEsc.Top = cmdEsc.Top + 2 * iAdd
    
    Width = Width + 2 * iAdd
    Height = Height + 2 * iAdd

    On Error Resume Next
    For Each c In Controls
        If (c.Container Is Me) Then
            c.Top = c.Top + iAdd2
        End If
    Next
    On Error GoTo DefErr
    
    
    Height = Height + iAdd2
    
    With nlcmdEsc
        .Init
        .Left = Me.ScaleWidth - .Width - 150
        .Top = flxRezSpeicher.Top + flxRezSpeicher.Height + iAdd + 600
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .Default = cmdEsc.Default
        .Cancel = cmdEsc.Cancel
        .Visible = True
    End With
    cmdEsc.Visible = False

    With nlcmdKKNeu
        .Init
        .AutoSize = True
        .Left = cmdKKNeu.Left
        .Top = nlcmdEsc.Top
        .Caption = cmdKKNeu.Caption
        .TabIndex = cmdKKNeu.TabIndex
        .Enabled = cmdKKNeu.Enabled
        .Default = cmdKKNeu.Default
        .Cancel = cmdKKNeu.Cancel
        .Visible = True
    End With
    cmdKKNeu.Visible = False

    With nlcmdKStamm
        .Init
        .AutoSize = True
        .Left = nlcmdKKNeu.Left + nlcmdKKNeu.Width + 60
        .Top = nlcmdEsc.Top
        .Caption = cmdKStamm.Caption
        .TabIndex = cmdKStamm.TabIndex
        .Enabled = cmdKStamm.Enabled
        .Default = cmdKStamm.Default
        .Cancel = cmdKStamm.Cancel
        .Visible = True
    End With
    cmdKStamm.Visible = False
    
    With nlcmdRezSpeicherLoeschen
        .Init
        .AutoSize = True
        .Left = nlcmdKStamm.Left + nlcmdKStamm.Width + 60
        .Top = nlcmdEsc.Top
        .Caption = cmdRezSpeicherLoeschen.Caption
        .TabIndex = cmdRezSpeicherLoeschen.TabIndex
        .Enabled = cmdRezSpeicherLoeschen.Enabled
        .Default = cmdRezSpeicherLoeschen.Default
        .Cancel = cmdRezSpeicherLoeschen.Cancel
        .Visible = True
    End With
    cmdRezSpeicherLoeschen.Visible = False
    
    With nlcmdDruck
        .Init
        .AutoSize = True
        .Left = nlcmdRezSpeicherLoeschen.Left + nlcmdRezSpeicherLoeschen.Width + 60
        .Top = nlcmdEsc.Top
        .Caption = cmdDruck.Caption
        .TabIndex = cmdDruck.TabIndex
        .Enabled = cmdDruck.Enabled
        .Default = cmdDruck.Default
        .Cancel = cmdDruck.Cancel
        .Visible = True
    End With
    cmdDruck.Visible = False
    
    With nlcmdAbrech
        .Init
        .AutoSize = True
        .Left = nlcmdDruck.Left + nlcmdDruck.Width + 60
        .Top = nlcmdEsc.Top
        .Caption = cmdAbrech.Caption
        .TabIndex = cmdAbrech.TabIndex
        .Enabled = cmdAbrech.Enabled
        .Default = cmdAbrech.Default
        .Cancel = cmdAbrech.Cancel
        .Visible = True
    End With
    cmdAbrech.Visible = False

    With nlcmdImpAlt(0)
        .Init
        .AutoSize = True
        .Left = nlcmdAbrech.Left + nlcmdAbrech.Width + 60
        .Top = nlcmdEsc.Top
        .Caption = cmdImpAlt(0).Caption
        .TabIndex = cmdImpAlt(0).TabIndex
        .Enabled = cmdImpAlt(0).Enabled
        .Default = cmdImpAlt(0).Default
        .Cancel = cmdImpAlt(0).Cancel
        .Visible = True
    End With
    cmdImpAlt(0).Visible = False
    
    With nlcmdImpAlt(1)
        .Init
        .AutoSize = True
        .Left = nlcmdImpAlt(0).Left + nlcmdImpAlt(0).Width + 60
        .Top = nlcmdEsc.Top
        .Caption = cmdImpAlt(1).Caption
        .TabIndex = cmdImpAlt(1).TabIndex
        .Enabled = cmdImpAlt(1).Enabled
        .Default = cmdImpAlt(1).Default
        .Cancel = cmdImpAlt(1).Cancel
        .Visible = True
    End With
    cmdImpAlt(1).Visible = False
    
    nlcmdRezSpeicherLoeschen.Visible = True
    If DarfRezSpeicher And (RezSpeicherModus% = 1) Then
        nlcmdAbrech.Enabled = True
        nlcmdImpAlt(0).Enabled = True
        nlcmdImpAlt(1).Enabled = True
    Else
        nlcmdAbrech.Enabled = False
        nlcmdImpAlt(0).Enabled = False
        nlcmdImpAlt(1).Enabled = False
        If (RezSpeicherModus% = 0) Then
'            nlcmdDruck.Visible = False
            nlcmdAbrech.Visible = False
            nlcmdImpAlt(0).Visible = False
            nlcmdImpAlt(1).Visible = False
            nlcmdRezSpeicherLoeschen.Visible = True
            
    '        cmdImpAlt(0).Caption = "FiveRx (F8)"
    '        cmdImpAlt(0).Enabled = True
        End If
    End If
    
    If (RezSpeicherModus% = 0) Then
        With nlcmdAbrech
            .Visible = True
            .Enabled = True
            .Caption = "IK-Auswertung (F7)"
        End With
    End If
    
    Me.Height = nlcmdEsc.Top + nlcmdEsc.Height + wpara.FrmCaptionHeight + 450

    Call wpara.NewLineWindow(Me, nlcmdEsc.Top)
    With flxRezSpeicher
        RoundRect hdc, (.Left - iAdd) / Screen.TwipsPerPixelX, (.Top - iAdd) / Screen.TwipsPerPixelY, (.Left + .Width + iAdd) / Screen.TwipsPerPixelX, (.Top + .Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
    End With

    On Error Resume Next
    For Each c In Controls
        If (c.tag <> "0") Then
            If (TypeOf c Is Label) Then
                c.BackStyle = 0 'duchsichtig
            ElseIf (TypeOf c Is TextBox) Or (TypeOf c Is ComboBox) Then
                If (TypeOf c Is ComboBox) Then
                    Call wpara.ControlBorderless(c)
                ElseIf (c.Appearance = 1) Then
                    Call wpara.ControlBorderless(c, 2, 2)
                Else
                    Call wpara.ControlBorderless(c, 1, 1)
                End If

                If (c.Enabled) Then
                    c.BackColor = vbWhite
                Else
                    c.BackColor = Me.BackColor
                End If

'                If (c.Visible) Then
                    With c.Container
                        .ForeColor = RGB(180, 180, 180) ' vbWhite
                        .FillStyle = vbSolid
                        .FillColor = c.BackColor

                        RoundRect .hdc, (c.Left - 60) / Screen.TwipsPerPixelX, (c.Top - 30) / Screen.TwipsPerPixelY, (c.Left + c.Width + 60) / Screen.TwipsPerPixelX, (c.Top + c.Height + 15) / Screen.TwipsPerPixelY, 10, 10
                    End With
'                End If
'            ElseIf (TypeOf c Is CheckBox) Then
'                c.Height = 0
'                c.Width = c.Height
'                If (c.Name = "chkHistorie") Then
'                    If (c.Index > 0) Then
'                        Load lblchkHistorie(c.Index)
'                    End If
'                    With lblchkHistorie(c.Index)
'                        .BackStyle = 0 'duchsichtig
'                        .Caption = c.Caption
'                        .Left = c.Left + 300
'                        .Top = c.Top
'                        .Width = TextWidth(.Caption) + 90
'                        .TabIndex = c.TabIndex
'                        .Visible = True
'                    End With
'                End If
            End If
        End If
    Next
    On Error GoTo DefErr
    
Else
    nlcmdEsc.Visible = False
    nlcmdAbrech.Visible = False
    nlcmdDruck.Visible = False
    nlcmdImpAlt(0).Visible = False
    nlcmdImpAlt(1).Visible = False
    nlcmdKKNeu.Visible = False
    nlcmdKStamm.Visible = False
    nlcmdRezSpeicherLoeschen.Visible = False
End If
'''''''''

'Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
'If Me.Left < 0 Then Me.Left = -flxRezSpeicher.Left
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2
'Me.Top = frmAction.Top + wpara.FrmCaptionHeight + wpara.FrmBorderHeight

Call DefErrPop
End Sub

Private Sub Form_Paint()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_Paint")
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
Dim i%, spBreite%, ind%, iAnzZeilen%, RowHe%, bis%, bis2%
Dim sp&
Dim h$, h2$
Dim iAdd%, iAdd2%, wi%
Dim c As Control

If (para.Newline) Then
    iAdd = wpara.NlFlexBackY
    iAdd2 = wpara.NlCaptionY
    
    Call wpara.NewLineWindow(Me, nlcmdEsc.Top, False)
    With flxRezSpeicher
        RoundRect hdc, (.Left - iAdd) / Screen.TwipsPerPixelX, (.Top - iAdd) / Screen.TwipsPerPixelY, (.Left + .Width + iAdd) / Screen.TwipsPerPixelX, (.Top + .Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
    End With

    Call Form_Resize
End If

Call DefErrPop
End Sub

Private Sub flxRezSpeicherBefuellen(AuswertungsMonat$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxRezSpeicherBefuellen")
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
Dim i%, j%
ReDim Werte#(8), pWerte#(0), pWerte1#(0), Wert#(0), typw#(MAX_KKTYP, 7)
Dim h$, h2$, Sort$, kkey$
Dim Quote#, ImpSoll#, diff#, Saldo#, SaldoV#, SaldoGes#, GutHaben#, dWert#
Dim RabWert#, RabWerte#, pRabWerte#, HerstRabatt#, HerstRabatte#
Dim ActRecNo&, l&, RezAnz&, AnzGes&
'Dim TypSatz(MAX_KKTYP) As Variant
Dim AbrechnungDa As Boolean, Typen As Boolean

AuswertungsDatum$ = AuswertungsMonat$

flxRezSpeicher.Redraw = False
flxRezSpeicher.Rows = 1

RabWerte# = 0
pRabWerte# = 0
HerstRabatte = 0

With KasseRec
    On Error Resume Next
    .MoveFirst
    On Error GoTo DefErr
    Do While Not .EOF
        'Summen für Kassentypen vorab unter Kommentar
        'If KasseRec!Nummer <= MAX_KKTYP Then
        '    TypSatz(KasseRec!typ) = .Bookmark
        'Else
            GoSub Kasse
        'End If
        .MoveNext
    Loop
    .index = "Unique"
    .Seek ">=", "000000001"
    
End With

'Typen = True
'For i% = 1 To MAX_KKTYP
'    If Not IsEmpty(TypSatz(i%)) Then
'        KasseRec.Bookmark = TypSatz(i%)
'        GoSub Kasse
'    End If
'Next i%

'nur Werte holen - keine Anzeige für "unbekannt"
'h2$ = GetRezeptSpeicher$("Unbekannt", AuswertungsMonat$, Wert#(), RabWert#, HerstRabatt#, RezAnz&, GutHaben#, Saldo#, SaldoV#, AbrechnungDa)

RabWerte = 0
HerstRabatte = 0
For j% = 0 To 2
    Werte(j) = 0
Next j
h2$ = GetRezeptSpeicher$("Alle Kk", AuswertungsMonat$, Wert#(), RabWert#, HerstRabatt#, RezAnz&, GutHaben#, Saldo#, SaldoV#, AbrechnungDa)

h2$ = ""
If (RezSpeicherModus% = 0) Then
    RabWerte# = RabWerte# + RabWert#
    HerstRabatte# = HerstRabatte# + HerstRabatt#
    
    For j% = 0 To 2
        h2$ = h2$ + vbTab
        
        Werte#(j%) = Werte#(j%) + Wert#(j%)
        
        If (Werte#(j%) <> 0#) Then
            If (j% = 0) Then
                h2$ = h2$ + Format(Werte#(j%), "0")
            Else
                h2$ = h2$ + Format(Werte#(j%), "0.00")
            End If
        End If
        
        If (j% = 1) Then
            h2$ = h2$ + vbTab
'            dWert# = Werte#(1) / VmRabattFaktor#
            dWert# = Werte#(1) - RabWerte#
            If (dWert# <> 0#) Then
                h2$ = h2$ + Format(dWert#, "0.00")
            End If
        ElseIf (j% = 2) Then
            h2$ = h2$ + vbTab
            dWert# = dWert# - Werte#(2)
            If (dWert# <> 0#) Then
                h2$ = h2$ + Format(dWert#, "0.00")
            End If
            
            h2$ = h2$ + vbTab
            dWert# = HerstRabatte
            If (dWert# <> 0#) Then
                h2$ = h2$ + Format(dWert#, "0.00")
            End If
                        
            h2$ = h2$ + vbTab
            If (Werte#(0) > 0) Then
                dWert# = Werte#(1) / Werte#(0)
                If (dWert# <> 0#) Then
                    h2$ = h2$ + Format(dWert#, "0.00")
                End If
            End If
            
            Werte#(3) = Werte#(3) + Wert#(3)
        
            h2$ = h2$ + vbTab
            If (Werte#(0) > 0) Then
                dWert# = Werte#(3) / Werte#(0)
                If (dWert# <> 0#) Then
                    h2$ = h2$ + Format(dWert#, "0.00")
                End If
            End If
        End If
    Next j%
            
    SaldoGes# = 0#
Else
    For i% = 0 To 4
        If i% = 2 Then
            AnzGes& = AnzGes& + RezAnz&
            h2$ = h2$ + vbTab
            If AnzGes& > 0 Then h2$ = h2$ + Format(AnzGes&, "0")
        End If
        Werte#(i%) = Werte#(i%) + Wert#(i%)
        h2$ = h2$ + vbTab
        If (Werte#(i%) <> 0) Then h2$ = h2$ + Format(Werte#(i%), "0.00")
    Next i%
    If Saldo# = 0# And Not AbrechnungDa Then
        Saldo# = diff# + SaldoV#
'        Saldo# = diff#
'        If diff# < 0# And SaldoV# > 0 Then
'          Saldo# = diff# + SaldoV#
'        End If
    End If
    SaldoGes# = SaldoGes# + Saldo#
End If

h$ = vbTab + vbTab + "AlleKK   " + vbTab + vbTab + "Kassenrezepte" + h2$
h$ = h$ + vbTab + vbTab + vbTab + Format(Werte#(W_IMPERSPART), "0.00") + vbTab
If SaldoGes# > 0# Then
    h$ = h$ + vbTab + Format(SaldoGes#, "0.00")
End If
Sort$ = "1"
h$ = Sort$ + vbTab + h$
flxRezSpeicher.AddItem h$



h$ = vbTab + vbTab + "Privat VK" + vbTab + vbTab + "Privat-Rezepte VK"
h2$ = GetRezeptSpeicher$("Privat VK", AuswertungsMonat$, pWerte#(), RabWert#, HerstRabatt#, RezAnz&, GutHaben#, Saldo#, SaldoV#, AbrechnungDa)
h2$ = ""
If (RezSpeicherModus% = 0) Then
    pRabWerte# = pRabWerte# + RabWert#
    HerstRabatte# = HerstRabatte# + HerstRabatt#
    For j% = 0 To 3
        h2$ = h2$ + vbTab
        
        If (pWerte#(j%) <> 0#) Then
            If (j% = 0) Then
                h2$ = h2$ + Format(pWerte#(j%), "0")
            Else
                h2$ = h2$ + Format(pWerte#(j%), "0.00")
            End If
        End If
        
        If (j% = 1) Then
            h2$ = h2$ + vbTab
'            dWert# = pWerte#(1) / VmRabattFaktor#
            dWert# = pWerte#(1) - RabWert#
            If (dWert# <> 0#) Then
                h2$ = h2$ + Format(dWert#, "0.00")
            End If
        ElseIf (j% = 2) Then
            h2$ = h2$ + vbTab
            dWert# = dWert# - pWerte#(2)
            If (dWert# <> 0#) Then
                h2$ = h2$ + Format(dWert#, "0.00")
            End If
            
            h2$ = h2$ + vbTab
            dWert# = HerstRabatt
            If (dWert# <> 0#) Then
                h2$ = h2$ + Format(dWert#, "0.00")
            End If
                        
                        
            h2$ = h2$ + vbTab
            If (pWerte#(0) > 0) Then
                dWert# = pWerte#(1) / pWerte#(0)
                If (dWert# <> 0#) Then
                    h2$ = h2$ + Format(dWert#, "0.00")
                End If
            End If
            
            h2$ = h2$ + vbTab
            If (pWerte#(0) > 0) Then
                dWert# = pWerte#(3) / pWerte#(0)
                If (dWert# <> 0#) Then
                    h2$ = h2$ + Format(dWert#, "0.00")
                End If
            End If
        End If
    Next j%
Else
    For i% = 0 To 4
        If i% = 2 Then
            h2$ = h2$ + vbTab
            If RezAnz& > 0 Then h2$ = h2$ + Format(RezAnz&, "0")
            AnzGes& = AnzGes& + RezAnz&
        End If
        h2$ = h2$ + vbTab
        If (pWerte#(i%) > 0) Then h2$ = h2$ + Format(pWerte#(i%), "0.00")
        'h2$ = h2$ + Format(9999999.99, "0.00")
    Next i%
End If

h$ = h$ + h2$
Sort$ = "2"
h$ = Sort$ + vbTab + h$
flxRezSpeicher.AddItem h$



h$ = vbTab + vbTab + "Privat   " + vbTab + vbTab + "Privat-Rezepte gedruckt"
h2$ = GetRezeptSpeicher$("Privat   ", AuswertungsMonat$, pWerte1#(), RabWert#, HerstRabatt#, RezAnz&, GutHaben#, Saldo#, SaldoV#, AbrechnungDa)
h2$ = ""
If (RezSpeicherModus% = 0) Then
    pRabWerte# = pRabWerte# + RabWert#
    HerstRabatte# = HerstRabatte# + HerstRabatt#
    
    For j% = 0 To 3
        h2$ = h2$ + vbTab
        
        If (pWerte1#(j%) <> 0#) Then
            If (j% = 0) Then
                h2$ = h2$ + Format(pWerte1#(j%), "0")
            Else
                h2$ = h2$ + Format(pWerte1#(j%), "0.00")
            End If
        End If
        
        If (j% = 1) Then
            h2$ = h2$ + vbTab
'            dWert# = pWerte1#(1) / VmRabattFaktor#
            dWert# = pWerte1#(1) - RabWert#
            If (dWert# <> 0#) Then
                h2$ = h2$ + Format(dWert#, "0.00")
            End If
        ElseIf (j% = 2) Then
            h2$ = h2$ + vbTab
            dWert# = dWert# - pWerte1#(2)
            If (dWert# <> 0#) Then
                h2$ = h2$ + Format(dWert#, "0.00")
            End If
            
            h2$ = h2$ + vbTab
            dWert# = HerstRabatt
            If (dWert# <> 0#) Then
                h2$ = h2$ + Format(dWert#, "0.00")
            End If
                        
            h2$ = h2$ + vbTab
            If (pWerte1#(0) > 0) Then
                dWert# = pWerte1#(1) / pWerte1#(0)
                If (dWert# <> 0#) Then
                    h2$ = h2$ + Format(dWert#, "0.00")
                End If
            End If
            
            h2$ = h2$ + vbTab
            If (pWerte1#(0) > 0) Then
                dWert# = pWerte1#(3) / pWerte1#(0)
                If (dWert# <> 0#) Then
                    h2$ = h2$ + Format(dWert#, "0.00")
                End If
            End If
        End If
    Next j%
Else
    For i% = 0 To 4
        If i% = 2 Then
            h2$ = h2$ + vbTab
            If RezAnz& > 0 Then h2$ = h2$ + Format(RezAnz&, "0")
        End If
        h2$ = h2$ + vbTab
        If (pWerte1#(i%) > 0) Then h2$ = h2$ + Format(pWerte1#(i%), "0.00")
        'h2$ = h2$ + Format(9999999.99, "0.00")
    Next i%
End If

h$ = h$ + h2$
Sort$ = "3"
h$ = Sort$ + vbTab + h$
flxRezSpeicher.AddItem h$




h2$ = ""
If (RezSpeicherModus% = 0) Then
    For j% = 0 To 2
        h2$ = h2$ + vbTab
        
        Werte#(j%) = Werte#(j%) + pWerte#(j%) + pWerte1#(j%)
        
        If (Werte#(j%) <> 0#) Then
            If (j% = 0) Then
                h2$ = h2$ + Format(Werte#(j%), "0")
            Else
                h2$ = h2$ + Format(Werte#(j%), "0.00")
            End If
        End If
    
        
        If (j% = 1) Then
            h2$ = h2$ + vbTab
'            dWert# = Werte#(1) / VmRabattFaktor#
            dWert# = Werte#(1) - (RabWerte# + pRabWerte#)
            If (dWert# <> 0#) Then
                h2$ = h2$ + Format(dWert#, "0.00")
            End If
        ElseIf (j% = 2) Then
            h2$ = h2$ + vbTab
            dWert# = dWert# - Werte#(2)
            If (dWert# <> 0#) Then
                h2$ = h2$ + Format(dWert#, "0.00")
            End If
            
            h2$ = h2$ + vbTab
            dWert# = HerstRabatte
            If (dWert# <> 0#) Then
                h2$ = h2$ + Format(dWert#, "0.00")
            End If
            
            h2$ = h2$ + vbTab
            If (Werte#(0) > 0) Then
                dWert# = Werte#(1) / Werte#(0)
                If (dWert# <> 0#) Then
                    h2$ = h2$ + Format(dWert#, "0.00")
                End If
            End If
            
            Werte#(3) = Werte#(3) + pWerte#(3)
            
            h2$ = h2$ + vbTab
            If (Werte#(0) > 0) Then
                dWert# = Werte#(3) / Werte#(0)
                If (dWert# <> 0#) Then
                    h2$ = h2$ + Format(dWert#, "0.00")
                End If
            End If
        End If
    Next j%
            
    SaldoGes# = 0#
Else
    For i% = 0 To 4
        If i% = 2 Then
            h2$ = h2$ + vbTab
            If AnzGes& > 0 Then h2$ = h2$ + Format(AnzGes&, "0")
        End If
        h2$ = h2$ + vbTab
        If (Werte#(i%) + pWerte#(i%) <> 0) Then h2$ = h2$ + Format(Werte#(i%) + pWerte#(i%), "0.00")
        'h2$ = h2$ + Format(9999999.99, "0.00")
    Next i%
End If

h$ = vbTab + vbTab + vbTab + vbTab + "alle Rezepte" + h2$
If SaldoGes# > 0# Then
    h$ = h$ + vbTab + vbTab + vbTab + vbTab + vbTab + Format(SaldoGes#, "0.00")
End If
Sort$ = "0"
h$ = Sort$ + vbTab + h$
flxRezSpeicher.AddItem h$


With flxRezSpeicher
  .FillStyle = flexFillRepeat
  .row = 1
  .col = 2
  .RowSel = .Rows - 1
  .ColSel = 1
  .CellFontName = "Symbol"
  .FillStyle = flexFillSingle

  .row = 1
  .col = 0
  .RowSel = .Rows - 1
  .ColSel = 1
  .Sort = 5
  .col = 0
  .ColSel = .Cols - 1
  .Redraw = True

    If (RezSpeicherModus% <> 0) Then
        For l& = 1 To .Rows - 1
            If xVal(.TextMatrix(l&, 13)) < 0# Then
                .row = l&
                .col = 13
                .CellForeColor = vbRed
            End If
            If xVal(.TextMatrix(l&, 14)) < 0# Then
                .row = l&
                .col = 14
                .CellForeColor = vbRed
            End If
        Next l&
        .row = 1
    End If
End With

Call DefErrPop
Exit Sub
'------------------------------------------------------------------------------------------------------------------------
Kasse:
'FabsErrf% = Kkasse.IndexSearch(0, KasseRec!Nummer, FabsRecno&)
If FabsRecno& = 0 Then FabsRecno& = -1

If (KasseRec!Anzeige) Then
    h2$ = Chr$(214)
    Sort$ = "4"
Else
    h2$ = " "
    Sort$ = "5"
End If

h$ = Format(FabsRecno&, "0") + vbTab + h2$
h$ = h$ + vbTab + KasseRec!Nummer + vbTab

h2$ = GetRezeptSpeicher$(KasseRec!Nummer, AuswertungsMonat$, Wert#(), RabWert#, HerstRabatt#, RezAnz&, GutHaben#, Saldo#, SaldoV#, AbrechnungDa)
If AbrechnungDa Then h$ = h$ + "*"
h$ = h$ + vbTab + KasseRec!Name

h2$ = ""
If (RezSpeicherModus% = 0) Then
    RabWerte# = RabWerte# + RabWert#
    HerstRabatte# = HerstRabatte# + HerstRabatt#
    For j% = 0 To 3
        h2$ = h2$ + vbTab
        
        Werte#(j%) = Werte#(j%) + Wert#(j%)
        typw#(KasseRec!Typ, j%) = typw#(KasseRec!Typ, j%) + Wert#(j%)
        
        If (Wert#(j%) <> 0#) Then
            If (j% = 0) Then
                h2$ = h2$ + Format(Wert#(j%), "0")
            Else
                h2$ = h2$ + Format(Wert#(j%), "0.00")
            End If
        End If
        
        If (j% = 1) Then
            h2$ = h2$ + vbTab
'            dWert# = Wert#(1) / VmRabattFaktor#
            dWert# = Wert#(1) - RabWert#
            If (dWert# <> 0#) Then
                h2$ = h2$ + Format(dWert#, "0.00")
            End If
        ElseIf (j% = 2) Then
            h2$ = h2$ + vbTab
            dWert# = dWert# - Wert#(2)
            If (dWert# <> 0#) Then
                h2$ = h2$ + Format(dWert#, "0.00")
            End If
            
            h2$ = h2$ + vbTab
            dWert# = HerstRabatt
            If (dWert# <> 0#) Then
                h2$ = h2$ + Format(dWert#, "0.00")
            End If
            
            h2$ = h2$ + vbTab
            If (Wert#(0) > 0) Then
                dWert# = Wert#(1) / Wert#(0)
                If (dWert# <> 0#) Then
                    h2$ = h2$ + Format(dWert#, "0.00")
                End If
            End If
            
            h2$ = h2$ + vbTab
            If (Wert#(0) > 0) Then
                dWert# = Wert#(3) / Wert#(0)
                If (dWert# <> 0#) Then
                    h2$ = h2$ + Format(dWert#, "0.00")
                End If
            End If
        End If
    Next j%
            
    h$ = h$ + h2$
    Sort$ = "5" + KasseRec!Name
    h$ = Sort$ + vbTab + h$
Else
    For j% = 0 To 4
        If j% = 2 Then
            'Rezeptanzahl einfügen
            AnzGes& = AnzGes& + RezAnz&
            typw#(KasseRec!Typ, W_REZANZ) = typw#(KasseRec!Typ, W_REZANZ) + RezAnz&
            h2$ = h2$ + vbTab
            If Typen Then
                If typw#(KasseRec!Typ, W_REZANZ) > 0# Then h2$ = h2$ + Format(typw#(KasseRec!Typ, W_REZANZ), "0")
            Else
                If RezAnz& > 0 Then h2$ = h2$ + Format(RezAnz&, "0")
            End If
        End If
        Werte#(j%) = Werte#(j%) + Wert#(j%)
        typw#(KasseRec!Typ, j%) = typw#(KasseRec!Typ, j%) + Wert#(j%)
        h2$ = h2$ + vbTab
        If Typen Then
            Wert#(j%) = typw#(KasseRec!Typ, j%)
            If typw#(KasseRec!Typ, j%) <> 0# Then h2$ = h2$ + Format(typw#(KasseRec!Typ, j%), "0.00")
        Else
            If Wert#(j%) > 0# Then h2$ = h2$ + Format(Wert#(j%), "0.00")
        End If
    Next j%
    h$ = h$ + h2$
    
    
    'Sort$ = Sort$ + Right(Space(10) + Format(Saldo#, "0.00"), 10)
    Sort$ = Sort$ + KasseRec!Name
    Quote# = ImportQuote#(Wert#(W_FAM), Wert#(W_IMPFÄHIG), Left$(AuswertungsMonat$, 4))
    h$ = h$ + vbTab
    If Quote# <> 0# Then h$ = h$ + Format(Quote#, "0.0")
    
    ImpSoll# = Wert#(W_FAM) * (Quote# / 100#) / 10#
    h$ = h$ + vbTab
    If ImpSoll# <> 0# Then h$ = h$ + Format(ImpSoll#, "0.00")
    
    h$ = h$ + vbTab
    
    If Wert#(W_IMPERSPART) <> 0# Then
      h$ = h$ + Format(Wert#(W_IMPERSPART), "0.00")
      Werte#(W_IMPERSPART) = Werte#(W_IMPERSPART) + Wert#(W_IMPERSPART)
    End If
    h$ = h$ + vbTab
    
    diff# = Wert#(W_IMPERSPART) - ImpSoll#
    If diff# <> 0# Then h$ = h$ + Format(diff#, "0.00")
    
    h$ = h$ + vbTab
    If GutHaben# <> 0# Then
        typw#(KasseRec!Typ, W_ABZUG) = typw#(KasseRec!Typ, W_ABZUG) + GutHaben#
        If Typen Then GutHaben# = typw#(KasseRec!Typ, W_ABZUG)
        h$ = h$ + Format(GutHaben#, "0.00")
    Else
'        diff# = diff# / 10#
        If Abs(diff#) > 5# Then
            typw#(KasseRec!Typ, W_ABZUG) = typw#(KasseRec!Typ, W_ABZUG) + diff#
            If Typen Then diff# = typw#(KasseRec!Typ, W_ABZUG)
            h$ = h$ + Format(diff#, "0.00")
        End If
    End If
    h$ = h$ + vbTab
    If Saldo# = 0# And Not AbrechnungDa Then
        Saldo# = diff# + SaldoV#
'        Saldo# = diff#
'        If diff# <= 0# And SaldoV# > 0 Then
 '         Saldo# = diff# + SaldoV#
 '       End If
    End If
    If Saldo# > 0# Then
        typw#(KasseRec!Typ, W_SALDO) = typw#(KasseRec!Typ, W_SALDO) + Saldo#
        If Typen Then Saldo# = typw#(KasseRec!Typ, W_SALDO)
        h$ = h$ + Format(Saldo, "0.00")
    End If
    h$ = Sort$ + vbTab + h$
End If

flxRezSpeicher.AddItem h$
If Saldo# <> 0# Then
  SaldoGes# = SaldoGes# + Saldo#
End If

Return

Call DefErrPop
End Sub

Private Function GetRezeptSpeicher$(kNummer$, kMonat$, Werte#(), RabWert#, HerstRabatt#, RezAnz&, GutHaben#, Saldo#, SaldoV#, Abrechnung As Boolean)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("GetRezeptSpeicher$")
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
Dim ret$, AktMo$, LastMo$
Dim i%, iOk%
Dim Jahr As Boolean, AbrechVorMon As Boolean
Dim bmk As Variant
Dim w#(4)

If (RezSpeicherModus% = 0) Then
    RabWert# = 0
    HerstRabatt = 0
    
    ReDim Werte#(8)
    With RezepteRec
        If Len(kMonat$) = 2 Then Jahr = True
        If (UCase(Left(kNummer, 7)) = "ALLE KK") Then
            .index = "DruckDatum"
            .Seek ">=", kMonat
        Else
            .index = "KasseDruck"
            .Seek ">=", kNummer$, kMonat$
        End If
        
        
        If Not .NoMatch Then
            Do While Not .EOF
                If (UCase(Left(kNummer, 7)) = "ALLE KK") Then
                    If (Left(RezepteRec!DruckDatum, Len(kMonat$)) <> kMonat$) Then Exit Do
                    iOk = (Left$(RezepteRec!Kkasse, 7) <> "Privat ")
                Else
                    If (RezepteRec!Kkasse <> kNummer$) Or (Left(RezepteRec!DruckDatum, Len(kMonat$)) <> kMonat$) Then Exit Do
                    iOk = True
                End If
                
                If (iOk) Then
                    Werte#(0) = Werte#(0) + 1
                    Werte#(1) = Werte#(1) + RezepteRec!RezSumme
                    Werte#(3) = Werte#(3) + RezepteRec!RezGebSumme
                    Werte#(6) = Werte#(6) + RezepteRec!AnzArtikel
                    Werte#(8) = Werte#(8) + RezepteRec!ImpErspart
                    
                    If (IsNull(RezepteRec!RabattWert)) Then
                    Else
                        RabWert# = RabWert# + RezepteRec!RabattWert
                    End If
                    
                    If (IsNull(RezepteRec!HerstRabatt)) Then
                    Else
                        HerstRabatt# = HerstRabatt# + RezepteRec!HerstRabatt
                    End If
                End If
                
                .MoveNext
            Loop
            
            ret$ = ""
            For i% = 0 To 6
              ret$ = ret$ + vbTab
              If (Werte#(i%) <> 0#) Then
                If (i% = 0) Then
                    ret$ = ret$ + Format(Werte#(i%), "0")
                Else
                    ret$ = ret$ + Format(Werte#(i%), "0.00")
                End If
              End If
            Next i%
            
            Werte#(2) = Werte#(3)
            Werte#(3) = Werte#(6)
        Else
            ret$ = vbTab + vbTab + vbTab + vbTab + vbTab + vbTab + vbTab
        End If
    End With
    GetRezeptSpeicher$ = ret$
    
    Call DefErrPop
    Exit Function
End If

Saldo# = 0#
SaldoV# = 0#
GutHaben# = 0#
RezAnz& = 0#
Abrechnung = False
AktMo$ = Format(Now, "YYMM")
ReDim Werte#(8)
With AuswertungRec
    .index = "Unique"
    If Len(kMonat$) = 2 Then Jahr = True
    .Seek ">=", kNummer$, Left(kMonat$, 2)
    If Not .NoMatch Then
        Do While Not AuswertungRec.EOF
            LastMo$ = AuswertungRec!Monat
            If AuswertungRec!Kkasse <> kNummer$ Or Left(AuswertungRec!Monat, 2) <> Left(kMonat$, 2) Then Exit Do
            If Jahr Or AuswertungRec!Monat = kMonat$ Then
                GoSub satz
            Else
                GoSub JahrSumme
            End If
            .MoveNext
        Loop
        If Not Jahr Then
            If Val(Mid(kMonat$, 3, 2)) = 1 Then
                .Seek "<=", kNummer$, Format(Val(Left(kMonat$, 2)) - 1, "00") + "12"
            Else
                .Seek "<=", kNummer$, Left(kMonat$, 2) + Format(Val(Mid(kMonat$, 3, 2)) - 1, "00")
            End If
            If Not .NoMatch Then
                If AuswertungRec!Kkasse = kNummer$ Then
                    SaldoV# = dCheckNull(AuswertungRec!Saldo)
                    If SaldoV# = 0 Then GoSub SaldoVorMonRechnen
                End If
            End If
        End If
        ret$ = ""
        For i% = 0 To 4
            ret$ = ret$ + vbTab
            If Werte#(i%) <> 0# Then ret$ = ret$ + Format(Werte#(i%), "0.00")
        Next i%
    Else
        ret$ = vbTab + vbTab + vbTab + vbTab
    End If
End With
If Jahr Then Abrechnung = False
GetRezeptSpeicher$ = ret$

Call DefErrPop
Exit Function
'------------------------------------------------------------------------------------------------------------------------
satz:

For i% = 1 To 4
    'fieldnr(1) = rez_gesamt, 3 = rez_gesamtFAM, 5 = rez_ImpFähig, 7 = rez_ImpIst
    If (dCheckNull(AuswertungRec.Fields(FieldNr%(2 * i%))) = 0#) Then
        Werte#(i%) = Werte#(i%) + dCheckNull(AuswertungRec.Fields(FieldNr%(2 * i% - 1)))
    Else
        Werte#(i%) = Werte#(i%) + AuswertungRec.Fields(FieldNr%(2 * i%))
        If i% = 1 Then Abrechnung = True
    End If
Next i%
If Abrechnung Then
  Werte#(W_IMPERSPART) = Werte#(W_IMPERSPART) + dCheckNull(AuswertungRec!abr_ImpErspart)
Else
  Werte#(W_IMPERSPART) = Werte#(W_IMPERSPART) + dCheckNull(AuswertungRec!ImpErspart)
End If

GutHaben# = dCheckNull(AuswertungRec!GutHaben)
If GutHaben# <> 0# Then Abrechnung = True
RezAnz& = dCheckNull(AuswertungRec!RezAnzahl)
Saldo# = dCheckNull(AuswertungRec!Saldo)
'If Saldo# = 0 And Not Abrechnung Then
'    bmk = AuswertungRec.Bookmark
'    If Not AuswertungRec.BOF Then
'      AuswertungRec.MovePrevious
'      If Not AuswertungRec.BOF Then
'            If AuswertungRec!Kkasse = kNummer$ And Val(AuswertungRec!Monat) = Val(LastMo$) - 1 Then
'                SaldoV# = dCheckNull(AuswertungRec!Saldo)
'            End If
'      End If
'    End If
'    AuswertungRec.Bookmark = bmk
'End If
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
JahrSumme:
'Jahressumme: Summe d. Abrechnungsstelle, wenn da, sonst aus Rezeptspeicher
If Jahr Then
  RezAnz& = RezAnz& + dCheckNull(AuswertungRec!RezAnzahl)
End If
If (dCheckNull(AuswertungRec.Fields(FieldNr%(2))) = 0#) Then
   Werte#(0) = Werte#(0) + dCheckNull(AuswertungRec.Fields(FieldNr%(1)))
Else
    Werte#(0) = Werte#(0) + AuswertungRec.Fields(FieldNr%(2))
End If
Return
'-----------------------------------------------------------------------------------------------------------------------
SaldoVorMonRechnen:
With AuswertungRec
    Do
      AbrechVorMon = False
      For i% = 1 To 4
          'fieldnr(1) = rez_gesamt, 3 = rez_gesamtFAM, 5 = rez_ImpFähig, 7 = rez_ImpIst
          If (dCheckNull(AuswertungRec.Fields(FieldNr%(2 * i%))) = 0#) Then
              w#(i%) = w#(i%) + dCheckNull(AuswertungRec.Fields(FieldNr%(2 * i% - 1)))
          Else
              w#(i%) = w#(i%) + AuswertungRec.Fields(FieldNr%(2 * i%))
              If i% = 1 Then AbrechVorMon = True
          End If
      Next i%
      
      If dCheckNull(AuswertungRec!GutHaben) = 0# And Not AbrechVorMon Then
          SaldoV# = (w#(W_IMPIST) - w#(W_FAM) * (ImportQuote#(w#(W_FAM), w#(W_IMPFÄHIG), Left(AuswertungRec!Monat, 4)) / 100#)) / 10#
      End If
      AuswertungRec.MovePrevious
      If .BOF Then Exit Do
    Loop Until AbrechVorMon Or SaldoV# <> 0 Or AuswertungRec!Kkasse <> kNummer$
End With
Return

Call DefErrPop
End Function

Sub KKLoeschen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("KKLoeschen")
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
Dim SQLStr$
Dim Monat As String
Dim bmk As Variant
Dim RezGesamt#, RezGesamtFAM#, RezImpFähig#, RezImpIst#, RezAnzahl#, RezImpErspart#


SQLStr$ = "UPDATE Rezepte SET Rezepte.Kkasse = " + Chr$(34) + "Unbekannt" + Chr$(34) + " WHERE Rezepte.Kkasse = " + Chr$(34) + KKSatz + Chr$(34)
RezSpeicherDB.Execute (SQLStr)

With AuswertungRec
'Zahlen übertragen auf "Unbekannt"
    .index = "Unique"
    .Seek ">=", KKSatz, "0000"
    If Not .NoMatch And Not .EOF Then
        Do While AuswertungRec!Kkasse = KKSatz
            Monat = AuswertungRec!Monat
            RezGesamt# = AuswertungRec!Rez_Gesamt
            RezGesamtFAM# = AuswertungRec!Rez_GesamtFAM
            RezImpFähig# = AuswertungRec!Rez_ImpFähig
            RezImpIst# = AuswertungRec!Rez_ImpIst
            RezAnzahl = AuswertungRec!RezAnzahl
            RezImpErspart = AuswertungRec!ImpErspart
            
            bmk = .Bookmark
            .Seek "=", "Unbekannt", Monat
            If .NoMatch Then
                .AddNew
                AuswertungRec!Kkasse = "Unbekannt"
                AuswertungRec!Monat = Monat
                AuswertungRec!Rez_Gesamt = 0
                AuswertungRec!Rez_GesamtFAM = 0
                AuswertungRec!Rez_ImpFähig = 0
                AuswertungRec!Rez_ImpIst = 0
            Else
                .Edit
            End If
            
            AuswertungRec!Rez_Gesamt = dCheckNull(AuswertungRec!Rez_Gesamt) + RezGesamt#
            AuswertungRec!Rez_GesamtFAM = dCheckNull(AuswertungRec!Rez_GesamtFAM) + RezGesamtFAM#
            AuswertungRec!Rez_ImpFähig = dCheckNull(AuswertungRec!Rez_ImpFähig) + RezImpFähig#
            AuswertungRec!Rez_ImpIst = dCheckNull(AuswertungRec!Rez_ImpIst) + RezImpIst#
            AuswertungRec!RezAnzahl = dCheckNull(AuswertungRec!RezAnzahl) + RezAnzahl
            AuswertungRec!ImpErspart = dCheckNull(AuswertungRec!ImpErspart) + RezImpErspart
            AuswertungRec!Saldo = 0#
            .Update
            .Bookmark = bmk
            If AuswertungRec!Kkasse = KKSatz Then           'zur Sicherheit nochmal abfragen
                .Delete
            End If
            .Seek ">=", KKSatz, "0000"
            If .NoMatch Or .EOF Then Exit Do
        Loop
    End If
End With

If Not KasseRec.NoMatch Then
    If KasseRec!Nummer = KKSatz Then
        KasseRec.Delete
    End If
End If
'FabsErrf% = Kkasse.IndexSearch(0, KKSatz, FabsRecno&)
'If FabsErrf% = 0 Then
'    FabsErrf% = Kkasse.IndexDelete(0, FabsRecno&, KKSatz, FabsRecno&)
'End If
Call DefErrPop
End Sub

Function TesteGH%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TesteGH%")
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
'Dim i%, ind%, ind2%, ret%, Anzahl%, ArtikelStatus%, CrLf%, einzug%
'Dim StatusCode&, BranchNumber&
'Dim buf$, h$, h2$, AnzeigeStr$, pzn$, ResultTeil$(6), sMsgBox$
'Dim HttpReq As New MSXML2.XMLHTTP40
'Dim tim1
'Dim Pzns, Mengen, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12
'Dim KundenNr$, Passwort$, ApoIk$, DefektGrund$
'Dim s, s2
'
'KundenNr = "12345"
'ApoIk = "301234561"
'Passwort = "passwort"
'
'buf$ = "<?xml version='1.0' encoding='utf-8'?>"
'buf$ = buf$ + vbCrLf + "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:SOAP-ENC='http://schemas.xmlsoap.org/soap/encoding/' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema'>"
'buf$ = buf$ + vbCrLf + "  <soapenv:Body>"
'
''buf$ = buf$ + vbCrLf + "    <m:ladeRzVersion xmlns:m='http://fiverx.de/spec/abrechnungsservice/types'>"
''buf$ = buf$ + vbCrLf + "      <rzeParamLadeVersion>"
'buf$ = buf$ + vbCrLf + "    <m:ladeRzDienste xmlns:m='http://fiverx.de/spec/abrechnungsservice/types'>"
'buf$ = buf$ + vbCrLf + "      <rzeParamDienste>"
'
'buf$ = buf$ + vbCrLf + "<![CDATA["
''buf$ = buf$ + vbCrLf + "  <?xml version='1.0' encoding='ISO-8859-15'?>"
''buf$ = buf$ + vbCrLf + "        <rzeParamLadeVersion xmlns='http://fiverx.de/spec/abrechnungsservice'>"
'buf$ = buf$ + vbCrLf + "        <rzeParamDienste xmlns='http://fiverx.de/spec/abrechnungsservice'>"
'buf$ = buf$ + vbCrLf + "          <sendHeader>"
'buf$ = buf$ + vbCrLf + "            <rzKdNr>" + KundenNr + "</rzKdNr>"
'buf$ = buf$ + vbCrLf + "            <avsSw>"
'buf$ = buf$ + vbCrLf + "              <hrst>OPTIPHARM Software GmbH</hrst>"
'buf$ = buf$ + vbCrLf + "              <nm>Testapplication</nm>"
'buf$ = buf$ + vbCrLf + "              <vs>1.0</vs>"
'buf$ = buf$ + vbCrLf + "            </avsSw>"
'buf$ = buf$ + vbCrLf + "            <apoIk>" + ApoIk + "</apoIk>"
'buf$ = buf$ + vbCrLf + "            <test>false</test>"
'buf$ = buf$ + vbCrLf + "            <pw>" + Passwort + "</pw>"
'buf$ = buf$ + vbCrLf + "          </sendHeader>"
''buf$ = buf$ + vbCrLf + "        </rzeParamLadeVersion>"
'buf$ = buf$ + vbCrLf + "        </rzeParamDienste>"
'buf$ = buf$ + vbCrLf + "]]>"
'
''buf$ = buf$ + vbCrLf + "      </rzeParamLadeVersion>"
''buf$ = buf$ + vbCrLf + "    </m:ladeRzVersion>"
'buf$ = buf$ + vbCrLf + "      </rzeParamDienste>"
'
'buf$ = buf$ + vbCrLf + "      <rzeParamVersion>"
'buf$ = buf$ + vbCrLf + "<![CDATA["
'buf$ = buf$ + vbCrLf + "        <rzeParamVersion xmlns='http://fiverx.de/spec/abrechnungsservice'>"
'buf$ = buf$ + vbCrLf + "          <versionNr>01.06</versionNr>"
'buf$ = buf$ + vbCrLf + "        </rzeParamVersion>"
'buf$ = buf$ + vbCrLf + "]]>"
'buf$ = buf$ + vbCrLf + "      </rzeParamVersion>"
'
'buf$ = buf$ + vbCrLf + "    </m:ladeRzDienste>"
'
'buf$ = buf$ + vbCrLf + "  </soapenv:Body>"
'buf$ = buf$ + vbCrLf + "</soapenv:Envelope>"
'
'
'
'buf$ = "<?xml version='1.0' encoding='utf-8'?>"
'buf$ = buf$ + vbCrLf + "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:SOAP-ENC='http://schemas.xmlsoap.org/soap/encoding/' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema'>"
'buf$ = buf$ + vbCrLf + "  <soapenv:Body>"
'
''buf$ = buf$ + vbCrLf + "    <m:ladeRzVersion xmlns:m='http://fiverx.de/spec/abrechnungsservice/types'>"
''buf$ = buf$ + vbCrLf + "      <rzeParamLadeVersion>"
'buf$ = buf$ + vbCrLf + "    <m:sendeRezepte xmlns:m='http://fiverx.de/spec/abrechnungsservice/types'>"
'buf$ = buf$ + vbCrLf + "      <rzeLeistung>"
'
'buf$ = buf$ + vbCrLf + "<![CDATA["
''buf$ = buf$ + vbCrLf + "  <?xml version='1.0' encoding='ISO-8859-15'?>"
''buf$ = buf$ + vbCrLf + "        <rzeParamLadeVersion xmlns='http://fiverx.de/spec/abrechnungsservice'>"
'buf$ = buf$ + vbCrLf + "        <rzeLeistung xmlns='http://fiverx.de/spec/abrechnungsservice'>"
'buf$ = buf$ + vbCrLf + "          <rzLeistungHeader>"
'buf$ = buf$ + vbCrLf + "            <sendHeader>"
'buf$ = buf$ + vbCrLf + "              <rzKdNr>" + KundenNr + "</rzKdNr>"
'buf$ = buf$ + vbCrLf + "              <avsSw>"
'buf$ = buf$ + vbCrLf + "                <hrst>OPTIPHARM Software GmbH</hrst>"
'buf$ = buf$ + vbCrLf + "                <nm>Testapplication</nm>"
'buf$ = buf$ + vbCrLf + "                <vs>1.0</vs>"
'buf$ = buf$ + vbCrLf + "              </avsSw>"
'buf$ = buf$ + vbCrLf + "              <apoIk>" + ApoIk + "</apoIk>"
'buf$ = buf$ + vbCrLf + "              <test>false</test>"
'buf$ = buf$ + vbCrLf + "              <pw>" + Passwort + "</pw>"
'buf$ = buf$ + vbCrLf + "            </sendHeader>"
'buf$ = buf$ + vbCrLf + "            <sndId>17</sndId>"
'buf$ = buf$ + vbCrLf + "          </rzLeistungHeader>"
'
'buf$ = buf$ + vbCrLf + "          <rzLeistungInhalt>"
'buf$ = buf$ + vbCrLf + "            <eLeistungHeader>"
'buf$ = buf$ + vbCrLf + "              <avsId>16</avsId>"
'buf$ = buf$ + vbCrLf + "            </eLeistungHeader>"
'buf$ = buf$ + vbCrLf + "            <eLeistungBody>"
'buf$ = buf$ + vbCrLf + "              <eMuster16>"
'buf$ = buf$ + vbCrLf + "                <rezeptTyp>STANDARDREZEPT</rezeptTyp>"
'buf$ = buf$ + vbCrLf + "                <muster16Id>730000006</muster16Id>"
'buf$ = buf$ + vbCrLf + "                <kArt>19</kArt>"
'buf$ = buf$ + vbCrLf + "                <apoIk>" + ApoIk + "</apoIk>"
'buf$ = buf$ + vbCrLf + "                <rTyp>GKV</rTyp>"
'buf$ = buf$ + vbCrLf + "                <gesBrutto>295.05</gesBrutto>"
'buf$ = buf$ + vbCrLf + "                <zuzahlung>10.00</zuzahlung>"
'buf$ = buf$ + vbCrLf + "                <artikel>"
'buf$ = buf$ + vbCrLf + "                  <pzn>0429364</pzn>"
'buf$ = buf$ + vbCrLf + "                  <posNr>1</posNr>"
'buf$ = buf$ + vbCrLf + "                  <faktor>1</faktor>"
'buf$ = buf$ + vbCrLf + "                  <taxe>0.01</taxe>"
'buf$ = buf$ + vbCrLf + "                  <autidem>0</autidem>"
'buf$ = buf$ + vbCrLf + "                </artikel>"
'buf$ = buf$ + vbCrLf + "                <abDatum>2009-05-17</abDatum>"
'buf$ = buf$ + vbCrLf + "                <noctu>0</noctu>"
'buf$ = buf$ + vbCrLf + "                <bvg>0</bvg>"
'buf$ = buf$ + vbCrLf + "                <hilf>2</hilf>"
'buf$ = buf$ + vbCrLf + "                <impf>0</impf>"
'buf$ = buf$ + vbCrLf + "                <bgrPfl>0</bgrPfl>"
'buf$ = buf$ + vbCrLf + "                <gebFrei>0</gebFrei>"
'buf$ = buf$ + vbCrLf + "                <unfall>0</unfall>"
'buf$ = buf$ + vbCrLf + "                <aUnfall>0</aUnfall>"
'buf$ = buf$ + vbCrLf + "                <sonstige>0</sonstige>"
'buf$ = buf$ + vbCrLf + "              </eMuster16>"
'buf$ = buf$ + vbCrLf + "            </eLeistungBody>"
'buf$ = buf$ + vbCrLf + "          </rzLeistungInhalt>"
'
''buf$ = buf$ + vbCrLf + "        </rzeParamLadeVersion>"
'buf$ = buf$ + vbCrLf + "        </rzeLeistung>"
'buf$ = buf$ + vbCrLf + "]]>"
'
''buf$ = buf$ + vbCrLf + "      </rzeParamLadeVersion>"
''buf$ = buf$ + vbCrLf + "    </m:ladeRzVersion>"
'buf$ = buf$ + vbCrLf + "      </rzeLeistung>"
'
'buf$ = buf$ + vbCrLf + "      <rzeParamVersion>"
'buf$ = buf$ + vbCrLf + "<![CDATA["
'buf$ = buf$ + vbCrLf + "        <rzeParamVersion xmlns='http://fiverx.de/spec/abrechnungsservice'>"
'buf$ = buf$ + vbCrLf + "          <versionNr>01.06</versionNr>"
'buf$ = buf$ + vbCrLf + "        </rzeParamVersion>"
'buf$ = buf$ + vbCrLf + "]]>"
'buf$ = buf$ + vbCrLf + "      </rzeParamVersion>"
'
'buf$ = buf$ + vbCrLf + "    </m:sendeRezepte>"
'
'buf$ = buf$ + vbCrLf + "  </soapenv:Body>"
'buf$ = buf$ + vbCrLf + "</soapenv:Envelope>"
'
'
'
'
''txtActionWert(0).text = buf$
''DoEvents
'MsgBox (buf$)
'
'HttpReq.Open "POST", "http://ws.fiverx.de/axis2/services/FiverxLinkService", False
''HttpReq.Open "POST", "https://fiverx.arz-darmstadt.de/axis2/services/FiverxLinkService", False
'
'If Not (Err) Then
'    'Set a standard SOAP/ XML header for the content-type
'    HttpReq.setRequestHeader "Content-Type", "text/xml"
'End If
'
'If Not (Err) Then
'    'Set a header for the method to be called
''    HttpReq.setRequestHeader "SOAPMethodName", "http://194.59.150.80/webservices/wws/PointGateway11?wsdl#checkStockWws"
''    HttpReq.setRequestHeader "SOAPAction", "http://194.59.150.80/webservices/wws/PointGateway11?wsdl#checkStockWws"
''    HttpReq.setRequestHeader "SOAPAction", "https://fiverx.arz-darmstadt.de/axis2/services/FiverxLinkService?wsdl"
''    HttpReq.setRequestHeader "SOAPAction", "ladeRzVersion"
'    HttpReq.setRequestHeader "SOAPAction", "ladeRzDienste"
''    HttpReq.setRequestHeader "SOAPAction", ""
'End If
'
'If Not (Err) Then
'    HttpReq.Send buf
'End If
'
'If Not (Err) Then
'    buf = HttpReq.responseText
'
'    Do
'        ind = InStr(buf, "&lt;")
'        If (ind > 0) Then
'            buf = Left(buf, ind - 1) + "<" + Mid(buf, ind + 4)
'        Else
'            Exit Do
'        End If
'    Loop
'
'    ind = 1
'    Do
'        ind = InStr(ind + 2, buf, vbLf)
'        If (ind > 0) Then
'            If (Mid$(buf, ind - 1, 1) <> vbCr) Then
'                buf = Left(buf, ind - 1) + vbCr + Mid(buf, ind)
'            End If
'        Else
'            Exit Do
'        End If
'    Loop
'
'    ind = 1
'    Do
'        ind = InStr(ind + 3, buf, "<")
'        If (ind > 0) Then
'            CrLf = True
'            If (Mid$(buf, ind + 1, 1) = "/") Then
'                If (Mid$(buf, ind - 1, 1) <> ">") Then
'                    CrLf = 0
'                End If
'            ElseIf (Mid$(buf, ind - 2, 2) = vbCrLf) Then
'                CrLf = 0
'            End If
'            If (CrLf) Then
'                buf = Left(buf, ind - 1) + vbCrLf + Mid(buf, ind)
'            End If
'        Else
'            Exit Do
'        End If
'    Loop
'
'    einzug = 0
'    ind = 1
'    Do
'        ind = InStr(ind, buf, vbCrLf)
'        If (ind > 0) Then
'            If (Mid(buf, ind + 2, 1) = "<") Then
'                If (Mid(buf, ind + 3, 1) = "/") Then
'                    einzug = einzug - 2
'                Else
'                    For i = ind To 1 Step -1
'                        If (Mid$(buf, i, 1) = "<") Then
'                            If (Mid$(buf, i + 1, 1) = "/") Then
'                                einzug = einzug - 2
'                            ElseIf (Mid$(buf, i + 1, 1) = "?") Then
'                                einzug = einzug - 2
'                            End If
'                            Exit For
'                        End If
'                    Next i
'                    einzug = einzug + 2
'                End If
'                buf = Left(buf, ind + 1) + Space(einzug) + Mid(buf, ind + 2)
'            End If
'            ind = ind + 5
'        Else
'            Exit Do
'        End If
'    Loop
'    MsgBox (buf)
'End If
'
'
'
'If (sMsgBox$ <> "") Then
'    Call MsgBox(sMsgBox$, vbInformation Or vbOKOnly)
'End If
'
'TesteGH% = ret%


Call DefErrPop
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_MouseDown")
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
    
If (y <= wpara.NlCaptionY) Then
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

Call DefErrPop
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_MouseMove")
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
Dim c As Object

On Error Resume Next
For Each c In Controls
    If (TypeOf c Is nlCommand) Then
        If (c.MouseOver) Then
            c.MouseOver = 0
        End If
    End If
Next
On Error GoTo DefErr

Call DefErrPop
End Sub

Private Sub Form_Resize()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_Resize")
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

If (para.Newline) And (Me.Visible) Then
    CurrentX = wpara.NlFlexBackY
    CurrentY = (wpara.NlCaptionY - TextHeight(Caption)) / 2
    ForeColor = vbBlack
    Me.Print Caption
End If

Call DefErrPop
End Sub

Private Sub nlcmdEsc_Click()
Call cmdEsc_Click
End Sub

Private Sub nlcmdAbrech_click()
Call cmdAbrech_Click
End Sub

Private Sub nlcmdDruck_click()
Call cmdDruck_Click
End Sub

Private Sub nlcmdImpAlt_click(index As Integer)
Call cmdImpAlt_Click(index)
End Sub

Private Sub nlcmdkkneu_Click()
Call cmdKKNeu_Click
End Sub

Private Sub nlcmdKStamm_Click()
Call cmdKStamm_Click
End Sub

Private Sub nlcmdRezSpeicherLoeschen_Click()
Call cmdRezSpeicherLoeschen_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If (para.Newline) Then
    If (KeyAscii = 13) Then
'        Call nlcmdOk_Click
        Exit Sub
    ElseIf (KeyAscii = 27) And (nlcmdEsc.Visible) Then
        Call nlcmdEsc_Click
        Exit Sub
'    ElseIf (KeyAscii = Asc("<")) And (nlcmdImport(0).Visible) Then
''        Call nlcmdChange_Click(0)
'        nlcmdImport(0).Value = 1
'    ElseIf (KeyAscii = Asc(">")) And (nlcmdImport(1).Visible) Then
''        Call nlcmdChange_Click(1)
'        nlcmdImport(1).Value = 1
    End If
End If
    
If (TypeOf ActiveControl Is TextBox) Then
    If (iEditModus% <> 1) Then
        If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
        If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And (((iEditModus% <> 2) And (iEditModus% <> 4)) Or (Chr$(KeyAscii) <> ".")) Then
            Beep
            KeyAscii = 0
        End If
    End If
End If

End Sub

Private Sub picControlBox_Click(index As Integer)

If (index = 0) Then
    Me.WindowState = vbMinimized
ElseIf (index = 1) Then
    Me.WindowState = vbNormal
Else
    Unload Me
End If

End Sub

