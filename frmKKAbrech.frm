VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlbutton.ocx"
Begin VB.Form frmKKAbrech 
   AutoRedraw      =   -1  'True
   Caption         =   "Werte der Abrechnungsstelle"
   ClientHeight    =   5310
   ClientLeft      =   1575
   ClientTop       =   2580
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   5760
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   3240
      Picture         =   "frmKKAbrech.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   3480
      Picture         =   "frmKKAbrech.frx":00A9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   3720
      Picture         =   "frmKKAbrech.frx":015D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txtEin 
      Alignment       =   1  'Rechts
      BorderStyle     =   0  'Kein
      Height          =   375
      Index           =   6
      Left            =   3600
      TabIndex        =   6
      Text            =   "Text7"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox txtEin 
      Alignment       =   1  'Rechts
      BorderStyle     =   0  'Kein
      Height          =   375
      Index           =   5
      Left            =   3600
      TabIndex        =   8
      Text            =   "Text6"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtEin 
      Alignment       =   1  'Rechts
      BorderStyle     =   0  'Kein
      Height          =   375
      Index           =   4
      Left            =   3600
      TabIndex        =   7
      Text            =   "Text5"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtEin 
      Alignment       =   1  'Rechts
      BorderStyle     =   0  'Kein
      Height          =   375
      Index           =   3
      Left            =   3600
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtEin 
      Alignment       =   1  'Rechts
      BorderStyle     =   0  'Kein
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtEin 
      Alignment       =   1  'Rechts
      BorderStyle     =   0  'Kein
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtEin 
      Alignment       =   1  'Rechts
      BorderStyle     =   0  'Kein
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   390
      Left            =   1680
      TabIndex        =   10
      Top             =   3600
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   390
      Left            =   240
      TabIndex        =   9
      Top             =   3600
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxAbrech 
      Height          =   2700
      Left            =   240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
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
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   495
      Left            =   1560
      TabIndex        =   14
      Top             =   4200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   4200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin VB.Label lblKK 
      AutoSize        =   -1  'True
      Caption         =   "000000000 RVO/Primärkasse "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2580
   End
End
Attribute VB_Name = "frmKKAbrech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "KKABRECH.FRM"

Sub flxAbrechBefuellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxAbrechBefuellen")
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
Dim i%, mo%
Dim Quote#, ImpSoll#, diff#, ImpFähig#, ImpIst#, UmsGes#, UmsFAM#, Saldo#, GutHaben#, Erspart#
Dim s$
Dim rs As Recordset
  
  
KasseRec.index = "Unique"
KasseRec.Seek "=", KKSatz$
If Not KasseRec.NoMatch Then

  lblKK.Caption = KasseRec!Nummer + "  " + KasseRec!Name
  AuswertungRec.index = "Unique"
  AuswertungRec.Seek "=", KasseRec!Nummer, AbrechMo
  If Not AuswertungRec.NoMatch Then
      UmsGes# = dCheckNull(AuswertungRec!Rez_Gesamt)
      UmsFAM# = dCheckNull(AuswertungRec!Rez_GesamtFAM)
      ImpFähig# = dCheckNull(AuswertungRec!Rez_ImpFähig)
      ImpIst# = dCheckNull(AuswertungRec!Rez_ImpIst)
      Erspart# = dCheckNull(AuswertungRec!ImpErspart)
  Else
      UmsGes# = 0#
      UmsFAM# = 0#
      ImpFähig# = 0#
      ImpIst# = 0#
      Erspart# = 0#
  End If
  GutHaben# = dCheckNull(AuswertungRec!GutHaben)
  Saldo# = dCheckNull(AuswertungRec!Saldo)
  
  
  flxAbrech.TextMatrix(1, 4) = Format(UmsGes#, "0.00")
  flxAbrech.TextMatrix(2, 4) = Format(UmsFAM#, "0.00")
  flxAbrech.TextMatrix(3, 4) = Format(ImpFähig#, "0.00")
  flxAbrech.TextMatrix(4, 4) = Format(ImpIst#, "0.00")
  Quote# = ImportQuote#(UmsFAM#, ImpFähig#, Left$(AbrechMo, 4))
  flxAbrech.TextMatrix(5, 4) = Format(Quote#, "0.00")
  ImpSoll# = UmsFAM# * (Quote# / 100#) / 10#
  flxAbrech.TextMatrix(6, 4) = Format(ImpSoll#, "0.00")
  flxAbrech.TextMatrix(7, 4) = Format(Erspart#, "0.00")
  diff# = Erspart# - ImpSoll#
  'diff# = ImpIst# - ImpSoll#
  flxAbrech.TextMatrix(8, 4) = Format(diff#, "0.00")
  'diff# = diff# / 10#
  If Abs(diff#) > 5# Then
      flxAbrech.row = 9
      flxAbrech.col = 2
      If diff# < 0# Then
          flxAbrech.CellForeColor = vbRed
      Else
          flxAbrech.CellForeColor = txtEin(0).ForeColor
      End If
      flxAbrech.TextMatrix(9, 4) = Format(diff#, "0.00")
  Else
      flxAbrech.TextMatrix(9, 4) = ""
  End If
  
 
  mo% = Val(Mid(AbrechMo, 3, 2))
  Do While mo% Mod 3 <> 1
    mo% = mo% - 1
    If mo% < 1 Then Exit Do 'kann nicht vorkommen, aber sicher ist sicher
    AuswertungRec.Seek "=", KasseRec!Nummer, Left(AbrechMo, 2) + Format(mo%, "00")
    If Not AuswertungRec.NoMatch Then
        UmsGes# = UmsGes# + dCheckNull(AuswertungRec!Rez_Gesamt)
        UmsFAM# = UmsFAM# + dCheckNull(AuswertungRec!Rez_GesamtFAM)
        ImpFähig# = ImpFähig# + dCheckNull(AuswertungRec!Rez_ImpFähig)
        ImpIst# = ImpIst# + dCheckNull(AuswertungRec!Rez_ImpIst)
        Erspart# = Erspart# + dCheckNull(AuswertungRec!ImpErspart)
    End If
  Loop
  AuswertungRec.Seek "=", KasseRec!Nummer, AbrechMo
  
  flxAbrech.TextMatrix(1, 2) = Format(UmsGes#, "0.00")
  flxAbrech.TextMatrix(2, 2) = Format(UmsFAM#, "0.00")
  flxAbrech.TextMatrix(3, 2) = Format(ImpFähig#, "0.00")
  flxAbrech.TextMatrix(4, 2) = Format(ImpIst#, "0.00")
  Quote# = ImportQuote#(UmsFAM#, ImpFähig#, Left$(AbrechMo, 4))
  flxAbrech.TextMatrix(5, 2) = Format(Quote#, "0.00")
  ImpSoll# = UmsFAM# * (Quote# / 100#) / 10#
  flxAbrech.TextMatrix(6, 2) = Format(ImpSoll#, "0.00")
  flxAbrech.TextMatrix(7, 2) = Format(Erspart#, "0.00")
  diff# = Erspart# - ImpSoll#
  'diff# = ImpIst# - ImpSoll#
  flxAbrech.TextMatrix(8, 2) = Format(diff#, "0.00")
  'diff# = diff# / 10#
  If Abs(diff#) > 5# Then
      flxAbrech.row = 9
      flxAbrech.col = 2
      If diff# < 0# Then
          flxAbrech.CellForeColor = vbRed
      Else
          flxAbrech.CellForeColor = txtEin(0).ForeColor
      End If
      flxAbrech.TextMatrix(9, 2) = Format(diff#, "0.00")
  Else
      flxAbrech.TextMatrix(9, 2) = ""
  End If
  
  If Not AuswertungRec.NoMatch Then
      UmsGes# = dCheckNull(AuswertungRec!abr_gesamt)
      UmsFAM# = dCheckNull(AuswertungRec!abr_gesamtFAM)
      ImpFähig# = dCheckNull(AuswertungRec!abr_ImpFähig)
      ImpIst# = dCheckNull(AuswertungRec!abr_ImpIst)
      Erspart# = dCheckNull(AuswertungRec!abr_ImpErspart)
  Else
      UmsGes# = 0#
      UmsFAM# = 0#
      ImpFähig# = 0#
      ImpIst# = 0#
      Erspart# = 0#
  End If
  For i% = 0 To 6
      txtEin(i%).tag = ""
  Next i%
  txtEin(0).text = Format(UmsGes#, "0.00")
  txtEin(1).text = Format(UmsFAM#, "0.00")
  txtEin(2).text = Format(ImpFähig#, "0.00")
  txtEin(3).text = Format(ImpIst#, "0.00")
  txtEin(4).text = Format(GutHaben#, "0.00")
  txtEin(5).text = Format(Saldo#, "0.00")
  txtEin(6).text = Format(Erspart#, "0.00")
  If GutHaben# <> 0 Then txtEin(4).tag = "1"
  For i% = 0 To 3
      flxAbrech.TextMatrix(i% + 1, 1) = txtEin(i%).text
  Next i%
  flxAbrech.TextMatrix(9, 1) = txtEin(4).text
  flxAbrech.TextMatrix(10, 1) = txtEin(5).text
  Call DiffSpalte
  If txtEin(0).Visible Then txtEin(0).SetFocus
End If
Call DefErrPop
End Sub

Private Sub DiffSpalte()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DiffSpalte")
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
Dim Quote#, diff#, ImpSoll#, GutHaben#

Quote# = ImportQuote#(xVal(txtEin(1).text), xVal(txtEin(2).text), Left$(AbrechMo, 4))
flxAbrech.TextMatrix(5, 1) = Format(Quote#, "0.00")
ImpSoll# = xVal(txtEin(1).text) * (Quote# / 100#) / 10#
flxAbrech.TextMatrix(6, 1) = Format(ImpSoll#, "0.00")
flxAbrech.TextMatrix(7, 1) = Format(xVal(txtEin(6).text), "0.00")
'diff# = xVal(txtEin(3).text) - ImpSoll#
diff# = xVal(txtEin(6).text) - ImpSoll#
flxAbrech.TextMatrix(8, 1) = Format(diff#, "0.00")
GutHaben# = xVal(txtEin(4).text)

If txtEin(4).tag = "1" Then
    If GutHaben# <> 0# Then
        flxAbrech.TextMatrix(9, 1) = Format(GutHaben#, "0.00")
    ElseIf diff# <> 0 Then
        diff# = diff# / 10#
        If Abs(diff#) > 5# Then
            flxAbrech.TextMatrix(9, 1) = Format(diff#, "0.00")
            txtEin(4).text = Format(diff#, "0.00")
            txtEin(4).tag = ""
        Else
            flxAbrech.TextMatrix(9, 1) = ""
        End If
    Else
        flxAbrech.TextMatrix(9, 1) = ""
    End If
ElseIf diff# <> 0 Then
'    diff# = diff# / 10#
    If Abs(diff#) > 5# Then
'        flxAbrech.row = 9
'        flxAbrech.col = 1
'        If diff# < 0# Then
'            flxAbrech.CellForeColor = vbRed
'        Else
'            flxAbrech.CellForeColor = txtEin(0).ForeColor
'        End If
        flxAbrech.TextMatrix(9, 1) = Format(diff#, "0.00")
        txtEin(4).text = Format(diff#, "0.00")
    Else
        flxAbrech.TextMatrix(9, 1) = ""
    End If
End If
With flxAbrech
    For i% = 1 To 4
        .TextMatrix(i%, 3) = Format(xVal(txtEin(i% - 1).text) - xVal(.TextMatrix(i%, 2)), "0.00")
    Next i%
    .TextMatrix(9, 3) = Format(xVal(txtEin(4).text) - xVal(.TextMatrix(9, 2)), "0.00")
    .TextMatrix(10, 3) = Format(xVal(txtEin(5).text) - xVal(.TextMatrix(10, 2)), "0.00")
    
    For i% = 1 To .Rows - 1
        .TextMatrix(i%, 3) = Format(xVal(.TextMatrix(i%, 1)) - xVal(.TextMatrix(i%, 2)), "0.00")
    Next i%
End With

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


Private Sub cmdOk_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdOK_Click")
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
Dim Quote#, ImpSoll#

KasseRec.index = "Unique"
KasseRec.Seek "=", KKSatz$

AuswertungRec.index = "Unique"
AuswertungRec.Seek "=", KKSatz$, AbrechMo
If AuswertungRec.NoMatch Then
  AuswertungRec.AddNew
  AuswertungRec!Kkasse = KKSatz$
  AuswertungRec!Monat = AbrechMo
Else
  AuswertungRec.Edit
End If
AuswertungRec!abr_gesamt = xVal(txtEin(0).text)
AuswertungRec!abr_gesamtFAM = xVal(txtEin(1).text)
AuswertungRec!abr_ImpFähig = xVal(txtEin(2).text)
AuswertungRec!abr_ImpIst = xVal(txtEin(3).text)
AuswertungRec!GutHaben = xVal(txtEin(4).text)
AuswertungRec!Saldo = xVal(txtEin(5).text)
AuswertungRec!abr_ImpErspart = xVal(txtEin(6).text)
If AuswertungRec!Saldo < 0# Then AuswertungRec!Saldo = 0#
AuswertungRec.Update

'nach Neuanlage kein aktueller DS
AuswertungRec.index = "Unique"
AuswertungRec.Seek "=", KKSatz$, AbrechMo


Quote# = ImportQuote#(AuswertungRec!abr_gesamtFAM, AuswertungRec!abr_ImpFähig, Left$(AuswertungRec!Monat, 4))
ImpSoll# = AuswertungRec!abr_gesamtFAM * (Quote# / 100#)
If AuswertungRec!abr_ImpIst < ImpSoll# Then
    If Not KasseRec.NoMatch Then
        If Not KasseRec!Anzeige Then
            KasseRec.Edit
            KasseRec!Anzeige = True
            KasseRec.Update
'            FabsErrf% = Kkasse.IndexSearch(0, KKSatz$, FabsRecno&)
'            If FabsErrf% = 0 Then
'                Kkasse.GetRecord (FabsRecno& + 1)
'                Kkasse.Anzeige = 1
'                Kkasse.PutRecord (FabsRecno& + 1)
'            End If
        End If
    End If
End If
   
frmKKMatch.Show vbModal
If FormErg% Then
  If Val(KKSatz) > 0 Then Call flxAbrechBefuellen
Else
  Unload Me
End If
Call DefErrPop
End Sub

Private Sub flxAbrech_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxAbrech_GotFocus")
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
txtEin(0).SetFocus
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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%, lief%, spalte%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%, FeldInd%
Dim iAdd%, iAdd2%
Dim h$, h2$, FormStr$, AuswertungsMonat$
Dim Quote#, ImpSoll#, diff#, ImpFähig#, ImpIst#, UmsGes#, UmsFAM#
Dim c As Control

Me.KeyPreview = True

Call wpara.InitFont(Me)

lblKK.Top = wpara.TitelY
lblKK.Left = wpara.LinksX

With flxAbrech
    .Cols = 5
    .Rows = 11
    .FixedRows = 1
    .FixedCols = 1
    .FormatString = "|>Abrechnungsstelle|>Rezeptspeicher Quartal|>Differenz|>Rezeptspeicher " + Mid$(AbrechMo, 3, 2) + "/" + Left(AbrechMo, 2)
    .ColWidth(0) = Me.TextWidth(" Abzug/Guthaben      ")
    .ColWidth(1) = Me.TextWidth(" Abrechnungsstelle      ")
    .ColWidth(2) = Me.TextWidth(" Rezeptspeicher Quartal    ")
    .ColWidth(3) = Me.TextWidth("  Differenz       ")
    .ColWidth(4) = Me.TextWidth("  Rezeptspeicher 99/99       ")
    
    Breite1% = 0
    For i% = 0 To (.Cols - 1)
        Breite1% = Breite1% + .ColWidth(i%)
    Next i%
    .Width = Breite1% + 90
    .Height = .RowHeight(0) * .Rows + 90
    
    .Top = lblKK.Top + lblKK.Height + wpara.TitelY
    .Left = wpara.LinksX
    
    .TextMatrix(1, 0) = "Umsatz gesamt"
    .TextMatrix(2, 0) = "Umsatz FAM"
    .TextMatrix(3, 0) = "importfähig"
    .TextMatrix(4, 0) = "Importumsatz"
    .TextMatrix(5, 0) = "Quote %"
    .TextMatrix(6, 0) = "Quote EUR"
    .TextMatrix(7, 0) = "Ersparnis"
    .TextMatrix(8, 0) = "Differenz"
    .TextMatrix(9, 0) = "Abzug/Guthaben"
    .TextMatrix(10, 0) = "Saldo"
    '.RowHeight(8) = 0
    .col = 1
    
    For i% = 0 To 6
      txtEin(i%).Width = .ColWidth(1) - 10
      txtEin(i%).Height = .RowHeight(1) - 5
      txtEin(i%).Left = .Left + .ColWidth(0) + 45
      If i% >= 4 Then
        If i% = 6 Then
          txtEin(i%).Top = .Top + .RowHeight(0) + CLng(i%) * (.RowHeight(1)) + 50
        Else
          txtEin(i%).Top = .Top + .RowHeight(0) + CLng(i% + 4) * (.RowHeight(1)) + 50
        End If
      Else
          txtEin(i%).Top = .Top + .RowHeight(0) + CLng(i%) * (.RowHeight(1)) + 50
      End If
    Next i%
    
    If Val(KKSatz) = 0 Then
      frmKKMatch.Show vbModal
      If FormErg% Then
        If Val(KKSatz) > 0 Then Call flxAbrechBefuellen
      Else
        Unload Me
        Call DefErrPop: Exit Sub
      End If
    Else
      Call flxAbrechBefuellen
    End If
End With


Me.Caption = Me.Caption + " - " + frmRezSpeicher!cmbDatum.List(frmRezSpeicher!cmbDatum.ListIndex)
If Val(Mid(AbrechMo, 3, 2)) Mod 3 = 0 Then Me.Caption = Me.Caption + " / " + CStr(Val(Mid(AbrechMo, 3, 2)) \ 4 + 1) + ". Quartal " + Left(AbrechMo, 2)

Font.Bold = False   ' True
Me.Width = flxAbrech.Left + flxAbrech.Width + 2 * wpara.LinksX

cmdEsc.Top = flxAbrech.Top + flxAbrech.Height + 150
cmdOk.Top = cmdEsc.Top
cmdEsc.Width = wpara.ButtonX
cmdOk.Width = wpara.ButtonX
cmdEsc.Height = wpara.ButtonY
cmdOk.Height = wpara.ButtonY
cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

If (para.Newline) Then
    iAdd = wpara.NlFlexBackY
    iAdd2 = wpara.NlCaptionY
    
    With flxAbrech
        .ScrollBars = flexScrollBarNone
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
''        .BackColor = vbWhite
'        .BackColorFixed = RGB(199, 176, 123)
''        .BackColorSel = RGB(232, 217, 172)
''        .ForeColorSel = vbBlack
        
        .Left = .Left + iAdd
        .Top = .Top + iAdd
    
        For i% = 0 To 6
            txtEin(i%).Width = .ColWidth(1) '- 10
            txtEin(i%).Height = .RowHeight(1) '- 5
            txtEin(i%).Left = .Left + .ColWidth(0) '+ 45
            If i% >= 4 Then
                If i% = 6 Then
                    txtEin(i%).Top = .Top + .RowHeight(0) + CLng(i%) * (.RowHeight(1)) '+ 50
                Else
                    txtEin(i%).Top = .Top + .RowHeight(0) + CLng(i% + 4) * (.RowHeight(1)) '+ 50
                End If
            Else
                txtEin(i%).Top = .Top + .RowHeight(0) + CLng(i%) * (.RowHeight(1)) '+ 50
            End If
        Next i%
    End With
    
    cmdOk.Top = cmdOk.Top + 2 * iAdd
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
'        .Left = Me.ScaleWidth - .Width - 150
        .Top = flxAbrech.Top + flxAbrech.Height + iAdd + 600
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .Default = cmdEsc.Default
        .Cancel = cmdEsc.Cancel
        .Visible = True
    End With
    cmdEsc.Visible = False

    With nlcmdOk
        .Init
'        .Left = Me.ScaleWidth - .Width - 150
        .Top = nlcmdEsc.Top
        .Caption = cmdOk.Caption
        .TabIndex = cmdOk.TabIndex
        .Enabled = cmdOk.Enabled
        .Default = cmdOk.Default
        .Cancel = cmdOk.Cancel
        .Visible = True
    End With
    cmdOk.Visible = False

    nlcmdOk.Left = (Me.Width - (nlcmdOk.Width * 2 + 300)) / 2
    nlcmdEsc.Left = nlcmdOk.Left + nlcmdEsc.Width + 300

    Me.Height = nlcmdEsc.Top + nlcmdEsc.Height + wpara.FrmCaptionHeight + 450

    Call wpara.NewLineWindow(Me, nlcmdEsc.Top)
    With flxAbrech
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
    nlcmdOk.Visible = False
    nlcmdEsc.Visible = False
End If
'''''''''

Me.Left = frmRezSpeicher.Left + (frmRezSpeicher.Width - Me.Width) / 2
If Me.Left < 0 Then Me.Left = 0
Me.Top = frmRezSpeicher.Top + (frmRezSpeicher.Height - Me.Height) / 2

Call DefErrPop
End Sub


Private Sub txtEin_Change(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtEin_Change(" + CStr(index) + ")")
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
If index < 4 Or index = 6 Then
  Call DiffSpalte
End If
Call DefErrPop
End Sub

Private Sub txtEin_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtEin_GotFocus(" + CStr(index) + ")")
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
Call SelectAll(txtEin(index))
Call DefErrPop


End Sub

Private Sub txtEin_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtEin_KeyDown(" + CStr(index) + ")")
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

If KeyCode = vbKeyDown Then
    If index < 5 Then
      If index = 3 Then
        txtEin(6).SetFocus
      Else
        txtEin(index + 1).SetFocus
      End If
    ElseIf index = 5 Then
        txtEin(0).SetFocus
    ElseIf index = 6 Then
      txtEin(4).SetFocus
    End If
    KeyCode = 0
ElseIf KeyCode = vbKeyUp Then
    If index > 0 Then
      If index = 4 Then
        txtEin(6).SetFocus
      ElseIf index = 6 Then
        txtEin(3).SetFocus
      Else
        txtEin(index - 1).SetFocus
      End If
    Else
        txtEin(5).SetFocus
    End If
    KeyCode = 0
End If
Call DefErrPop
End Sub


Private Sub txtEin_LostFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtEin_LostFocus")
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

If index < 4 Then
  txtEin(index).text = flxAbrech.TextMatrix(index + 1, 1)
Else
  If index = 6 Then
    txtEin(index).text = flxAbrech.TextMatrix(7, 1)
  Else
  txtEin(index).text = flxAbrech.TextMatrix(index + 4, 1)
  End If
End If
Call DefErrPop

End Sub


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

Private Sub nlcmdOk_Click()
Call cmdOk_Click
End Sub

Private Sub nlcmdEsc_Click()
Call cmdEsc_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If (para.Newline) Then
    If (KeyAscii = 13) Then
        Call nlcmdOk_Click
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
    
If KeyAscii = vbKeyReturn Then
    If TypeOf Me.ActiveControl Is TextBox Then
        If Me.ActiveControl.index < 4 Then
            If Me.ActiveControl.tag <> "1" Then
                If flxAbrech.TextMatrix(Me.ActiveControl.index + 1, 1) <> txtEin(Me.ActiveControl.index).text Then
                    Me.ActiveControl.tag = "1"
                End If
            End If
            flxAbrech.TextMatrix(Me.ActiveControl.index + 1, 1) = txtEin(Me.ActiveControl.index).text
        Else
            If Me.ActiveControl.tag <> "1" Then
                If flxAbrech.TextMatrix(Me.ActiveControl.index + 4, 1) <> txtEin(Me.ActiveControl.index).text Then
                    Me.ActiveControl.tag = "1"
                End If
            End If
            flxAbrech.TextMatrix(Me.ActiveControl.index + 4, 1) = txtEin(Me.ActiveControl.index).text
        End If
        If Me.ActiveControl.index < 5 Then
            txtEin(Me.ActiveControl.index + 1).SetFocus
        Else
            If (para.Newline) Then
                nlcmdOk.SetFocus
            Else
                cmdOk.SetFocus
            End If
        End If
    End If
    KeyAscii = 0
End If

'If (TypeOf ActiveControl Is TextBox) Then
'    If (iEditModus% <> 1) Then
'        If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
'        If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And (((iEditModus% <> 2) And (iEditModus% <> 4)) Or (Chr$(KeyAscii) <> ".")) Then
'            Beep
'            KeyAscii = 0
'        End If
'    End If
'End If

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






