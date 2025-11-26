VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmWuFrage 
   Caption         =   "Artikel-Frage"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4305
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   480
      TabIndex        =   1
      Top             =   3600
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2520
      TabIndex        =   2
      Top             =   3600
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxWuFrage 
      Height          =   2280
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4022
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483633
      BackColorBkg    =   -2147483633
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Label lblWuFrage 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmWuFrage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type LieferungenRec
    LifDat As String * 11
    AnzArtikel As Integer
    Wert As Double
    fertig As Byte
    Name As String * 12
    Sort As String * 14
    RetourKz As String * 1
End Type

Dim AnzLieferungen%
Dim AlleLieferungen() As LieferungenRec

Private Const DefErrModul = "WUFRAGE.FRM"

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
Call DefErrFnc("cmdOk_Click")
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

If (Me.Caption = "Artikel zu WÜ hinzufügen") Then
    With flxWuFrage
        Call UpdateEinzelZeile(.TextMatrix(.row, 7))
    End With
End If

With flxWuFrage
    WuFrageErg% = .row
    If (.TextMatrix(.row, 0) = "Nachlieferung") Then
        WuFrageErg% = 4
    End If
End With

Unload Me

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
Dim i%, spBreite%, ind%, iLief%, iRufzeit%
Dim h$, h2$

lblWuFrage.Caption = WuFrageTxt$
Call wpara.InitFont(Me)

With lblWuFrage
    .Top = wpara.TitelY
    .Left = wpara.LinksX
End With

With flxWuFrage
    If (WuFrageErg%) Then
        Me.Caption = "Artikel zu WÜ hinzufügen"
    
        .Rows = 2
        .FixedRows = 1
        .Rows = 1
        .FormatString = "|Lieferant|^Datum|^Uhrzeit|>Anz.Artikel|>Warenwert||||"
        
        Font.Bold = True
        .ColWidth(0) = TextWidth("X")
        .ColWidth(1) = TextWidth("XXXXXXXXXXXXXXX")
        .ColWidth(2) = TextWidth("99:99:9999")
        .ColWidth(3) = TextWidth("Uhrzeit ")
        .ColWidth(4) = TextWidth("Anz.Artikel ")
        .ColWidth(5) = TextWidth("Warenwert ")
        .ColWidth(6) = wpara.FrmScrollHeight + 2 * wpara.FrmBorderHeight
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        Font.Bold = False
        
        spBreite% = 0
        For i% = 0 To .Cols - 1
            If (.ColWidth(i%) > 0) Then
                .ColWidth(i%) = .ColWidth(i%) + TextWidth("X")
            End If
            spBreite% = spBreite% + .ColWidth(i%)
        Next i%
        .Left = lblWuFrage.Left
        .Width = spBreite% + 90
        .Height = .RowHeight(0) * 11 + 90
        
        Call SortWuHeader
    Else
        Me.Caption = "Artikel-Frage"
        .Cols = 1
        .Rows = 0
            
        Font.Bold = True
        .ColWidth(0) = TextWidth("Alternativ-Artikel (auf Dauer)     ")
        Font.Bold = False
        
        spBreite% = 0
        For i% = 0 To .Cols - 1
            If (.ColWidth(i%) > 0) Then
                .ColWidth(i%) = .ColWidth(i%) + TextWidth("X")
            End If
            spBreite% = spBreite% + .ColWidth(i%)
        Next i%
        .Width = spBreite% + 90
       
        .AddItem "Neuer Artikel"
        .AddItem "Alternativ-Artikel (einmalig)"
        .AddItem "Alternativ-Artikel (auf Dauer)"
        If (Left$(WuFrageTxt$, 11) = "Strichcode ") Then
            .AddItem "Strichcode zuordnen"
        End If
        If (ActProgram.IstInNachlieferung%) Then
            .AddItem "Nachlieferung"
        End If
        
        .Left = lblWuFrage.Left + (lblWuFrage.Width - .Width) / 2
        .Height = .Rows * .RowHeight(0) + 90
        .row = 0
    End If
        
    .Top = lblWuFrage.Top + lblWuFrage.Height + 300
    .SelectionMode = flexSelectionByRow
    .col = 0
    .ColSel = .Cols - 1
    
    spBreite% = lblWuFrage.Width
    If (.Width > spBreite%) Then spBreite% = .Width
    
End With
    

With cmdOk
    .Top = flxWuFrage.Top + flxWuFrage.Height + 300 * wpara.BildFaktor
    .Width = wpara.ButtonX%
    .Height = wpara.ButtonY%
End With
With cmdEsc
    .Top = cmdOk.Top
    .Width = cmdOk.Width
    .Height = cmdOk.Height
End With

Me.Width = lblWuFrage.Left + spBreite% + 2 * wpara.LinksX

cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY% + 90 + wpara.FrmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

WuFrageErg% = -1

Call DefErrPop
End Sub

Function EinlesenWuHeader%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinlesenWuHeader%")
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
Dim i%, j%, Max%, IstNeu%, lief%
Dim zWert#
Dim LifDat$, h$, h2$, lName$

AnzLieferungen% = 0

ww.GetRecord (1)
Max% = ww.erstmax

For i% = 1 To Max%
    ww.GetRecord (i% + 1)

    If (ww.status = 2) And (Asc(ww.WuBestDatum) <> 0) And (ww.lief = Lieferant%) And (ww.IstAltLast = 0) Then
        LifDat$ = Chr$(ww.lief) + ww.WuBestDatum + Format(CVI(ww.WuBestZeit), "0000")

        zWert# = ww.WuAEP * ww.WuRm
        
        IstNeu% = True
        For j% = 0 To (AnzLieferungen% - 1)
            If (AlleLieferungen(j%).LifDat = LifDat$) Then
                IstNeu% = False
                Exit For
            End If
        Next j%
        
        If (IstNeu%) Then
            ReDim Preserve AlleLieferungen(AnzLieferungen%)
            AlleLieferungen(AnzLieferungen%).LifDat = LifDat$
            AlleLieferungen(AnzLieferungen%).AnzArtikel = 1
            AlleLieferungen(AnzLieferungen%).Wert = zWert#
            
            AlleLieferungen(AnzLieferungen%).fertig = ww.IstAltLast
            
            AlleLieferungen(AnzLieferungen%).RetourKz = "R"
            If (ww.WuNeuLm >= 0) Then AlleLieferungen(AnzLieferungen%).RetourKz = "?"
            
            IstNeu% = True
            lief% = Asc(Left$(LifDat$, 1))
            lif.GetRecord (lief% + 1)
            h$ = Trim$(lif.kurz)
            If (h$ <> "") Then
                If (Asc(Left$(h$, 1)) < 32) Then
                    h$ = ""
                End If
            End If
            lName$ = h$ + " (" + Mid$(Str$(lief%), 2) + ")"
            AlleLieferungen(AnzLieferungen%).Name = lName$
            
            
            h$ = Format(AlleLieferungen(AnzLieferungen%).fertig, "0")
            h2 = Mid$(LifDat$, 2, 6)
            h$ = h$ + Right$(h2$, 2) + Mid$(h2$, 3, 2) + Left$(h2$, 2)
            h2$ = Mid$(LifDat$, 8, 4)   'Format(CVI(Right$(LifDat$, 2)), "0000")
            h$ = h$ + h2$
            h2$ = Format(lief%, "000")
            h$ = h$ + h2$
            AlleLieferungen(AnzLieferungen%).Sort = h$
            
            AnzLieferungen% = AnzLieferungen% + 1
            
        Else
            AlleLieferungen(j%).AnzArtikel = AlleLieferungen(j%).AnzArtikel + 1
            AlleLieferungen(j%).Wert = AlleLieferungen(j%).Wert + zWert#
'            If (ww.IstAltLast) Then AlleLieferungen(j%).fertig = ww.IstAltLast
            If (ww.WuNeuLm >= 0) Then AlleLieferungen(j%).RetourKz = "?"
        End If
    End If
Next i%

EinlesenWuHeader% = AnzLieferungen%

Call DefErrPop
End Function

Sub SortWuHeader()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SortWuHeader")
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
Dim i%, row%
Dim h$, zeit$

With flxWuFrage
    .Rows = 1
    
    For i% = 0 To (AnzLieferungen% - 1)
        .Rows = .Rows + 1
        row% = .Rows - 1
        
        h$ = " "
        If (AlleLieferungen(i%).fertig) Then
            h$ = AlleLieferungen(i%).RetourKz
        End If
        .TextMatrix(row%, 0) = h$
        .TextMatrix(row%, 1) = AlleLieferungen(i%).Name
        
        h$ = Mid$(AlleLieferungen(i%).LifDat, 2, 6)
        .TextMatrix(row%, 2) = Left$(h$, 2) + "." + Mid$(h$, 3, 2) + "." + Right$(h$, 2)
        
        zeit$ = Mid$(AlleLieferungen(i%).LifDat, 8, 4) ' Format(CVI(Right$(AlleLieferungen(i%).LifDat, 2)), "0000")
        .TextMatrix(row%, 3) = Left$(zeit$, 2) + ":" + Mid$(zeit$, 3)
        .TextMatrix(row%, 4) = Format(AlleLieferungen(i%).AnzArtikel, "0")
        .TextMatrix(row%, 5) = Format(AlleLieferungen(i%).Wert, "0.00")
        .TextMatrix(row%, 7) = AlleLieferungen(i%).LifDat
        .TextMatrix(row%, 8) = .TextMatrix(row%, 0)
        .TextMatrix(row%, 9) = AlleLieferungen(i%).Sort
    Next i%
    
    .row = 1
    .col = 9
    .RowSel = .Rows - 1
    .ColSel = 9
    .Sort = 5
    
    .TopRow = 1
    .row = .Rows - 1
    .col = 0
    .RowSel = .row
    .ColSel = .Cols - 1
End With

Call DefErrPop
End Sub

