VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRezHistorie 
   Caption         =   "Historie Rezeptspeicher für "
   ClientHeight    =   6960
   ClientLeft      =   510
   ClientTop       =   375
   ClientWidth     =   8565
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   8565
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2280
      TabIndex        =   7
      Top             =   6240
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   840
      TabIndex        =   6
      Top             =   6240
      Width           =   1200
   End
   Begin TabDlg.SSTab tabRezHistorie 
      Height          =   5325
      Left            =   240
      TabIndex        =   8
      Top             =   360
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   9393
      _Version        =   327681
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&1 - Monate"
      TabPicture(0)   =   "RezHistorie.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fmeRezHistorie(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&2 - Tage"
      TabPicture(1)   =   "RezHistorie.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fmeRezHistorie(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&3 - Rezepte"
      TabPicture(2)   =   "RezHistorie.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fmeRezHistorie(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&4 - Ansicht"
      TabPicture(3)   =   "RezHistorie.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "fmeRezHistorie(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame fmeRezHistorie 
         Height          =   4215
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   8895
         Begin VB.PictureBox picRezept 
            AutoRedraw      =   -1  'True
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1320
            ScaleHeight     =   435
            ScaleWidth      =   915
            TabIndex        =   10
            Top             =   1680
            Width           =   975
         End
         Begin MSFlexGridLib.MSFlexGrid flxInfo 
            Height          =   540
            Index           =   0
            Left            =   480
            TabIndex        =   14
            Top             =   2880
            Width           =   6720
            _ExtentX        =   11853
            _ExtentY        =   953
            _Version        =   65541
            Rows            =   0
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            BackColor       =   -2147483633
            BackColorBkg    =   -2147483633
            FocusRect       =   0
            HighLight       =   2
            GridLinesFixed  =   1
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin VB.Label lblAv 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   3720
            TabIndex        =   13
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lblAv 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   12
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label lblAv 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   3720
            TabIndex        =   11
            Top             =   960
            Width           =   1815
         End
      End
      Begin VB.Frame fmeRezHistorie 
         Height          =   4215
         Index           =   2
         Left            =   -75000
         TabIndex        =   4
         Top             =   840
         Width           =   8895
         Begin MSFlexGridLib.MSFlexGrid flxRezHistorie 
            Height          =   2700
            Index           =   2
            Left            =   840
            TabIndex        =   5
            Top             =   840
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   4763
            _Version        =   65541
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
      End
      Begin VB.Frame fmeRezHistorie 
         Height          =   4215
         Index           =   1
         Left            =   -74520
         TabIndex        =   2
         Top             =   480
         Width           =   8895
         Begin MSFlexGridLib.MSFlexGrid flxRezHistorie 
            Height          =   2700
            Index           =   1
            Left            =   840
            TabIndex        =   3
            Top             =   840
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   4763
            _Version        =   65541
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
      End
      Begin VB.Frame fmeRezHistorie 
         Height          =   4215
         Index           =   0
         Left            =   -74280
         TabIndex        =   0
         Top             =   480
         Width           =   8895
         Begin MSFlexGridLib.MSFlexGrid flxRezHistorie 
            Height          =   2700
            Index           =   0
            Left            =   720
            TabIndex        =   1
            Top             =   840
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   4763
            _Version        =   65541
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
      End
   End
End
Attribute VB_Name = "frmRezHistorie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const DefErrModul = "REZHISTORIE.FRM"

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
Call ActProgram.RezHistorieExit
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
Dim i%, VERBAND%
Dim h$

'If (ActiveControl.Name = cmdOk.Name) Then
'    Call AuslesenFlexTaetigkeiten
'
'    RezApoNr$ = Right$(Space$(7) + Trim(txtOptionen0(0).text), 7)
'
''    VmRabattFaktor# = Val(txtOptionen0(1).text)
'    VmRabattFaktor# = 100# / (100# - Val(txtOptionen0(1).text))
'
'
'    OrgBundesland% = Val(flxOptionen1(0).TextMatrix(flxOptionen1(0).row, 1))
'
'    VERBAND% = FileOpen("verbandm.dat", "RW", "B")
'
'    h$ = MKI(OrgBundesland%)
'    Seek #VERBAND%, 7
'    Put #VERBAND%, , h$
'
'    h$ = Format(VmRabattFaktor#, "0.0000")
'    For i% = 1 To Len(h$)
'        If (Mid$(h$, i%, 1) = ",") Then Mid$(h$, i%, 1) = "."
'    Next i%
'    Seek #VERBAND%, 9
'    Put #VERBAND%, , h$
'    Close #VERBAND%
'
'    OptionenNeu% = True
'    Unload Me
'ElseIf (ActiveControl.Name = flxOptionen1(0).Name) Then
'    If (ActiveControl.index = 1) Then
'        Call EditOptionenLstMulti
'    End If
'End If

Call DefErrPop
End Sub

Private Sub flxRezHistorie_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxRezHistorie_GotFocus")
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

With flxRezHistorie(Index)
    .col = 0
    .ColSel = .Cols - 1
    .HighLight = flexHighlightAlways
End With

Call DefErrPop
End Sub

Private Sub flxRezHistorie_LostFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxRezHistorie_LostFocus")
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

With flxRezHistorie(Index)
    .HighLight = flexHighlightNever
End With

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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%, FormVersatzY%
Dim Breite%, Breite2%, Hoehe1%, Hoehe2%
Dim h$, h2$, h3$, FormStr$


Call wpara.InitFont(Me)

FormVersatzY% = wpara.FrmCaptionHeight + wpara.FrmBorderHeight
Me.Width = frmRezSpeicher.Width
Me.Height = frmRezSpeicher.Height - FormVersatzY%
Me.Left = frmRezSpeicher.Left
Me.Top = frmRezSpeicher.Top + FormVersatzY%



'''''''''''''''''''''''''''''''''''
With tabRezHistorie
    .Left = wpara.LinksX
    .Top = wpara.TitelY
    .Width = ScaleWidth - 2 * .Left
    .Height = (Me.ScaleHeight - .Top - wpara.ButtonY - 300)
End With


For i% = 0 To 3
    With fmeRezHistorie(i%)
        .Left = wpara.LinksX
        .Top = 3 * wpara.TitelY
        .Width = tabRezHistorie.Width - 2 * .Left
        .Height = tabRezHistorie.Height - 2 * .Top
    End With
Next i%



For i% = 0 To 2
    tabRezHistorie.Tab = i%
    With flxRezHistorie(i%)
        .Cols = 8
        .Rows = 2
        .FixedRows = 1
        
        If (i% = 0) Then
            h$ = "Monat"
        ElseIf (i% = 1) Then
            h$ = "Tag"
        Else
            h$ = "Rezept"
        End If
        .FormatString = "|^" + h$ + "|>Gesamt|>Anzahl|>FAM|>ImpFähig|>ImpIst|>"
        .Rows = 1
        
        .Left = wpara.LinksX
        .Top = 2 * wpara.TitelY
        
        .Height = ((fmeRezHistorie(0).Height - .Top - 300) \ .RowHeight(0)) * .RowHeight(0) + 90
        .Width = fmeRezHistorie(0).Width - 2 * .Left
        
        .ColWidth(0) = 0
        .ColWidth(1) = 0 ' TextWidth(String(15, "9"))
        .ColWidth(2) = TextWidth(String(11, "9"))
        .ColWidth(3) = TextWidth(String(7, "9"))
        .ColWidth(4) = TextWidth(String(11, "9"))
        .ColWidth(5) = TextWidth(String(11, "9"))
        .ColWidth(6) = TextWidth(String(11, "9"))
        .ColWidth(7) = wpara.FrmScrollHeight
        
        Breite% = 0
        For j% = 0 To (.Cols - 1)
            Breite% = Breite% + .ColWidth(j%)
        Next j%
        Breite% = .Width - Breite% - 90
        If (Breite% > 0) Then .ColWidth(1) = Breite%
        
        .SelectionMode = flexSelectionByRow
        .col = 0
        .ColSel = .Cols - 1
    End With
Next i%

tabRezHistorie.Tab = 3
With picRezept
    .Width = flxRezHistorie(0).Width
    .Height = .Width * 10.3 / 15
    If (.Height > flxRezHistorie(0).Height) Then
        .Height = flxRezHistorie(0).Height
        .Width = .Height * 15 / 10.3
    End If
    .Left = wpara.LinksX
    .Top = flxRezHistorie(0).Top
    
    Call ActProgram.RezHistorieInit(Me)
'    Call ActProgram.PaintRezept
End With


Call RezHistorieBefuellen(0)

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

cmdOk.Top = tabRezHistorie.Top + tabRezHistorie.Height + 150
cmdEsc.Top = cmdOk.Top

'Me.Width = tabRezHistorie.Width + 2 * wpara.LinksX

cmdOk.Width = wpara.ButtonX
cmdOk.Height = wpara.ButtonY
cmdEsc.Width = wpara.ButtonX
cmdEsc.Height = wpara.ButtonY
cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

'Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

'Breite1% = frmAction.Left + (frmAction.Width - Me.Width) / 2
'If (Breite1% < 0) Then Breite1% = 0
'Me.Left = Breite1%
'Hoehe1% = frmAction.Top + (frmAction.Height - Me.Height) / 2
'If (Hoehe1% < 0) Then Hoehe1% = 0
'Me.Top = Hoehe1%

Caption = Caption + RezHistorieKassenNr$ + " - " + RezHistorieKassenName$

tabRezHistorie.Tab = 0
Call TabDisable
Call TabEnable(tabRezHistorie.Tab)

Call DefErrPop
End Sub

Private Sub tabRezHistorie_Click(PreviousTab As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("tabRezHistorie_Click")
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
If (tabRezHistorie.Visible = False) Then Call DefErrPop: Exit Sub

Call TabDisable
Call TabEnable(tabRezHistorie.Tab)

Call DefErrPop
End Sub

Private Sub flxRezHistorie_DblClick(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxRezHistorie_DblClick")
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
Call cmdOk_Click
Call DefErrPop
End Sub

Sub TabDisable()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TabDisable")
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

For i% = 0 To 3
    fmeRezHistorie(i%).Visible = False
Next i%

Call DefErrPop
End Sub

Sub TabEnable(hTab%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TabEnable")
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
Dim Unique$

fmeRezHistorie(hTab%).Visible = True
If (hTab% < 3) Then
    If (flxRezHistorie(hTab%).Visible) Then flxRezHistorie(hTab%).SetFocus
    Call RezHistorieBefuellen(hTab%)
Else
    Unique$ = flxRezHistorie(2).TextMatrix(flxRezHistorie(2).row, 0)
    RezepteRec.Index = "Unique"
    RezepteRec.Seek "=", Unique$
    Call ActProgram.HoleAusRezeptSpeicher
    Call ActProgram.PaintRezept
End If

Call DefErrPop
End Sub

Sub RezHistorieBefuellen(ind%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RezHistorieBefuellen")
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
Dim i%, anz%
Dim l&
Dim Gesamt#, Fam#, ImpFähig#, ImpIst#
Dim h$, SuchMonat$, SuchTag$

If (ind% = 0) Then
    flxRezHistorie(0).Rows = 1
    With AuswertungRec
        .Index = "Unique"
        .Seek ">=", RezHistorieKassenNr$
        If Not .NoMatch Then
            Do While Not .EOF
                If (AuswertungRec!Kkasse = RezHistorieKassenNr$) Then
                    h$ = AuswertungRec!Monat
                    h$ = h$ + vbTab + Format(CDate("01." + Mid(AuswertungRec!Monat, 3, 2) + ".20" + Left(AuswertungRec!Monat, 2)), "MM/YYYY")
                    h$ = h$ + vbTab + Format(dCheckNull(AuswertungRec!Rez_Gesamt), "0.00")
                    h$ = h$ + vbTab + Format(dCheckNull(AuswertungRec!RezAnzahl), "0")
                    h$ = h$ + vbTab + Format(dCheckNull(AuswertungRec!Rez_GesamtFAM), "0.00")
                    h$ = h$ + vbTab + Format(dCheckNull(AuswertungRec!Rez_ImpFähig), "0.00")
                    h$ = h$ + vbTab + Format(dCheckNull(AuswertungRec!Rez_ImpIst), "0.00")
                    h$ = h$ + vbTab
                    flxRezHistorie(0).AddItem h$
                Else
                    Exit Do
                End If
                .MoveNext
            Loop
        Else
            flxRezHistorie(0).AddItem " "
        End If
    End With
    With flxRezHistorie(0)
        .row = 1
        .col = 0
        .RowSel = .Rows - 1
        .ColSel = .col
        .Sort = 6
        .col = 0
        .ColSel = .Cols - 1
    End With

ElseIf (ind% = 1) Then
    flxRezHistorie(1).Rows = 1
    With RezepteRec
        .Index = "Kasse"
        SuchMonat$ = flxRezHistorie(0).TextMatrix(flxRezHistorie(0).row, 0)
        .Seek ">=", RezHistorieKassenNr$, SuchMonat$ + "01"
        If Not .NoMatch Then
            Do While Not .EOF
                If (RezepteRec!Kkasse = RezHistorieKassenNr$) And (Left$(RezepteRec!VerkDatum, 4) = SuchMonat$) Then
                    h$ = vbTab + RezepteRec!VerkDatum
                    h$ = h$ + vbTab + Format(dCheckNull(RezepteRec!RezSumme), "0.00")
                    h$ = h$ + vbTab + "1"
    '                    h$ = h$ + vbTab + Format(dCheckNull(RezepteRec!AnzArtikel), "0")
                    h$ = h$ + vbTab + Format(dCheckNull(RezepteRec!Fam), "0.00")
                    h$ = h$ + vbTab + Format(dCheckNull(RezepteRec!ImpFähig), "0.00")
                    h$ = h$ + vbTab + Format(dCheckNull(RezepteRec!ImpIst), "0.00")
                    h$ = h$ + vbTab
                    flxRezHistorie(1).AddItem h$
                Else
                    Exit Do
                End If
                .MoveNext
            Loop
        Else
            flxRezHistorie(1).AddItem " "
        End If
    End With
    With flxRezHistorie(1)
        .row = 1
        .col = 1
        .RowSel = .Rows - 1
        .ColSel = .col
        .Sort = 5
    
        h$ = ""
        l& = 1
        Do
            If (l& >= .Rows) Then Exit Do
            
            If (.TextMatrix(l&, 1) = h$) Then
                For i% = 2 To 6
                    .TextMatrix(l& - 1, i%) = Format(xVal(.TextMatrix(l& - 1, i%)) + xVal(.TextMatrix(l&, i%)), "0.00")
                Next i%
                
                .RemoveItem l&
            Else
                h$ = .TextMatrix(l&, 1)
                l& = l& + 1
            End If
        Loop
    End With

ElseIf (ind% = 2) Then
    flxRezHistorie(2).Rows = 1
    With RezepteRec
        .Index = "Kasse"
        SuchTag$ = flxRezHistorie(1).TextMatrix(flxRezHistorie(1).row, 1)
        .Seek ">=", RezHistorieKassenNr$, SuchTag$
        If Not .NoMatch Then
            Do While Not .EOF
                If (RezepteRec!Kkasse = RezHistorieKassenNr$) And (RezepteRec!VerkDatum = SuchTag$) Then
                    h$ = RezepteRec!Unique + vbTab + RezepteRec!VerkDatum
                    h$ = h$ + vbTab + Format(dCheckNull(RezepteRec!RezSumme), "0.00")
                    h$ = h$ + vbTab + Format(dCheckNull(RezepteRec!AnzArtikel), "0")
                    h$ = h$ + vbTab + Format(dCheckNull(RezepteRec!Fam), "0.00")
                    h$ = h$ + vbTab + Format(dCheckNull(RezepteRec!ImpFähig), "0.00")
                    h$ = h$ + vbTab + Format(dCheckNull(RezepteRec!ImpIst), "0.00")
                    h$ = h$ + vbTab
                    flxRezHistorie(2).AddItem h$
                Else
                    Exit Do
                End If
                .MoveNext
            Loop
        Else
            flxRezHistorie(2).AddItem " "
        End If
    End With
End If

Call DefErrPop
End Sub

