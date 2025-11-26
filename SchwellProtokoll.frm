VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSchwellProtokoll 
   Caption         =   "Form1"
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
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   360
      TabIndex        =   1
      Top             =   3000
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxSchwellProt 
      Height          =   2280
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4022
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLines       =   0
      ScrollBars      =   0
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmSchwellProtokoll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LiefInd%(10)
Dim MindDiff#(10)

Private Const DefErrModul = "SCHWELLPROTOKOLL.FRM"

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

Private Sub flxSchwellProt_KeyPress(KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxSchwellProt_KeyPress")
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
Dim row%, col%, iLief%
Dim h$, h2$

If (KeyAscii = 13) Then
    With flxSchwellProt
        row% = .row
        col% = .col
    End With
    
    iLief% = LiefInd%((col% - 1) \ 2)
    h$ = Format(iLief%, "000")
    
    If ((col% Mod 2) = 0) Then col% = col% - 1
    h2$ = flxSchwellProt.TextMatrix(0, col%) + ": "
    
    If (row% = 2) Or (row% = 3) Then
        If (row% = 2) Then
            h$ = "S " + h$
            h2$ = h2$ + "Umsatz/Sendung"
        Else
            h$ = "U " + h$
            h2$ = h2$ + "Prognose Umsatz"
        End If
        SchwellInfoSuch$ = h$
        SchwellInfoName$ = h2$
        frmSchwellProtInfo.Show 1
    ElseIf (row% >= 5) And (row% <= 8) Then
        If (row% = 5) Then
            h$ = "Z " + h$
            h2$ = h2$ + "Zugeordnete Artikel"
        ElseIf (row% = 6) Then
            h$ = "M " + h$
            h2$ = h2$ + "Auffüllen Mindest-Umsatz (Diff: " + Format(MindDiff#((col% - 1) \ 2), "### ##0.00") + ")"
        ElseIf (row% = 7) Then
            h$ = "P " + h$
            h2$ = h2$ + "Auffüllen Schwellwert-Sprung"
        ElseIf (row% = 8) Then
            h$ = "B " + h$
            h2$ = h2$ + "Günstigster Lieferant"
        End If
        SchwellInfoSuch$ = h$
        SchwellInfoName$ = h2$
        frmSchwellProtInfo.Show 1
    End If
End If


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
Dim i%, spBreite%, ind%, iLief%, iRufzeit%, row%, col%, iModus%, maxSp%, iToggle%
Dim DRUCKHANDLE%
Dim h$, h2$, FormStr$, SollStr$

Call wpara.InitFont(Me)

With flxSchwellProt
    .Cols = 2
    .Rows = 9
    .FixedRows = 1
    .FixedCols = 1
    
    .Top = wpara.TitelY
    .Left = wpara.LinksX
    .Height = .RowHeight(0) * .Rows + 90
    
    FormStr$ = ""
    For i% = 1 To 30
        FormStr$ = FormStr$ + "|^" + Mid$(Str$(i%), 2)
    Next i%
'    FormStr$ = FormStr$ + ";|Rufzeit|Sendungen bisher|Sendungen geplant|Sendungen %|Umsatz/Sendung|"
'    FormStr$ = FormStr$ + "Mindest-Umsatz|aliq.Mindest-Ums|Umsatz bisher|Umsatz m.Zuord|"
'    FormStr$ = FormStr$ + "Umsatz m.Mindest|Prognose Umsatz|Prognose Rabatt|Umsatz m.BestLief"
    
    FormStr$ = FormStr$ + ";|Rufzeit|Umsatz/Sendung|Prognose Umsatz|"
    FormStr$ = FormStr$ + "Umsatz bisher| + Zuordnungen|"
    FormStr$ = FormStr$ + " + aliquot. Mindest-Ums.| + Schwellw.Sprung| + günstigster Lief|"
    FormStr$ = FormStr$ + " = Umsatz inkl. Sendung"
    .FormatString = FormStr$
    .ColAlignment(0) = flexAlignLeftCenter
    .SelectionMode = flexSelectionFree
    
    .FillStyle = flexFillRepeat
    col% = 3
    Do
        If (col% >= .Cols) Then Exit Do
        .row = 1
        .col = col%
        .RowSel = .Rows - 1
        .ColSel = col% + 1
        .CellBackColor = vbButtonFace
        col% = col% + 4
    Loop
    .FillStyle = flexFillSingle
    
    
    DRUCKHANDLE% = FileOpen("winw\" + GesendetDatei$, "I")
    col% = -1
    SollStr$ = "XXX"
    Do While Not EOF(DRUCKHANDLE%)
        Line Input #DRUCKHANDLE%, h$
        If (Left$(h$, 1) = "*") Then
            iModus% = True
            col% = col% + 2
            iToggle% = 0
        
            h2$ = ""
            iLief% = Val(Mid$(h$, 2))
            LiefInd%((col% - 1) / 2) = iLief%
            If (iLief% > 0) And (iLief% <= lif.AnzRec) Then
                lif.GetRecord (iLief% + 1)
                h2$ = RTrim$(lif.kurz)
                If (h2$ = String$(Len(h2$), 0)) Then h2$ = ""
                If (h2$ = "") Then
                    h2$ = "(" + Str$(iLief%) + ")"
                End If
            End If
            .TextMatrix(0, col%) = h2$
            .TextMatrix(0, col% + 1) = "rabattf."
            row% = 1
            SollStr$ = "A " + Format(iLief%, "000")
        ElseIf (Left$(h$, 2) = "Q ") Then
            h$ = Mid$(h$, 6)
            MindDiff#((col% - 1) / 2) = Val(h$)
        ElseIf (Left$(h$, 5) = SollStr$) Then
            h$ = Mid$(h$, 6)
            .TextMatrix(row%, col% + iToggle%) = Trim(h$)
            If (iToggle% = 0) Then
                iToggle% = 1
            Else
                row% = row% + 1
                iToggle% = 0
            End If
        End If
    Loop
    Close #DRUCKHANDLE%


    .Cols = col% + 2
    .ColWidth(0) = TextWidth("= Umsatz inkl. Sendung    ")
    For i% = 1 To .Cols - 1
        .ColWidth(i%) = TextWidth("9 999 999.99 ")
        .ColAlignment(i%) = flexAlignRightCenter
    Next i%

    maxSp% = (Screen.Width - (2 * wpara.LinksX) - 900 - .ColWidth(0)) \ .ColWidth(1) + 1
    If (.Cols <= maxSp%) Then
        maxSp% = .Cols
    Else
'        .Height = .Height + wpara.FrmScrollHeight
    End If
    
    spBreite% = 0
    For i% = 0 To maxSp% - 1
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
    
    .row = 1
    .col = 1
End With


Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

Me.Width = flxSchwellProt.Left + flxSchwellProt.Width + 2 * wpara.LinksX

With cmdEsc
    .Top = flxSchwellProt.Top + flxSchwellProt.Height + 150 * wpara.BildFaktor
    .Width = wpara.ButtonX%
    .Height = wpara.ButtonY%
    .Left = (ScaleWidth - .Width) / 2
End With

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY% + 90 + wpara.FrmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2


h$ = "Schwellwert-Automatik: "
iLief% = Val(Left$(GesendetDatei$, 3))
If (iLief% > 0) And (iLief% <= lif.AnzRec) Then
    lif.GetRecord (iLief% + 1)
    h2$ = RTrim$(lif.Name(0))
    
    iRufzeit% = Val(Mid$(GesendetDatei$, 4, 4))
    h2$ = h2$ + "  (" + Format(iRufzeit% \ 100, "00") + ":" + Format(iRufzeit% Mod 100, "00") + ")"
    
    If (InStr(GesendetDatei$, "m.") > 0) Then h2$ = h2$ + "  manuell"
    h$ = h$ + h2$
End If
Caption = h$

Call DefErrPop
End Sub

