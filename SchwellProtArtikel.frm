VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSchwellProtArtikel 
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
      Default         =   -1  'True
      Height          =   450
      Left            =   360
      TabIndex        =   1
      Top             =   3000
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxSchwellProtArtikel 
      Height          =   2280
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4022
      _Version        =   65541
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
End
Attribute VB_Name = "frmSchwellProtArtikel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "SCHWELLPROTARTIKEL.FRM"

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
Dim DRUCKHANDLE%
Dim zWertSumme#
Dim h$, h2$

Call wpara.InitFont(Me)

With flxSchwellProtArtikel
    .Cols = 9
    .Rows = 2
    .FixedRows = 1
    .FormatString = "<PZN|<Name|>Menge|^Meh|>BM|>NR|>Zeilenwert|Rab|"
        
    Font.Bold = True
    .ColWidth(0) = TextWidth("9999999")
    .ColWidth(1) = TextWidth("Xxxxxx Xxxxxx Xxxxxx Xxxxxx")
    .ColWidth(2) = TextWidth("XXXXXX")
    .ColWidth(3) = TextWidth("XXX")
    .ColWidth(4) = TextWidth("9999")
    .ColWidth(5) = TextWidth("9999")
    .ColWidth(6) = TextWidth("999999.99")
    .ColWidth(7) = TextWidth("Rab")
    .ColWidth(8) = wpara.FrmScrollHeight
    Font.Bold = False
    
    spBreite% = 0
    For i% = 0 To .Cols - 1
        If (.ColWidth(i%) > 0) Then
            .ColWidth(i%) = .ColWidth(i%) + TextWidth("X")
        End If
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
    
    .Rows = 1
    zWertSumme# = 0#
    
    DRUCKHANDLE% = FileOpen("winw\" + GesendetDatei$, "I")
    Do While Not EOF(DRUCKHANDLE%)
        Line Input #DRUCKHANDLE%, h$
        If (Left$(h$, 1) = "*") Then
            Exit Do
        ElseIf (Left$(h$, 5) = SchwellArtikelSuch$) Then
            .Rows = .Rows + 1
            .row = .Rows - 1
            If (Mid$(h$, 7, 1) = "1") Then
                .TextMatrix(.row, 7) = " " + Chr$(214)
            Else
                .TextMatrix(.row, 7) = " "
            End If
            h$ = Mid$(h$, 8)
            For i% = 0 To 5
                ind% = InStr(h$, vbTab)
                h2$ = Left$(h$, ind% - 1)
                h$ = Mid$(h$, ind% + 1)
                .TextMatrix(.row, i% + 1) = h2$
            Next i%
            .TextMatrix(.row, 0) = Left$(h$, Len(h$) - 1)
            zWertSumme# = zWertSumme# + CDbl(.TextMatrix(.row, 6))
        End If
    Loop
    Close #DRUCKHANDLE%
    
    If (.Rows > 1) Then
        .row = 1
        .col = 1
        .RowSel = .Rows - 1 ' AnzBestellArtikel%
        .ColSel = 3
        .Sort = 5
    End If

    .Rows = .Rows + 1
    .row = .Rows - 1
    If (.row = 1) Then
        .TextMatrix(.row, 1) = "keine Artikel zugeordnet !"
    Else
        .TextMatrix(.row, 1) = Format(.Rows - 2, "0") + " Positionen"
        .TextMatrix(.row, 6) = Format(zWertSumme#, "0.00")
    End If

    .FillStyle = flexFillRepeat
    .row = 1
    .col = 7
    .RowSel = .Rows - 1
    .ColSel = .col
    .CellFontName = "Symbol"
    .FillStyle = flexFillSingle
    
    
    .Top = wpara.TitelY
    .Left = wpara.LinksX
    .Height = 11 * .RowHeight(0) + 90

    .SelectionMode = flexSelectionByRow
    
    .row = 1
    .col = 0
    .ColSel = .Cols - 1
End With
    

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

Me.Width = flxSchwellProtArtikel.Left + flxSchwellProtArtikel.Width + 2 * wpara.LinksX

With cmdEsc
    .Top = flxSchwellProtArtikel.Top + flxSchwellProtArtikel.Height + 150 * wpara.BildFaktor
    .Width = wpara.ButtonX%
    .Height = wpara.ButtonY%
    .Left = (ScaleWidth - .Width) / 2
End With

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY% + 90 + wpara.FrmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2


h$ = "Von Schwellwert-Automatik zugeordnete Artikel"
Caption = h$

Call DefErrPop
End Sub

