VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBstatus 
   Caption         =   "Bestellstatus"
   ClientHeight    =   3555
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   4875
   Icon            =   "frmBstatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4875
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   2760
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid flxbstatus 
      Height          =   2280
      Left            =   0
      TabIndex        =   0
      Top             =   240
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
Attribute VB_Name = "frmBstatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()

Unload Me

End Sub

Private Sub Form_Load()

Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%
Dim h$, h2$, FormStr$

Call wPara1.InitFont(Me)


Font.Bold = False   ' True

With flxbstatus
    If (AnzeigeFensterTyp$ = TEXT_BESTELLSTATUS) Then
        .Cols = 5
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 0
        
        .FormatString = "^Typ|^Datum|^Zeit|^Lieferant|^Menge"
        .Rows = 1
        .SelectionMode = flexSelectionByRow
        
        .ColWidth(0) = TextWidth("WWWWWWWWWWWWWWWW")
        .ColWidth(1) = TextWidth("99.99.9999")
        .ColWidth(2) = TextWidth("9999:99")
        .ColWidth(3) = TextWidth("WWWWWWW (999)")
        .ColWidth(4) = TextWidth("999999") + wPara1.FrmScrollHeight
    ElseIf (AnzeigeFensterTyp$ = TEXT_NACHBEARBEITUNG) Then
        .Cols = 10
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 0
        
        .FormatString = "^Datum|^Pzn|<Artikel|>Menge|<Eh|>BM|>NM|>RM|>LM|<User"
        .Rows = 1
        .SelectionMode = flexSelectionByRow
        
        .ColWidth(0) = TextWidth("99.99.9999")
        .ColWidth(1) = TextWidth(String(9, "9"))
        .ColWidth(2) = TextWidth(String(23, "W"))
        .ColWidth(3) = TextWidth(String(7, "9"))
        .ColWidth(4) = TextWidth(String(3, "W"))
        .ColWidth(5) = TextWidth("9999")
        .ColWidth(6) = TextWidth("9999")
        .ColWidth(7) = TextWidth("9999")
        .ColWidth(8) = TextWidth("9999")
        .ColWidth(9) = TextWidth(String(8, "W")) + wPara1.FrmScrollHeight
    ElseIf (AnzeigeFensterTyp$ = TEXT_ABSAGEN) Then
        .Cols = 6
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 0
        
        .FormatString = "^Datum|^Pzn|<Artikel|>Menge|<Eh|<BM"
        .Rows = 1
        .SelectionMode = flexSelectionByRow
        
        .ColWidth(0) = TextWidth("99.99.9999")
        .ColWidth(1) = TextWidth(String(9, "9"))
        .ColWidth(2) = TextWidth(String(23, "W"))
        .ColWidth(3) = TextWidth(String(7, "9"))
        .ColWidth(4) = TextWidth(String(3, "W"))
        .ColWidth(5) = TextWidth("9999") + wPara1.FrmScrollHeight
    End If

    .Top = wPara1.TitelY
    .Left = wPara1.LinksX
    .Height = .RowHeight(0) * 11 + 90
    
    spBreite% = 0
    For i% = 0 To .Cols - 1
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
    .Rows = 1
End With

Font.Bold = False   ' True

Me.Width = flxbstatus.Width + 2 * wPara1.LinksX

cmdOk.Width = wPara1.ButtonX
cmdOk.Height = wPara1.ButtonY
cmdOk.Left = (Me.ScaleWidth - cmdOk.Width) / 2
cmdOk.Top = flxbstatus.Top + flxbstatus.Height + 150

Me.Height = cmdOk.Top + cmdOk.Height + wPara1.TitelY + 90 + wPara1.FrmCaptionHeight

'Call BstatusBefuellen
'
'Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
'Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

End Sub

