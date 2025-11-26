VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRueckmeldungen 
   Caption         =   "Rueckmeldungen"
   ClientHeight    =   3450
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4305
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   2520
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid flxRueck 
      Height          =   2280
      Left            =   0
      TabIndex        =   1
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
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmRueckmeldungen"
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

Call wpara.InitFont(Me)


Font.Bold = False   ' True

With flxRueck
    .Cols = 3
    .Rows = 2
    .FixedRows = 1
    .FixedCols = 0

    .Top = wpara.TitelY
    .Left = wpara.LinksX
    .Height = .RowHeight(0) * 11 + 90
    
    .FormatString = "^Satzart|^Länge|<Text"
    .Rows = 1
    .SelectionMode = flexSelectionByRow
    
    Call RueckmeldungenBefuellen
    
    MaxWi% = 0
    For i% = 1 To .Rows - 1
        h$ = .TextMatrix(i%, 2)
        If (TextWidth(h$) > MaxWi%) Then
            MaxWi% = TextWidth(h$)
        End If
    Next i%
    
    .ColWidth(0) = TextWidth("9999999")
    .ColWidth(1) = TextWidth("999999")
    .ColWidth(2) = MaxWi%

    spBreite% = 0
    For i% = 0 To .Cols - 1
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
    
    If (.Width > (frmAction.Width - 600)) Then
        .Width = frmAction.Width - 600
    End If
End With


Me.Width = flxRueck.Width + 2 * wpara.LinksX

cmdOk.Width = wpara.ButtonX
cmdOk.Height = wpara.ButtonY
cmdOk.Left = (Me.ScaleWidth - cmdOk.Width) / 2
cmdOk.Top = flxRueck.Top + flxRueck.Height + 150

Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

End Sub


