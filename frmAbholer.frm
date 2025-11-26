VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAbholer 
   Caption         =   "Abholerstatus"
   ClientHeight    =   5715
   ClientLeft      =   240
   ClientTop       =   1650
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   10215
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   6960
      TabIndex        =   21
      Top             =   3840
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxAbholer 
      Height          =   1320
      Left            =   2280
      TabIndex        =   20
      Top             =   3840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2328
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
   Begin VB.Label lblAbholer 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   5040
      TabIndex        =   19
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label lblAbholerWert 
      Caption         =   "WWWWWWWWWWWWWWW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   9
      Left            =   7680
      TabIndex        =   18
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label lblAbholer 
      Alignment       =   2  'Zentriert
      Caption         =   "Text von Kasse"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   4800
      TabIndex        =   17
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label lblAbholerWert 
      Caption         =   "WWWWWWWWWWWWWWW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   8
      Left            =   7440
      TabIndex        =   16
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label lblAbholer 
      Alignment       =   2  'Zentriert
      Caption         =   "Bestellmenge"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   4800
      TabIndex        =   15
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label lblAbholerWert 
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   7
      Left            =   7440
      TabIndex        =   14
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label lblAbholer 
      Alignment       =   2  'Zentriert
      Caption         =   "Gebühr/Aconto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4800
      TabIndex        =   13
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label lblAbholerWert 
      Caption         =   "99999.99"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   6
      Left            =   7440
      TabIndex        =   12
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label lblAbholer 
      Alignment       =   2  'Zentriert
      Caption         =   "Rezept-Nummer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4800
      TabIndex        =   11
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label lblAbholerWert 
      Caption         =   "9999999999999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   5
      Left            =   7440
      TabIndex        =   10
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblAbholer 
      Alignment       =   1  'Rechts
      Caption         =   "für"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   9
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblAbholerWert 
      Caption         =   "WWWWWWWWWW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   4
      Left            =   2040
      TabIndex        =   8
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label lblAbholer 
      Alignment       =   1  'Rechts
      Caption         =   "um"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblAbholerWert 
      Caption         =   "99:99"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   3
      Left            =   2040
      TabIndex        =   6
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label lblAbholer 
      Alignment       =   1  'Rechts
      Caption         =   "am"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblAbholerWert 
      Caption         =   "99.99.9999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   2
      Left            =   2040
      TabIndex        =   4
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label lblAbholer 
      Alignment       =   1  'Rechts
      Caption         =   "bei"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblAbholerWert 
      Caption         =   "99 RECHTS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   1
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label lblAbholer 
      Alignment       =   1  'Rechts
      Caption         =   "angelegt von"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblAbholerWert 
      Caption         =   "WWWWWWWWWW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmAbholer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()

Unload Me

End Sub

Private Sub Form_Load()
Dim i%, wi%, MaxWi%

Call clsDialog.BesorgerBefuellen

Call wPara1.InitFont(Me)

lblAbholer(0).Top = wPara1.TitelY
For i% = 1 To 4
    lblAbholer(i%).Top = lblAbholer(i% - 1).Top + lblAbholer(i% - 1).Height + 90
Next i%
For i% = 0 To 4
    lblAbholerWert(i%).Top = lblAbholer(i%).Top
Next i%

lblAbholer(0).Left = wPara1.LinksX
For i% = 1 To 4
    lblAbholer(i%).Left = lblAbholer(i% - 1).Left + lblAbholer(i% - 1).Width - lblAbholer(i%).Width
Next i%

MaxWi% = 0
For i% = 0 To 4
    wi% = lblAbholer(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

lblAbholerWert(0).Left = lblAbholer(0).Left + MaxWi% + 300
For i% = 1 To 4
    lblAbholerWert(i%).Left = lblAbholerWert(i% - 1).Left
Next i%


lblAbholer(5).Top = wPara1.TitelY
For i% = 6 To 9
    lblAbholer(i%).Top = lblAbholer(i% - 1).Top + lblAbholer(i% - 1).Height + 90
Next i%
For i% = 5 To 9
    lblAbholerWert(i%).Top = lblAbholer(i%).Top
Next i%

MaxWi% = 0
For i% = 0 To 4
    wi% = lblAbholerWert(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

lblAbholer(5).Left = lblAbholerWert(0).Left + MaxWi% + 300
For i% = 6 To 9
    lblAbholer(i%).Left = lblAbholer(i% - 1).Left
Next i%

MaxWi% = 0
For i% = 5 To 9
    wi% = lblAbholer(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

lblAbholerWert(5).Left = lblAbholer(5).Left + MaxWi% + 300
For i% = 6 To 9
    lblAbholerWert(i%).Left = lblAbholerWert(i% - 1).Left
Next i%

MaxWi% = 0
For i% = 5 To 9
    wi% = lblAbholerWert(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%


With flxAbholer
    .Cols = 1
    .Rows = 6
    .FixedRows = 1
        
    .FormatString = "Historie"
    .SelectionMode = flexSelectionFree
    
    .Top = lblAbholer(4).Top + lblAbholer(4).Height + 300
    .Left = wPara1.LinksX
'    .ColWidth(0) = lblAbholerWert(5).Left + lblAbholerWert(5).Width + 2 * wPara1.LinksX
    .ColWidth(0) = lblAbholerWert(5).Left + MaxWi% + 2 * wPara1.LinksX
    .Width = .ColWidth(0) + 90
    .Height = .RowHeight(0) * 6 + 90
        
End With

cmdOk.Top = flxAbholer.Top + flxAbholer.Height + 210

Me.Width = flxAbholer.Width + 2 * wPara1.LinksX + 120

cmdOk.Width = wPara1.ButtonX
cmdOk.Height = wPara1.ButtonY
cmdOk.Left = (Me.ScaleWidth - cmdOk.Width) / 2

Me.Height = cmdOk.Top + cmdOk.Height + wPara1.TitelY + 90 + wPara1.FrmCaptionHeight

'Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
'Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

End Sub

'Sub AbholerBefuellen()
'Dim erg%, PersCode%, KundenNr%
'Dim h$, tx$, RezTxt$
'
'erg% = AbholerNummer%(GlobalAbholNr%, GlobalPzn$)
'If (erg% = False) Then Call DefErrPop: Exit Sub
'
'
'Call SucheBesorgerInVk(GlobalAbholNr%, PersCode%, KundenNr%)
'
'Caption = "Abholnummer " + Str$(GlobalAbholNr%)
'
'tx$ = ""
'If (PersCode% > 0) Then
'    tx$ = Personal$(PersCode%)
'End If
'lblAbholerWert(0).Caption = tx$
'
'tx$ = kiste.VonWo \ 2
'If (kiste.VonWo Mod 2) Then
'    tx$ = tx$ + " Rechts"
'Else
'    tx$ = tx$ + " Links"
'End If
'lblAbholerWert(1).Caption = tx$
'
'tx$ = CVDatum(Left$(kiste.VonWann, 2))
'lblAbholerWert(2).Caption = Mid$(tx$, 7, 2) + "." + Mid$(tx$, 5, 2) + "." + Left$(tx$, 4)
'
'lblAbholerWert(3).Caption = Format(Asc(Mid$(kiste.VonWann, 3, 1)), "00") + ":" + Format(Asc(Mid$(kiste.VonWann, 4, 1)), "00")
'
'tx$ = ""
'If (KundenNr% > 0) Then
'    tx$ = HoleKundenName(KundenNr%)
'End If
'lblAbholerWert(4).Caption = tx$
'
'tx$ = ""
'If (kiste.RezeptNr) > 0 Then
'    RezTxt$ = Bcd2ascii(kiste.RezeptEan, 12)
'    RezTxt$ = RezTxt$ + "0": Call EanPruef(RezTxt$)
'    tx$ = RezTxt$
'End If
'lblAbholerWert(5).Caption = tx$
'
'lblAbholerWert(6).Caption = LTrim$(BesorgerAconto$)
'lblAbholerWert(7).Caption = Str$(BesorgerMenge%)
'lblAbholerWert(8).Caption = RTrim$(BesorgerInfo(3))
'lblAbholerWert(9).Caption = RTrim$(BesorgerInfo(4))
'
'End Sub
