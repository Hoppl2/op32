VERSION 5.00
Begin VB.Form frmLieferantenWahl 
   Caption         =   "Lieferant auswählen"
   ClientHeight    =   6480
   ClientLeft      =   1050
   ClientTop       =   1770
   ClientWidth     =   7485
   Icon            =   "frmLieferanten.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   7485
   Begin VB.Frame fmeLieferantenWahlManuell 
      Caption         =   "Manuelle Zuordnungen"
      Height          =   3855
      Left            =   3480
      TabIndex        =   11
      Top             =   1800
      Width           =   3615
      Begin VB.CheckBox chkLieferantenWahlExklusiv 
         Caption         =   "&Nur dieser Lieferant"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   3240
         Width           =   3015
      End
      Begin VB.TextBox txtLieferantenWahl 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   2400
         TabIndex        =   7
         Top             =   2580
         Width           =   855
      End
      Begin VB.TextBox txtLieferantenWahl 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   2400
         TabIndex        =   6
         Top             =   1995
         Width           =   855
      End
      Begin VB.TextBox txtLieferantenWahl 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   2400
         TabIndex        =   5
         Top             =   1410
         Width           =   855
      End
      Begin VB.TextBox txtLieferantenWahl 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   2400
         TabIndex        =   4
         Top             =   825
         Width           =   855
      End
      Begin VB.TextBox txtLieferantenWahl 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   2400
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblLieferantenWahl 
         Caption         =   "Bestell&code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   16
         Top             =   2580
         Width           =   2055
      End
      Begin VB.Label lblLieferantenWahl 
         Caption         =   "&Lagercode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   1995
         Width           =   2055
      End
      Begin VB.Label lblLieferantenWahl 
         Caption         =   "&Warengruppen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   1410
         Width           =   2055
      End
      Begin VB.Label lblLieferantenWahl 
         Caption         =   "&bis BM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   825
         Width           =   2055
      End
      Begin VB.Label lblLieferantenWahl 
         Caption         =   "&von BM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CheckBox chkLieferantenWahlOption 
      Caption         =   "&manuelle Zuordnungen"
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   2
      Top             =   960
      Width           =   3735
   End
   Begin VB.CheckBox chkLieferantenWahlOption 
      Caption         =   "mit &Zuordnungstabelle "
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   1
      Top             =   360
      Width           =   3735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   2520
      TabIndex        =   9
      Top             =   5760
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   390
      Left            =   3960
      TabIndex        =   10
      Top             =   5760
      Width           =   1200
   End
   Begin VB.ListBox lstLieferantenWahl 
      Height          =   5325
      ItemData        =   "frmLieferanten.frx":030A
      Left            =   480
      List            =   "frmLieferanten.frx":0311
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmLieferantenWahl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LIEF_VONMENGE = 0
Private Const LIEF_BISMENGE = 1
Private Const LIEF_WARENGRUPPEN = 2
Private Const LIEF_LAGERCODE = 3
Private Const LIEF_BESTELLCODE = 4

Private Const DefErrModul = "frmLieferanten.frm"

Private Sub chkLieferantenWahlOption_Click(index As Integer)
If (index = 1) Then
    If (chkLieferantenWahlOption(1).Value) Then
        fmeLieferantenWahlManuell.Visible = True
    Else
        fmeLieferantenWahlManuell.Visible = False
    End If
End If
End Sub

Private Sub cmdEsc_Click()

Unload Me

End Sub

Private Sub cmdOk_Click()
Dim j%, ind%
Dim h$

With lstLieferantenWahl
    h$ = RTrim$(.text)
End With

ind% = InStr(h$, "(")
h$ = Mid$(h$, ind% + 1)

Lieferant% = Val(Left$(h$, Len(h$) - 1))

zTabelleAktiv% = False
zManuellAktiv% = False

If (chkLieferantenWahlOption(0).Value) Then
    zTabelleAktiv% = True
End If

If (chkLieferantenWahlOption(1).Value) Then
    zManuellAktiv% = True
    
    Exklusiv% = chkLieferantenWahlExklusiv.Value
    
    VonBM% = 0
    h$ = RTrim$(txtLieferantenWahl(LIEF_VONMENGE).text)
    If (Len(h$) > 0) Then VonBM% = Val(h$)
    
    BisBM% = 0
    h$ = RTrim$(txtLieferantenWahl(LIEF_BISMENGE).text)
    If (Len(h$) > 0) Then BisBM% = Val(h$)
    
    WaGr$ = ""
    h$ = RTrim$(txtLieferantenWahl(LIEF_WARENGRUPPEN).text)
    If (Len(h$) > 0) Then WaGr$ = h$
    
    LaCo$ = ""
    h$ = RTrim$(txtLieferantenWahl(LIEF_LAGERCODE).text)
    If (Len(h$) > 0) Then LaCo$ = h$
    
    auto$ = ""
    h$ = RTrim$(txtLieferantenWahl(LIEF_BESTELLCODE).text)
    If (Len(h$) > 0) Then
        If (InStr(h$, "A") > 0) Then auto$ = "+"
        If (InStr(h$, "V") > 0) Then auto$ = auto$ + "V"
        If (InStr(h$, "M") > 0) Then auto$ = auto$ + "M"
    End If
End If

Call frmAction.AuslesenBestellung(True, False, True)
Unload Me

End Sub

Private Sub Form_Load()
Dim j%, ind%, LiefAktiv%, LiefInd%
Dim i%, lInd%, wi%, MaxWi%
Dim h$

Call wpara.InitFont(Me)

lstLieferantenWahl.Left = wpara.LinksX
lstLieferantenWahl.Top = wpara.TitelY
lstLieferantenWahl.Width = Me.TextWidth("WWWWWW (99)WW")

chkLieferantenWahlOption(0).Left = lstLieferantenWahl.Left + lstLieferantenWahl.Width + 750
chkLieferantenWahlOption(0).Top = wpara.TitelY
For i% = 1 To 1
    chkLieferantenWahlOption(i%).Left = chkLieferantenWahlOption(i% - 1).Left
    chkLieferantenWahlOption(i%).Top = chkLieferantenWahlOption(i% - 1).Top + chkLieferantenWahlOption(i% - 1).Height + 300
Next i%

fmeLieferantenWahlManuell.Left = chkLieferantenWahlOption(0).Left
fmeLieferantenWahlManuell.Top = chkLieferantenWahlOption(1).Top + chkLieferantenWahlOption(1).Height + 150

txtLieferantenWahl(0).Top = 2 * wpara.TitelY
For i% = 1 To 4
    txtLieferantenWahl(i%).Top = txtLieferantenWahl(i% - 1).Top + txtLieferantenWahl(i% - 1).Height + 45
Next i%


lblLieferantenWahl(0).Left = wpara.LinksX
lblLieferantenWahl(0).Top = txtLieferantenWahl(0).Top
For i% = 1 To 4
    lblLieferantenWahl(i%).Left = lblLieferantenWahl(i% - 1).Left
    lblLieferantenWahl(i%).Top = txtLieferantenWahl(i%).Top
Next i%

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

MaxWi% = 0
For i% = 0 To 4
    wi% = lblLieferantenWahl(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

txtLieferantenWahl(0).Left = lblLieferantenWahl(0).Left + MaxWi% + 300
For i% = 1 To 4
    txtLieferantenWahl(i%).Left = txtLieferantenWahl(i% - 1).Left
Next i%

chkLieferantenWahlExklusiv.Left = lblLieferantenWahl(0).Left
chkLieferantenWahlExklusiv.Top = lblLieferantenWahl(4).Top + lblLieferantenWahl(4).Height + 300


fmeLieferantenWahlManuell.Height = chkLieferantenWahlExklusiv.Top + chkLieferantenWahlExklusiv.Height + wpara.TitelY
fmeLieferantenWahlManuell.Width = txtLieferantenWahl(0).Left + txtLieferantenWahl(0).Width + 2 * wpara.LinksX

lstLieferantenWahl.Height = fmeLieferantenWahlManuell.Top + fmeLieferantenWahlManuell.Height

cmdOk.Top = lstLieferantenWahl.Top + lstLieferantenWahl.Height + 450
cmdEsc.Top = cmdOk.Top

Me.Width = fmeLieferantenWahlManuell.Left + fmeLieferantenWahlManuell.Width + wpara.LinksX + 90

cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

lInd% = -1
With lstLieferantenWahl
    .Clear
    For i% = 1 To AnzLiefNamen%
        h$ = LiefNamen$(i% - 1)
        .AddItem h$
        If (Lieferant% > 0) And (lInd% < 0) Then
            ind% = InStr(h$, "(")
            h$ = Mid$(h$, ind% + 1)
            If (Lieferant% = Val(Left$(h$, Len(h$) - 1))) Then
                lInd% = i% - 1
            End If
        End If
    Next i%
    If (lInd% < 0) Then
        .ListIndex = 0
    Else
        .ListIndex = lInd%
    End If
End With

chkLieferantenWahlOption(0).Value = 1
chkLieferantenWahlOption(1).Value = False
fmeLieferantenWahlManuell.Visible = False

End Sub

Private Sub txtLieferantenWahl_GotFocus(index As Integer)

With txtLieferantenWahl(index)
    .SelStart = 0
    .SelLength = Len(.text)
End With

End Sub
