VERSION 5.00
Begin VB.Form frmHinweis 
   Caption         =   "Bestellung senden"
   ClientHeight    =   7020
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   9855
   Icon            =   "hinweis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   9855
   Visible         =   0   'False
   Begin VB.Timer tmrFocus 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4080
      Top             =   5040
   End
   Begin VB.Timer tmrWarnung 
      Interval        =   1000
      Left            =   4080
      Top             =   5760
   End
   Begin VB.CommandButton cmdWarnungOk 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   480
      TabIndex        =   29
      Top             =   6120
      Width           =   1200
   End
   Begin VB.CommandButton cmdWarnungWeg 
      Caption         =   "Streichen"
      Height          =   450
      Left            =   2400
      TabIndex        =   28
      Top             =   6120
      Width           =   1200
   End
   Begin VB.Frame fmeZeit 
      Caption         =   "Rufzeiteneintrag"
      Height          =   2295
      Left            =   5280
      TabIndex        =   22
      Top             =   4200
      Width           =   4575
      Begin VB.PictureBox picZeit 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         DrawMode        =   10  'Stift maskieren
         FillColor       =   &H8000000D&
         FillStyle       =   0  'Ausgefüllt
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   720
         ScaleHeight     =   435
         ScaleWidth      =   2940
         TabIndex        =   23
         Top             =   915
         Width           =   3000
      End
      Begin VB.Label lblZeit 
         Caption         =   "Eingetragene Rufzeit: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblZeit 
         Caption         =   "Aktuelle Uhrzeit: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   1635
         Width           =   2295
      End
      Begin VB.Label lblZeitWert 
         Caption         =   "999:99:99"
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
         Left            =   3000
         TabIndex        =   25
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblZeitWert 
         Caption         =   "999:99:99"
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
         Left            =   3000
         TabIndex        =   24
         Top             =   1635
         Width           =   1095
      End
   End
   Begin VB.Frame fmeAuftrag 
      Caption         =   "Auftrag"
      Height          =   4335
      Left            =   120
      TabIndex        =   19
      Top             =   360
      Width           =   5055
      Begin VB.ComboBox cboAuftrag 
         Height          =   315
         Index           =   1
         Left            =   3240
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   2280
         Width           =   1695
      End
      Begin VB.ComboBox cboAuftrag 
         Height          =   315
         Index           =   0
         Left            =   3240
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CheckBox chkAuftrag 
         Caption         =   "Absagen &berücksichtigen"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   3840
         Width           =   4335
      End
      Begin VB.CheckBox chkAuftrag 
         Caption         =   "&Rückmeldungen anzeigen"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   3480
         Width           =   4335
      End
      Begin VB.OptionButton optAuftrag 
         Caption         =   "&Warten auf Anruf"
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
         Left            =   360
         TabIndex        =   5
         Top             =   2640
         Width           =   3655
      End
      Begin VB.OptionButton optAuftrag 
         Caption         =   "Anruf &durchführen"
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
         Left            =   360
         TabIndex        =   4
         Top             =   2160
         Width           =   3655
      End
      Begin VB.TextBox txtAuftrag 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   3840
         MaxLength       =   2
         TabIndex        =   3
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtAuftrag 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   3840
         MaxLength       =   2
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblAuftrag 
         Caption         =   "Auftrags&art"
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
         Left            =   360
         TabIndex        =   21
         Top             =   1200
         Width           =   3495
      End
      Begin VB.Label lblAuftrag 
         Caption         =   "Auftrags&ergänzung"
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
         Left            =   360
         TabIndex        =   20
         Top             =   480
         Width           =   3495
      End
   End
   Begin VB.Frame fmeLieferant 
      Caption         =   "Lieferant"
      Height          =   3615
      Left            =   5400
      TabIndex        =   8
      Top             =   360
      Width           =   5055
      Begin VB.Label lblLieferant 
         Caption         =   "IDF-Apotheke"
         Height          =   300
         Index           =   0
         Left            =   360
         TabIndex        =   18
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblLieferant 
         Caption         =   "IDF-Lieferant"
         Height          =   300
         Index           =   1
         Left            =   360
         TabIndex        =   17
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label lblLieferant 
         Caption         =   "Tel-Lieferant"
         Height          =   300
         Index           =   2
         Left            =   360
         TabIndex        =   16
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label lblLieferant 
         Height          =   300
         Index           =   3
         Left            =   360
         TabIndex        =   15
         Top             =   1980
         Width           =   3975
      End
      Begin VB.Label lblLieferant 
         Height          =   300
         Index           =   4
         Left            =   360
         TabIndex        =   14
         Top             =   2340
         Width           =   3975
      End
      Begin VB.Label lblLieferant 
         Height          =   300
         Index           =   5
         Left            =   360
         TabIndex        =   13
         Top             =   2700
         Width           =   3975
      End
      Begin VB.Label lblLieferant 
         Height          =   300
         Index           =   6
         Left            =   360
         TabIndex        =   12
         Top             =   3060
         Width           =   3975
      End
      Begin VB.Label lblLieferantWert 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   2400
         TabIndex        =   11
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblLieferantWert 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2400
         TabIndex        =   10
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblLieferantWert 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   2400
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmHinweis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim StartZeit&

Private Const DefErrModul = "HINWEIS.FRM"

Private Sub cmdWarnungWeg_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdWarnungWeg_Click")
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
Dim IstDatum&

IstDatum& = Val(Format(Day(Date), "00") + Format(Month(Date), "00") + Format(Year(Date), "0000"))
Rufzeiten(AutomaticInd%).LetztSend = IstDatum&
Rufzeiten(AutomaticInd%).Gewarnt = "N"
Call SpeicherIniRufzeiten

Unload Me
Call DefErrPop
End Sub

Private Sub cmdWarnungOk_Click()

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdWarnungOk_Click")
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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%, iLieferant%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%
Dim h$, h2$, FormStr$, sZeit$


iLieferant% = Rufzeiten(AutomaticInd%).Lieferant

Call HoleLieferantenDaten(iLieferant%)

Caption = "Sendeauftrag für " + LiefName1$
txtAuftrag(0).text = "ZH"
txtAuftrag(1).text = "  "
optAuftrag(0).Value = True
chkAuftrag(0).Value = 1
chkAuftrag(1).Value = 1
lblLieferant(3).Caption = LiefName1$
lblLieferant(4).Caption = LiefName2$
lblLieferant(5).Caption = LiefName3$
lblLieferant(6).Caption = LiefName4$
lblLieferantWert(0).Caption = ApoIDF$
lblLieferantWert(1).Caption = GhIDF$
lblLieferantWert(2).Caption = TelGh$
fmeAuftrag.Caption = "Auftrag"

'lblModemWert.Caption = ZeigeModemTyp$




'Call InitBestellung

Call wpara.InitFont(Me)


fmeAuftrag.Left = wpara.LinksX%
fmeAuftrag.Top = wpara.TitelY%

txtAuftrag(0).Top = 2 * wpara.TitelY%
For i% = 1 To 1
    txtAuftrag(i%).Top = txtAuftrag(i% - 1).Top + txtAuftrag(i% - 1).Height + 90
Next i%

lblAuftrag(0).Left = wpara.LinksX%
lblAuftrag(0).Top = txtAuftrag(0).Top
For i% = 1 To 1
    lblAuftrag(i%).Left = lblAuftrag(i% - 1).Left
    lblAuftrag(i%).Top = txtAuftrag(i%).Top
Next i%

MaxWi% = 0
For i% = 0 To 1
    wi% = lblAuftrag(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

txtAuftrag(0).Left = lblAuftrag(0).Left + MaxWi% + 300
For i% = 1 To 1
    txtAuftrag(i%).Left = txtAuftrag(i% - 1).Left
Next i%

optAuftrag(0).Left = wpara.LinksX%
optAuftrag(0).Top = txtAuftrag(1).Top + txtAuftrag(1).Height + 300
For i% = 1 To 1
    optAuftrag(i%).Left = optAuftrag(i% - 1).Left
    optAuftrag(i%).Top = optAuftrag(i% - 1).Top + optAuftrag(i% - 1).Height + 45
Next i%

chkAuftrag(0).Left = wpara.LinksX%
chkAuftrag(0).Top = optAuftrag(1).Top + optAuftrag(1).Height + 300
For i% = 1 To 1
    chkAuftrag(i%).Left = chkAuftrag(i% - 1).Left
    chkAuftrag(i%).Top = chkAuftrag(i% - 1).Top + chkAuftrag(i% - 1).Height + 45
Next i%



With cboAuftrag(0)
    .Width = TextWidth("Zeilenwert-Auftrag (ZW)") + 300 + wpara.FrmScrollHeight
    
    .Clear
    .AddItem "Zustellung Heute (ZH)"
    .AddItem "Zustellung Morgen (ZM)"
    .AddItem "Heute kein Auftrag (KA)"
    
    .Top = txtAuftrag(0).Top
    .Left = txtAuftrag(0).Left
    .ListIndex = 0
    .Visible = True
End With

With cboAuftrag(1)
    .Width = cboAuftrag(0).Width
    
    .Clear
    .AddItem "Normalauftrag (  )"
    .AddItem "Inventurliste (IN)"
    .AddItem "Lochkarten (LK)"
    .AddItem "Rückkauf-Anfrage (RK)"
    .AddItem "SBL-Auftrag (SB)"
    .AddItem "Sonder-Auftrag (SO)"
    .AddItem "Stapel-Auftrag (ST)"
    .AddItem "Test-Auftrag (TE)"
    .AddItem "Verfalldatenliste (VD)"
    .AddItem "Vorratskauf (VR)"
    .AddItem "10er-Auftrag (ZE)"
    .AddItem "Zeit-Auftrag (ZT)"
    .AddItem "Zeilenwert-Auftrag (ZW)"
    
    .Top = txtAuftrag(1).Top
    .Left = txtAuftrag(1).Left
    .ListIndex = 0
    .Visible = True
End With


txtAuftrag(0).Visible = False
txtAuftrag(1).Visible = False



'fmeAuftrag.Width = txtAuftrag(0).Left + txtAuftrag(0).Width + 2 * wpara.LinksX%
fmeAuftrag.Width = cboAuftrag(1).Left + cboAuftrag(1).Width + 2 * wpara.LinksX%
Hoehe1% = chkAuftrag(1).Top + chkAuftrag(1).Height

''''''''''''''''

fmeLieferant.Left = fmeAuftrag.Left + fmeAuftrag.Width + 300
fmeLieferant.Top = wpara.TitelY%

lblLieferant(0).Top = 2 * wpara.TitelY%
For i% = 1 To 2
    lblLieferant(i%).Top = lblLieferant(i% - 1).Top + lblLieferant(i% - 1).Height + 60
Next i%
lblLieferant(3).Top = lblLieferant(2).Top + lblLieferant(2).Height + 180
For i% = 4 To 6
    lblLieferant(i%).Top = lblLieferant(i% - 1).Top + lblLieferant(i% - 1).Height + 15
Next i%

lblLieferant(0).Left = wpara.LinksX%
lblLieferantWert(0).Top = lblLieferant(0).Top
For i% = 1 To 6
    lblLieferant(i%).Left = lblLieferant(i% - 1).Left
    If (i < 3) Then
        lblLieferantWert(i%).Top = lblLieferant(i%).Top
    End If
Next i%

MaxWi% = 0
For i% = 0 To 2
    wi% = lblLieferant(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

lblLieferantWert(0).Left = lblLieferant(0).Left + MaxWi% + 300
For i% = 1 To 2
    lblLieferantWert(i%).Left = lblLieferantWert(i% - 1).Left
Next i%

MaxWi% = 0
For i% = 0 To 2
    wi% = lblLieferantWert(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%
MaxWi% = MaxWi% + lblLieferantWert(0).Left - lblLieferant(0).Left
For i% = 3 To 6
    wi% = lblLieferant(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

fmeLieferant.Width = MaxWi% + 2 * wpara.LinksX%
'fmeLieferant.Width = lblLieferantWert(0).Left + lblLieferantWert(0).Width + 2 * LinksX%
Hoehe2% = lblLieferant(6).Top + lblLieferant(6).Height


If (Hoehe1% > Hoehe2%) Then
    fmeAuftrag.Height = Hoehe1% + wpara.TitelY%
Else
    fmeAuftrag.Height = Hoehe2% + wpara.TitelY%
End If
fmeLieferant.Height = fmeAuftrag.Height




'''''''''''''''
fmeZeit.Top = fmeLieferant.Top + fmeLieferant.Height + 150
fmeZeit.Left = fmeAuftrag.Left
fmeZeit.Width = fmeLieferant.Left + fmeLieferant.Width - fmeZeit.Left


lblZeit(0).Top = 2 * wpara.TitelY%
For i% = 1 To 1
    lblZeit(i%).Top = lblZeit(i% - 1).Top + lblZeit(i% - 1).Height + 60
Next i%

lblZeit(0).Left = wpara.LinksX%
lblZeitWert(0).Top = lblZeit(0).Top
For i% = 1 To 1
    lblZeit(i%).Left = lblZeit(i% - 1).Left
    lblZeitWert(i%).Top = lblZeit(i%).Top
Next i%

MaxWi% = 0
For i% = 0 To 1
    wi% = lblZeit(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

lblZeitWert(0).Left = lblZeit(0).Left + MaxWi% + 300
For i% = 1 To 1
    lblZeitWert(i%).Left = lblZeitWert(i% - 1).Left
Next i%


With picZeit
    .Left = wpara.LinksX%
    .Top = lblZeit(1).Top + lblZeit(1).Height + 300
    .Width = fmeZeit.Width - .Left - wpara.LinksX%
    .Height = picZeit.TextHeight("99 %") + 120
    fmeZeit.Height = .Top + .Height + wpara.TitelY%
End With


cmdWarnungOk.Top = fmeZeit.Top + fmeZeit.Height + 150
cmdWarnungWeg.Top = cmdWarnungOk.Top

Me.Width = fmeLieferant.Left + fmeLieferant.Width + 2 * wpara.LinksX%

cmdWarnungOk.Width = wpara.ButtonX%
cmdWarnungOk.Height = wpara.ButtonY%
cmdWarnungWeg.Width = wpara.ButtonX%
cmdWarnungWeg.Height = wpara.ButtonY%
cmdWarnungOk.Left = (ScaleWidth - (cmdWarnungOk.Width * 2 + 300)) / 2
cmdWarnungWeg.Left = cmdWarnungOk.Left + cmdWarnungWeg.Width + 300

Me.Height = cmdWarnungOk.Top + cmdWarnungOk.Height + wpara.TitelY% + 90 + wpara.FrmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2



''''''''''''''

h$ = "(" + Left$(RTrim$(Rufzeiten(AutomaticInd%).AuftragsErg) + Space$(2), 2) + ")"
With cboAuftrag(0)
    For i% = 0 To (.ListCount - 1)
        .ListIndex = i%
        If (InStr(.text, h$) > 0) Then
            Exit For
        End If
    Next i%
End With

h$ = "(" + Left$(RTrim$(Rufzeiten(AutomaticInd%).AuftragsArt) + Space$(2), 2) + ")"
With cboAuftrag(1)
    For i% = 0 To (.ListCount - 1)
        .ListIndex = i%
        If (InStr(.text, h$) > 0) Then
            Exit For
        End If
    Next i%
End With

If (Rufzeiten(AutomaticInd%).Aktiv = "J") Then
    optAuftrag(0).Value = True
Else
    optAuftrag(1).Value = True
End If

chkAuftrag(0).Value = 0
chkAuftrag(1).Value = 1

fmeAuftrag.Enabled = False


sZeit$ = Format(Rufzeiten(AutomaticInd%).RufZeit, "0000")
lblZeitWert(0).Caption = Left$(sZeit$, 2) + ":" + Mid$(sZeit$, 3)

sZeit$ = Format(Now, "HH:MM")
lblZeitWert(1).Caption = sZeit$

h$ = Format(Now, "HHMMSS")
StartZeit& = Val(Left$(h$, 2)) * 3600& + Val(Mid$(h$, 3, 2)) * 60& + Val(Mid$(h$, 5, 2))

tmrFocus.Enabled = True

Call DefErrPop
End Sub

Private Sub tmrFocus_Timer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("tmrFocus_Timer")
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

tmrFocus.Enabled = False
cmdWarnungOk.SetFocus

Call DefErrPop
End Sub

Private Sub tmrWarnung_Timer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("tmrWarnung_Timer")
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
Dim IstZeit%, rZeit%
Dim lIstZeit&, hZeit&
Dim Prozent!
Dim sZeit$, h$

sZeit$ = Format(Now, "HH:MM")
lblZeitWert(1).Caption = sZeit$

IstZeit% = Val(Format(Now, "HHMM"))
IstZeit% = (IstZeit% \ 100) * 60 + (IstZeit% Mod 100)

rZeit% = Rufzeiten(AutomaticInd%).RufZeit
rZeit% = (rZeit% \ 100) * 60 + (rZeit% Mod 100)

h$ = Format(Now, "HHMMSS")
lIstZeit& = Val(Left$(h$, 2)) * 3600& + Val(Mid$(h$, 3, 2)) * 60& + Val(Mid$(h$, 5, 2))


hZeit& = (rZeit% * 60& - StartZeit&)
If (hZeit& <= 0&) Then
    Prozent! = 100!
Else
    Prozent! = (lIstZeit& - StartZeit&) / (hZeit&) * 100!
End If
h$ = "Verbleibende Zeit: " + Str$(rZeit% * 60& - lIstZeit&) + " Sekunden"
'h$ = Format$(Prozent!, "##0") + " %"
With picZeit
    .Cls
    .CurrentX = (.ScaleWidth - .TextWidth(h$)) \ 2
    .CurrentY = (.ScaleHeight - .TextHeight(h$)) \ 2
    picZeit.Print h$
    picZeit.Line (0, 0)-((Prozent! * .ScaleWidth) \ 100, .ScaleHeight), vbHighlight, BF
'                Call BitBlt(.hdc, 0, 0, .ScaleWidth, .ScaleHeight, .hdc, 0, 0, &HCC0020)
End With

If (rZeit% = IstZeit%) Then
    Call cmdWarnungOk_Click
End If

Call DefErrPop
End Sub

