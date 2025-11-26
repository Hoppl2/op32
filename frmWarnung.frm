VERSION 5.00
Begin VB.Form frmWarnung 
   Caption         =   "Rufzeitentabelle"
   ClientHeight    =   3360
   ClientLeft      =   1710
   ClientTop       =   1485
   ClientWidth     =   5655
   ControlBox      =   0   'False
   Icon            =   "frmWarnung.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5655
   Begin VB.Timer tmrWarnung 
      Interval        =   1000
      Left            =   240
      Top             =   2760
   End
   Begin VB.CommandButton cmdWarnungWeg 
      Caption         =   "Tabelle übergehen"
      Height          =   450
      Left            =   3240
      TabIndex        =   1
      Top             =   2640
      Width           =   1200
   End
   Begin VB.CommandButton cmdWarnungOk 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   1320
      TabIndex        =   0
      Top             =   2640
      Width           =   1200
   End
   Begin VB.Label lblWarnungUhrzeitAnzeige 
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
      Left            =   3240
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblWarnungRufzeitAnzeige 
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
      Left            =   3240
      TabIndex        =   6
      Top             =   885
      Width           =   1095
   End
   Begin VB.Label lblWarnungLiefAnzeige 
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
      Left            =   3240
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblWarnungUhrzeit 
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
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label lblWarnungRufzeit 
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
      Left            =   360
      TabIndex        =   3
      Top             =   885
      Width           =   2295
   End
   Begin VB.Label lblWarnungLief 
      Caption         =   "Lieferant:"
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
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmWarnung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdWarnungWeg_Click()
Dim IstDatum&

IstDatum& = Val(Format(Day(Date), "00") + Format(Month(Date), "00") + Format(Year(Date), "0000"))
Rufzeiten(AutomaticInd%).LetztSend = IstDatum&
Rufzeiten(AutomaticInd%).Gewarnt = "N"
Call SpeicherIniRufzeiten

Unload Me
End Sub

Private Sub cmdWarnungOk_Click()

Unload Me
End Sub

Private Sub Form_Load()
Dim iLieferant%
Dim sLieferant$, sZeit$

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

iLieferant% = Rufzeiten(AutomaticInd%).Lieferant
If (iLieferant% > 0) And (iLieferant% < 200) Then
    lif.GetRecord (iLieferant% + 1)
    sLieferant$ = RTrim$(lif.Name(0))
    Call OemToChar(sLieferant$, sLieferant$)
End If
lblWarnungLiefAnzeige.Caption = sLieferant$

sZeit$ = Format(Rufzeiten(AutomaticInd%).RufZeit, "0000")
lblWarnungRufzeitAnzeige.Caption = Left$(sZeit$, 2) + ":" + Mid$(sZeit$, 3)

sZeit$ = Format(Now, "HH:MM")
lblWarnungUhrzeitAnzeige.Caption = sZeit$

'lblWarnung.Caption = "Achtung: Für " + Left$(sZeit$, 2) + ":" + Mid$(sZeit$, 3) + " ist ein Sendevorgang für Lieferant " + sLieferant$ + " eingetragen !"

End Sub

Private Sub tmrWarnung_Timer()
Dim IstZeit%, rZeit%
Dim sZeit$

sZeit$ = Format(Now, "HH:MM")
lblWarnungUhrzeitAnzeige.Caption = sZeit$

IstZeit% = Val(Format(Now, "HHMM"))
IstZeit% = (IstZeit% \ 100) * 60 + (IstZeit% Mod 100)

rZeit% = Rufzeiten(AutomaticInd%).RufZeit
rZeit% = (rZeit% \ 100) * 60 + (rZeit% Mod 100)

If (rZeit% = IstZeit%) Then
    Call cmdWarnungOk_Click
End If

End Sub
