VERSION 5.00
Begin VB.Form frmWarnung 
   Caption         =   "Rufzeitentabelle"
   ClientHeight    =   3360
   ClientLeft      =   3270
   ClientTop       =   4950
   ClientWidth     =   5655
   ControlBox      =   0   'False
   Icon            =   "Warnung.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5655
   Begin VB.PictureBox picProgress 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawMode        =   10  'Stift maskieren
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Ausgefüllt
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   4995
      TabIndex        =   8
      Top             =   1200
      Width           =   5055
   End
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
Dim StartZeit&

Private Const DefErrModul = "WARNUNG.FRM"

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
Dim iLieferant%
Dim sLieferant$, sZeit$, h$

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

h$ = Format(Now, "HHMMSS")
StartZeit& = Val(Left$(h$, 2)) * 3600& + Val(Mid$(h$, 3, 2)) * 60& + Val(Mid$(h$, 5, 2))

'lblWarnung.Caption = "Achtung: Für " + Left$(sZeit$, 2) + ":" + Mid$(sZeit$, 3) + " ist ein Sendevorgang für Lieferant " + sLieferant$ + " eingetragen !"

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
Dim lIstZeit&
Dim Prozent!
Dim sZeit$, h$

sZeit$ = Format(Now, "HH:MM")
lblWarnungUhrzeitAnzeige.Caption = sZeit$

IstZeit% = Val(Format(Now, "HHMM"))
IstZeit% = (IstZeit% \ 100) * 60 + (IstZeit% Mod 100)

rZeit% = Rufzeiten(AutomaticInd%).RufZeit
rZeit% = (rZeit% \ 100) * 60 + (rZeit% Mod 100)

h$ = Format(Now, "HHMMSS")
lIstZeit& = Val(Left$(h$, 2)) * 3600& + Val(Mid$(h$, 3, 2)) * 60& + Val(Mid$(h$, 5, 2))


Prozent! = (lIstZeit& - StartZeit&) / (rZeit% * 60& - StartZeit&) * 100!
h$ = Format$(Prozent!, "##0") + " %"
With picProgress
    .Cls
    .CurrentX = (.ScaleWidth - .TextWidth(h$)) \ 2
    .CurrentY = (.ScaleHeight - .TextHeight(h$)) \ 2
    picProgress.Print h$
    picProgress.Line (0, 0)-((Prozent! * .ScaleWidth) \ 100, .ScaleHeight), vbHighlight, BF
'                Call BitBlt(.hdc, 0, 0, .ScaleWidth, .ScaleHeight, .hdc, 0, 0, &HCC0020)
End With

If (rZeit% = IstZeit%) Then
    Call cmdWarnungOk_Click
End If

Call DefErrPop
End Sub
