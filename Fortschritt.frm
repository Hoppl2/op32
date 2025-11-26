VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlButton.ocx"
Begin VB.Form frmFortschritt 
   AutoRedraw      =   -1  'True
   Caption         =   "Privatrezepte einlesen"
   ClientHeight    =   6630
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   9375
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   9375
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   5400
      Picture         =   "Fortschritt.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   5640
      Picture         =   "Fortschritt.frx":00A9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   5880
      Picture         =   "Fortschritt.frx":015D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picfmeBestVorsDauer 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   8775
      TabIndex        =   9
      Tag             =   "Einspieldauer"
      Top             =   2520
      Width           =   8775
      Begin VB.Label lblFortschrittDauerWert 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "99999999"
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
         Left            =   6000
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblBestVorsDauer 
         Alignment       =   2  'Zentriert
         Caption         =   "Rest"
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
         Left            =   4800
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblFortschrittDauerWert 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "99999999"
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
         Left            =   1800
         TabIndex        =   11
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblBestVorsDauer 
         Alignment       =   2  'Zentriert
         Caption         =   "Bisher"
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
         Left            =   600
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.PictureBox picfmeBestVorsStatus 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1935
      ScaleWidth      =   4935
      TabIndex        =   4
      Tag             =   "Einspielstatus"
      Top             =   120
      Width           =   4935
      Begin VB.Label lblBestVorsStatus 
         Alignment       =   2  'Zentriert
         Caption         =   "Anzahl Datensätze"
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
         TabIndex        =   8
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblFortschrittStatusWert 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "99999999"
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
         Left            =   2880
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblBestVorsStatus 
         Alignment       =   2  'Zentriert
         Caption         =   "Anzahl Privatrezepte"
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
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label lblFortschrittStatusWert 
         Alignment       =   2  'Zentriert
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "99999999"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   1
         Left            =   2880
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.Timer tmrBestVors 
      Interval        =   500
      Left            =   3600
      Top             =   4920
   End
   Begin VB.PictureBox picBestVorsProgress 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawMode        =   10  'Stift maskieren
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Ausgefüllt
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   4920
      ScaleHeight     =   555
      ScaleWidth      =   1275
      TabIndex        =   3
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "Abbruch"
      Height          =   500
      Left            =   1560
      TabIndex        =   0
      Top             =   5280
      Width           =   1200
   End
   Begin ComctlLib.ProgressBar prgBestVors 
      Height          =   255
      Left            =   -120
      TabIndex        =   1
      Top             =   4080
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   495
      Left            =   1560
      TabIndex        =   17
      Top             =   5880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin VB.Label lblBestVorsProzent 
      Alignment       =   2  'Zentriert
      Caption         =   "999%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmFortschritt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "FORTSCHRITT.FRM"

Public FortschrittAbbruch%

Dim iEditModus%

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

FortschrittAbbruch% = True

Call DefErrPop
End Sub

Sub HoleVKPrivatRezepte()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim RezAnz&, l&, LastZeit$
Call DefErrFnc("HoleVKPrivatRezepte")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If RezAnz& > 0 Then     'GS 16.5.02
  l& = WritePrivateProfileString("Rezeptkontrolle", "PrivatrezepteVK", LastZeit$, INI_DATEI)
End If
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
Dim j&, found&, Max&, SatzAnz&
Dim erg%, VKAbZeit%, VKAbdatum%
Dim h$, AbDatum$, xc$
Dim Prozent!, StartZeit!, Dauer!, GesamtDauer!, RestDauer!

If (VerkaufDbOk%) Then
    If (VerkaufAdoDB.SqlServerDB) Then
        Call HoleVKPrivatRezepteDB
    Else
        Call HoleVKPrivatRezepteMDB
    End If
    Call DefErrPop: Exit Sub
End If

h$ = Space$(11)
l& = GetPrivateProfileString("Rezeptkontrolle", "PrivatrezepteVK", h$, h$, 11, INI_DATEI)
h$ = Left$(h$, l&)
If Trim(h$) > "" Then
    AbDatum$ = Left$(h$, 6)
    VKAbZeit% = Val(Mid$(h$, 7, 4))
Else
    AbDatum$ = "010102"
    VKAbZeit% = 0
End If


Call vk.GetRecord(1)
xc$ = MKDate(iDate(AbDatum$))
VKAbdatum% = iDate(AbDatum$)
found& = vk.DatumSuche(xc$)
If found& < 0 Then found& = Abs(found&) + 1
Max& = vk.DateiLen / vk.RecordLen

SatzAnz& = Max& - found& + 1
prgBestVors.Max = SatzAnz&
StartZeit! = Timer

j& = 1
Do While found& <= Max&
    vk.GetRecord (found&)
    If vk.Datum >= VKAbdatum% And vk.zeit > VKAbZeit% Then
        If Val(vk.RezEan) = 0 And vk.gebuehren > 0 Then     'Privatrezept
            RezNr$ = "P"
            ActProgram.SetVerkPtr& = found&
            erg% = ActProgram.RezeptHolen(2)
            If erg% Then
                found& = ActProgram.GetPrivEnde
                Call ActProgram.CalcRezeptWerte
                Call ActProgram.WriteRezeptSpeicher(1)
                RezAnz& = RezAnz& + 1
                LastZeit$ = sDate(vk.Datum) + Format(vk.zeit, "0000")
                VKAbZeit% = 0
            End If
        End If
    End If
    found& = found& + 1
    j& = j& + 1
    
    If (j& Mod 100 = 1) Then
        lblFortschrittStatusWert(0).Caption = j&
        lblFortschrittStatusWert(1).Caption = RezAnz&
        Dauer! = Timer - StartZeit!
        lblFortschrittDauerWert(0).Caption = Format$(Dauer! \ 60, "##0") + ":" + Format$(Dauer! Mod 60, "00")
        Prozent! = (j& / SatzAnz&) * 100!
        If (Prozent! > 0) Then
            GesamtDauer! = (Dauer! / Prozent!) * 100!
        Else
            GesamtDauer! = Dauer!
        End If
        RestDauer! = GesamtDauer! - Dauer!
        lblFortschrittDauerWert(1).Caption = Format$(RestDauer! \ 60, "##0") + ":" + Format$(RestDauer! Mod 60, "00")
        prgBestVors.Value = j&
        lblBestVorsProzent.Caption = Format$(Prozent!, "##0") + " %"
            
        h$ = Format$(Prozent!, "##0") + " %"
        With picBestVorsProgress
            .Cls
            .CurrentX = (.ScaleWidth - .TextWidth(h$)) \ 2
            .CurrentY = (.ScaleHeight - .TextHeight(h$)) \ 2
            picBestVorsProgress.Print h$
            picBestVorsProgress.Line (0, 0)-((Prozent! * .ScaleWidth) \ 100, .ScaleHeight), vbHighlight, BF
        End With
        
        DoEvents
        If (FortschrittAbbruch% = True) Then
            Exit Do
        End If

    End If
Loop
h$ = sDate(vk.Datum) + Format(vk.zeit, "0000")
l& = WritePrivateProfileString("Rezeptkontrolle", "PrivatrezepteVK", h$, INI_DATEI)
Call DefErrPop
End Sub

Sub HoleVKPrivatRezepteDB()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim RezAnz&, l&, LastZeit$
Call DefErrFnc("HoleVKPrivatRezepteDB")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If RezAnz& > 0 Then     'GS 16.5.02
  l& = WritePrivateProfileString("Rezeptkontrolle", "PrivatrezepteVK", LastZeit$, INI_DATEI)
End If
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
Dim j&, found&, Max&, SatzAnz&
Dim erg%, VKAbZeit%, VKAbdatum%
Dim h$, xc$
Dim Prozent!, StartZeit!, Dauer!, GesamtDauer!, RestDauer!
Dim AbDatum As Date

h$ = Space$(11)
l& = GetPrivateProfileString("Rezeptkontrolle", "PrivatrezepteVK", h$, h$, 11, INI_DATEI)
h$ = Left$(h$, l&)
If Trim(h$) > "" Then
    AbDatum = Left$(h$, 2) + "." + Mid$(h$, 3, 2) + ".20" + Mid$(h$, 5, 2) + " " + Mid$(h$, 7, 2) + ":" + Right$(h$, 2) + ":59"
    VKAbZeit% = Val(Mid$(h$, 7, 4))
Else
    AbDatum = "01.01.2002 00:00"
    VKAbZeit% = 0
End If

'Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
'VerkaufRec.index = "Unique"
'VerkaufRec.Seek ">", AbDatum
Max = 0
SQLStr = "SELECT COUNT(*) as iANz FROM Verkauf WHERE Datum>'" + Format(AbDatum, "DD.MM.YYYY HH:MM:SS") + "'"
FabsErrf = VerkaufAdoDB.OpenRecordset(VerkaufAdoRec, SQLStr, 0)
If (FabsErrf = 0) Then
    Max = CheckNullLong(VerkaufAdoRec!iAnz)
End If

LastZeit$ = Format(Now, "DDMMYYHHNN")
If (Max > 0) Then
    SQLStr = "SELECT * FROM Verkauf WHERE Datum>'" + Format(AbDatum, "DD.MM.YYYY HH:MM:SS") + "'"
    SQLStr = SQLStr + " ORDER BY Datum,LaufNr,ZeilenNr"
    FabsErrf = VerkaufAdoDB.OpenRecordset(VerkaufAdoRec, SQLStr, 0)

    prgBestVors.Max = 100
    StartZeit! = Timer

    j& = 1
    Do
        If (VerkaufAdoRec.EOF) Then
            Exit Do
        End If

    '    If (VerkaufRec!RezeptArt = REZEPTART_PRIVAT) Then
        If (VerkaufAdoRec!RezeptArt = 5) Or (VerkaufAdoRec!RezeptArt = 6) Then
            If (VerkaufAdoRec!PrivRezNr > 0) And (Val(VerkaufAdoRec!RezeptNr) = 0) Then
                RezNr$ = "P"
                erg% = ActProgram.RezeptHolen(2)
                If erg% Then
                    Call ActProgram.CalcRezeptWerte
                    Call ActProgram.WriteRezeptSpeicher(1)
                    RezAnz& = RezAnz& + 1
                    LastZeit$ = Format(VerkaufAdoRec!Datum, "DDMMYYHHNN")
                    VKAbZeit% = 0
                End If
            End If
        End If

        VerkaufAdoRec.MoveNext

        found& = found& + 1
        j& = j& + 1

    '    If (j& Mod 100 = 1) Then
            lblFortschrittStatusWert(0).Caption = j&
            lblFortschrittStatusWert(1).Caption = RezAnz&
            Dauer! = Timer - StartZeit!
            lblFortschrittDauerWert(0).Caption = Format$(Dauer! \ 60, "##0") + ":" + Format$(Dauer! Mod 60, "00")
            Prozent! = (j / Max) * 100!
            If (Prozent! > 0) Then
                GesamtDauer! = (Dauer! / Prozent!) * 100!
            Else
                GesamtDauer! = Dauer!
            End If
            RestDauer! = GesamtDauer! - Dauer!
            lblFortschrittDauerWert(1).Caption = Format$(RestDauer! \ 60, "##0") + ":" + Format$(RestDauer! Mod 60, "00")
            prgBestVors.Value = CLng(Prozent!)
            lblBestVorsProzent.Caption = Format$(Prozent!, "##0") + " %"

            h$ = Format$(Prozent!, "##0") + " %"
            With picBestVorsProgress
                .Cls
                .CurrentX = (.ScaleWidth - .TextWidth(h$)) \ 2
                .CurrentY = (.ScaleHeight - .TextHeight(h$)) \ 2
                picBestVorsProgress.Print h$
                picBestVorsProgress.Line (0, 0)-((Prozent! * .ScaleWidth) \ 100, .ScaleHeight), vbHighlight, BF
            End With

            DoEvents
            If (FortschrittAbbruch% = True) Then
                Exit Do
            End If

    '    End If
    Loop
End If

'h$ = Format(VerkaufRec!Datum, "DDMMYYHHNN")
l& = WritePrivateProfileString("Rezeptkontrolle", "PrivatrezepteVK", LastZeit$, INI_DATEI)

Call DefErrPop
End Sub

Sub HoleVKPrivatRezepteMDB()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim RezAnz&, l&, LastZeit$
Call DefErrFnc("HoleVKPrivatRezepteMDB")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If RezAnz& > 0 Then     'GS 16.5.02
  l& = WritePrivateProfileString("Rezeptkontrolle", "PrivatrezepteVK", LastZeit$, INI_DATEI)
End If
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
DefErrEnd:
Dim j&, found&, Max&, SatzAnz&
Dim erg%, VKAbZeit%, VKAbdatum%
Dim h$, xc$
Dim Prozent!, StartZeit!, Dauer!, GesamtDauer!, RestDauer!
Dim AbDatum As Date

h$ = Space$(11)
l& = GetPrivateProfileString("Rezeptkontrolle", "PrivatrezepteVK", h$, h$, 11, INI_DATEI)
h$ = Left$(h$, l&)
If Trim(h$) > "" Then
    AbDatum = Left$(h$, 2) + "." + Mid$(h$, 3, 2) + ".20" + Mid$(h$, 5, 2) + " " + Mid$(h$, 7, 2) + ":" + Right$(h$, 2) + ":59"
    VKAbZeit% = Val(Mid$(h$, 7, 4))
Else
    AbDatum = "01.01.2002 00:00"
    VKAbZeit% = 0
End If

Set VerkaufRec = VerkaufDB.OpenRecordset("Verkauf", dbOpenTable)
VerkaufRec.index = "Unique"
VerkaufRec.Seek ">", AbDatum


prgBestVors.Max = 100
StartZeit! = Timer
LastZeit$ = Format(Now, "DDMMYYHHNN")

j& = 1
Do
    If (VerkaufRec.EOF) Or (VerkaufRec.NoMatch) Then
        Exit Do
    End If
    
'    If (VerkaufRec!RezeptArt = REZEPTART_PRIVAT) Then
    If (VerkaufRec!RezeptArt = 5) Or (VerkaufRec!RezeptArt = 6) Then
        If (VerkaufRec!PrivRezNr > 0) And (Val(VerkaufRec!RezeptNr) = 0) Then
            RezNr$ = "P"
            erg% = ActProgram.RezeptHolen(2)
            If erg% Then
                Call ActProgram.CalcRezeptWerte
                Call ActProgram.WriteRezeptSpeicher(1)
                RezAnz& = RezAnz& + 1
                LastZeit$ = Format(VerkaufRec!Datum, "DDMMYYHHNN")
                VKAbZeit% = 0
            End If
        End If
    End If
    
    VerkaufRec.MoveNext
    
    found& = found& + 1
    j& = j& + 1
    
'    If (j& Mod 100 = 1) Then
        lblFortschrittStatusWert(0).Caption = j&
        lblFortschrittStatusWert(1).Caption = RezAnz&
        Dauer! = Timer - StartZeit!
        lblFortschrittDauerWert(0).Caption = Format$(Dauer! \ 60, "##0") + ":" + Format$(Dauer! Mod 60, "00")
        Prozent! = VerkaufRec.PercentPosition
        If (Prozent! > 0) Then
            GesamtDauer! = (Dauer! / Prozent!) * 100!
        Else
            GesamtDauer! = Dauer!
        End If
        RestDauer! = GesamtDauer! - Dauer!
        lblFortschrittDauerWert(1).Caption = Format$(RestDauer! \ 60, "##0") + ":" + Format$(RestDauer! Mod 60, "00")
        prgBestVors.Value = CLng(Prozent!)
        lblBestVorsProzent.Caption = Format$(Prozent!, "##0") + " %"
            
        h$ = Format$(Prozent!, "##0") + " %"
        With picBestVorsProgress
            .Cls
            .CurrentX = (.ScaleWidth - .TextWidth(h$)) \ 2
            .CurrentY = (.ScaleHeight - .TextHeight(h$)) \ 2
            picBestVorsProgress.Print h$
            picBestVorsProgress.Line (0, 0)-((Prozent! * .ScaleWidth) \ 100, .ScaleHeight), vbHighlight, BF
        End With
        
        DoEvents
        If (FortschrittAbbruch% = True) Then
            Exit Do
        End If

'    End If
Loop

'h$ = Format(VerkaufRec!Datum, "DDMMYYHHNN")
l& = WritePrivateProfileString("Rezeptkontrolle", "PrivatrezepteVK", LastZeit$, INI_DATEI)

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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%, y%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%
Dim iAdd%, iAdd2%
Dim lColor&(1)
Dim h$, h2$, FormStr$
Dim c As Control

iEditModus = 1

Call wpara.InitFont(Me)

For i% = 0 To 1
    lblFortschrittStatusWert(i%).Caption = ""
    lblFortschrittDauerWert(i%).Caption = ""
Next i%
lblBestVorsProzent.Caption = ""

picfmeBestVorsStatus.Left = wpara.LinksX
picfmeBestVorsStatus.Top = wpara.TitelY

lblBestVorsStatus(0).Top = 2 * wpara.TitelY
For i% = 1 To 1
    lblBestVorsStatus(i%).Top = lblBestVorsStatus(i% - 1).Top + lblBestVorsStatus(i% - 1).Height + 90
Next i%
For i% = 0 To 1
    lblFortschrittStatusWert(i%).Top = lblBestVorsStatus(i%).Top
Next i%

lblBestVorsStatus(0).Left = wpara.LinksX
For i% = 1 To 1
    lblBestVorsStatus(i%).Left = lblBestVorsStatus(i% - 1).Left
Next i%

MaxWi% = 0
For i% = 0 To 1
    wi% = lblBestVorsStatus(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

lblFortschrittStatusWert(0).Left = lblBestVorsStatus(0).Left + MaxWi% + 300
For i% = 1 To 1
    lblFortschrittStatusWert(i%).Left = lblFortschrittStatusWert(i% - 1).Left
Next i%

picfmeBestVorsStatus.Width = lblFortschrittStatusWert(0).Left + lblFortschrittStatusWert(0).Width + 2 * wpara.LinksX
picfmeBestVorsStatus.Height = lblBestVorsStatus(1).Top + lblBestVorsStatus(1).Height + wpara.TitelY



picfmeBestVorsDauer.Left = wpara.LinksX
picfmeBestVorsDauer.Top = picfmeBestVorsStatus.Top + picfmeBestVorsStatus.Height + 300

lblBestVorsDauer(0).Top = 2 * wpara.TitelY
For i% = 1 To 1
    lblBestVorsDauer(i%).Top = lblBestVorsDauer(i% - 1).Top + lblBestVorsDauer(i% - 1).Height + 90
Next i%
For i% = 0 To 1
    lblFortschrittDauerWert(i%).Top = lblBestVorsDauer(i%).Top
Next i%

lblBestVorsDauer(0).Left = wpara.LinksX
For i% = 1 To 1
    lblBestVorsDauer(i%).Left = lblBestVorsDauer(i% - 1).Left
Next i%

lblFortschrittDauerWert(0).Left = lblFortschrittStatusWert(0).Left
For i% = 1 To 1
    lblFortschrittDauerWert(i%).Left = lblFortschrittDauerWert(i% - 1).Left
Next i%

picfmeBestVorsDauer.Width = lblFortschrittDauerWert(0).Left + lblFortschrittDauerWert(0).Width + 2 * wpara.LinksX
picfmeBestVorsDauer.Height = lblBestVorsDauer(1).Top + lblBestVorsDauer(1).Height + wpara.TitelY


prgBestVors.Left = wpara.LinksX
prgBestVors.Top = picfmeBestVorsDauer.Top + picfmeBestVorsDauer.Height + 300
prgBestVors.Width = picfmeBestVorsDauer.Width

lblBestVorsProzent.Left = prgBestVors.Left + (prgBestVors.Width - lblBestVorsProzent.Width) / 2
lblBestVorsProzent.Top = prgBestVors.Top + prgBestVors.Height + 150

picBestVorsProgress.Left = wpara.LinksX
picBestVorsProgress.Top = picfmeBestVorsDauer.Top + picfmeBestVorsDauer.Height + 300
picBestVorsProgress.Width = picfmeBestVorsDauer.Width
picBestVorsProgress.Height = picBestVorsProgress.TextHeight("99 %") + 120


'cmdEsc.Top = lblBestVorsProzent.Top + lblBestVorsProzent.Height + 150
cmdEsc.Top = picBestVorsProgress.Top + picBestVorsProgress.Height + 210

Me.Width = picfmeBestVorsDauer.Width + 2 * wpara.LinksX + 120

cmdEsc.Width = wpara.ButtonX
cmdEsc.Height = wpara.ButtonY
cmdEsc.Left = (Me.ScaleWidth - cmdEsc.Width) / 2

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

For Each c In Controls
    If (Left(c.Name, 6) = "picfme") Then
        With c
            .BackColor = Me.BackColor
            
            .ForeColor = vbBlack ' RGB(150, 150, 150)
            .FillStyle = vbFSTransparent ' vbSolid
            .FillColor = vbWhite
        
            y = 30 + .TextHeight(.tag) / 2
            RoundRect .hdc, 30 / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY, (.Width - 30) / Screen.TwipsPerPixelX, (.Height - 30) / Screen.TwipsPerPixelY, 20, 20
            
            c.Line (240, y)-(300 + .TextWidth(.tag) + 90, y), RGB(180, 180, 180)
            
            .CurrentX = 300
            .CurrentY = 30
            c.Print .tag
        End With
    End If
Next

If (para.Newline) Then
    iAdd = wpara.NlFlexBackY
    iAdd2 = wpara.NlCaptionY
    
    On Error Resume Next
    For Each c In Controls
        If (c.Container Is Me) Then
            c.Top = c.Top + iAdd2
        End If
    Next
    On Error GoTo DefErr
    
    
    Height = Height + iAdd2
    
    With nlcmdEsc
        .Init
'        .Left = Me.ScaleWidth - .Width - 150
        .Top = picBestVorsProgress.Top + picBestVorsProgress.Height + iAdd + 600
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .Default = cmdEsc.Default
        .Cancel = cmdEsc.Cancel
        .Visible = True
    End With
    cmdEsc.Visible = False

    nlcmdEsc.Left = (Me.Width - nlcmdEsc.Width) / 2

    Me.Height = nlcmdEsc.Top + nlcmdEsc.Height + wpara.FrmCaptionHeight + 450

    Call wpara.NewLineWindow(Me, nlcmdEsc.Top)
    
    For Each c In Controls
        If (Left(c.Name, 6) = "picfme") Then
            With c
                lColor(0) = GetPixel(.Container.hdc, c.Left / Screen.TwipsPerPixelX - 2, c.Top / Screen.TwipsPerPixelY)
                lColor(1) = GetPixel(.Container.hdc, c.Left / Screen.TwipsPerPixelX - 2, (c.Top + c.Height) / Screen.TwipsPerPixelY)
'                Call wpara.FillGradient(c, 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, RGB(177, 177, 177), RGB(225, 225, 225))
                Call wpara.FillGradient(c, 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, lColor(0), lColor(1))
                
                .ForeColor = vbBlack ' RGB(150, 150, 150)
                .FillStyle = vbFSTransparent ' vbSolid
                .FillColor = vbWhite
            
                y = 30 + .TextHeight(.tag) / 2
                RoundRect .hdc, 30 / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY, (.Width - 30) / Screen.TwipsPerPixelX, (.Height - 30) / Screen.TwipsPerPixelY, 20, 20
                
                c.Line (240, y)-(300 + .TextWidth(.tag) + 90, y), RGB(180, 180, 180)
                
                .CurrentX = 300
                .CurrentY = 30
                c.Print .tag
            End With
        End If
    Next
    
    On Error Resume Next
    For Each c In Controls
        If (c.tag <> "0") Then
            If (TypeOf c Is Label) Then
                c.BackStyle = 0 'duchsichtig
            ElseIf (TypeOf c Is TextBox) Or (TypeOf c Is ComboBox) Then
                If (TypeOf c Is ComboBox) Then
                    Call wpara.ControlBorderless(c)
                ElseIf (c.Appearance = 1) Then
                    Call wpara.ControlBorderless(c, 2, 2)
                Else
                    Call wpara.ControlBorderless(c, 1, 1)
                End If

                If (c.Enabled) Then
                    c.BackColor = vbWhite
                Else
                    c.BackColor = Me.BackColor
                End If

'                If (c.Visible) Then
                    With c.Container
                        .ForeColor = RGB(180, 180, 180) ' vbWhite
                        .FillStyle = vbSolid
                        .FillColor = c.BackColor

                        RoundRect .hdc, (c.Left - 60) / Screen.TwipsPerPixelX, (c.Top - 30) / Screen.TwipsPerPixelY, (c.Left + c.Width + 60) / Screen.TwipsPerPixelX, (c.Top + c.Height + 15) / Screen.TwipsPerPixelY, 10, 10
                    End With
'                End If
            End If
        End If
    Next
    On Error GoTo DefErr
    
Else
    nlcmdEsc.Visible = False
End If
'''''''''

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2


Call DefErrPop
End Sub

Private Sub tmrBestVors_Timer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("tmrBestVors_Timer")
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

tmrBestVors.Enabled = False
Call HoleVKPrivatRezepte
Unload Me

Call DefErrPop
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_MouseDown")
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
    
If (y <= wpara.NlCaptionY) Then
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

Call DefErrPop
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_MouseMove")
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
Dim c As Object

On Error Resume Next
For Each c In Controls
    If (TypeOf c Is nlCommand) Then
        If (c.MouseOver) Then
            c.MouseOver = 0
        End If
    End If
Next
On Error GoTo DefErr

Call DefErrPop
End Sub

Private Sub Form_Resize()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_Resize")
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

If (para.Newline) And (Me.Visible) Then
    CurrentX = wpara.NlFlexBackY
    CurrentY = (wpara.NlCaptionY - TextHeight(Caption)) / 2
    ForeColor = vbBlack
    Me.Print Caption
End If

Call DefErrPop
End Sub

'Private Sub nlcmdOk_Click()
'Call cmdOk_Click
'End Sub

Private Sub nlcmdEsc_Click()
Call cmdEsc_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If (para.Newline) Then
    If (KeyAscii = 13) Then
'        Call nlcmdOk_Click
        Exit Sub
    ElseIf (KeyAscii = 27) And (nlcmdEsc.Visible) Then
        Call nlcmdEsc_Click
        Exit Sub
'    ElseIf (KeyAscii = Asc("<")) And (nlcmdImport(0).Visible) Then
''        Call nlcmdChange_Click(0)
'        nlcmdImport(0).Value = 1
'    ElseIf (KeyAscii = Asc(">")) And (nlcmdImport(1).Visible) Then
''        Call nlcmdChange_Click(1)
'        nlcmdImport(1).Value = 1
    End If
End If
    
If (TypeOf ActiveControl Is TextBox) Then
    If (iEditModus% <> 1) Then
        If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
        If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And (((iEditModus% <> 2) And (iEditModus% <> 4)) Or (Chr$(KeyAscii) <> ".")) Then
            Beep
            KeyAscii = 0
        End If
    End If
End If

End Sub

Private Sub picControlBox_Click(index As Integer)

If (index = 0) Then
    Me.WindowState = vbMinimized
ElseIf (index = 1) Then
    Me.WindowState = vbNormal
Else
    Unload Me
End If

End Sub








