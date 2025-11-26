VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlbutton.ocx"
Begin VB.Form frmnlRezkOptionen 
   AutoRedraw      =   -1  'True
   Caption         =   "Optionen"
   ClientHeight    =   9345
   ClientLeft      =   540
   ClientTop       =   735
   ClientWidth     =   17295
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   17295
   Begin VB.PictureBox picStammdatenBack 
      AutoRedraw      =   -1  'True
      Height          =   1575
      Index           =   7
      Left            =   13080
      ScaleHeight     =   1515
      ScaleWidth      =   3675
      TabIndex        =   39
      Top             =   6360
      Width           =   3735
      Begin MSFlexGridLib.MSFlexGrid flxOptionen1 
         Height          =   780
         Index           =   6
         Left            =   600
         TabIndex        =   40
         Top             =   360
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   1376
         _Version        =   393216
         Rows            =   0
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         BackColorBkg    =   -2147483633
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
      End
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   10200
      Picture         =   "nlRezkOptionen.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   10440
      Picture         =   "nlRezkOptionen.frx":00A9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   10680
      Picture         =   "nlRezkOptionen.frx":015D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picTab 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   735
      Index           =   0
      Left            =   11160
      ScaleHeight     =   735
      ScaleWidth      =   2055
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   960
      Width           =   2055
   End
   Begin VB.PictureBox picStammdatenBack 
      AutoRedraw      =   -1  'True
      Height          =   1575
      Index           =   6
      Left            =   9000
      ScaleHeight     =   1515
      ScaleWidth      =   3675
      TabIndex        =   28
      Top             =   6480
      Width           =   3735
      Begin MSFlexGridLib.MSFlexGrid flxOptionen1 
         Height          =   780
         Index           =   5
         Left            =   600
         TabIndex        =   29
         Top             =   360
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   1376
         _Version        =   393216
         Rows            =   0
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         BackColorBkg    =   -2147483633
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
      End
   End
   Begin VB.PictureBox picStammdatenBack 
      AutoRedraw      =   -1  'True
      Height          =   2775
      Index           =   5
      Left            =   9000
      ScaleHeight     =   2715
      ScaleWidth      =   3915
      TabIndex        =   22
      Top             =   3480
      Width           =   3975
      Begin VB.TextBox txtOptionen5 
         Height          =   495
         Index           =   0
         Left            =   2160
         TabIndex        =   25
         Text            =   "999,999"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtOptionen5 
         Height          =   495
         Index           =   1
         Left            =   2160
         TabIndex        =   27
         Text            =   "999,999"
         Top             =   1920
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid flxOptionen1 
         Height          =   780
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   1376
         _Version        =   393216
         Rows            =   0
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         BackColorBkg    =   -2147483633
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin VB.Label lblOptionen5 
         Caption         =   "Aufschlag &Spezialitäten  (%)"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label lblOptionen5 
         Caption         =   "Aufschlag &Gefässe  (%)"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   2040
         Width           =   2055
      End
   End
   Begin VB.PictureBox picStammdatenBack 
      AutoRedraw      =   -1  'True
      Height          =   1335
      Index           =   4
      Left            =   8760
      ScaleHeight     =   1275
      ScaleWidth      =   3195
      TabIndex        =   20
      Top             =   1920
      Width           =   3255
      Begin MSFlexGridLib.MSFlexGrid flxOptionen1 
         Height          =   660
         Index           =   3
         Left            =   360
         TabIndex        =   21
         Top             =   240
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   1164
         _Version        =   393216
         Rows            =   0
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         BackColorBkg    =   -2147483633
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
      End
   End
   Begin VB.PictureBox picStammdatenBack 
      AutoRedraw      =   -1  'True
      Height          =   1575
      Index           =   3
      Left            =   5160
      ScaleHeight     =   1515
      ScaleWidth      =   3555
      TabIndex        =   18
      Top             =   4680
      Width           =   3615
      Begin MSFlexGridLib.MSFlexGrid flxOptionen1 
         Height          =   900
         Index           =   2
         Left            =   600
         TabIndex        =   19
         Top             =   240
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   1588
         _Version        =   393216
         Rows            =   0
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         BackColorBkg    =   -2147483633
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
      End
   End
   Begin VB.PictureBox picStammdatenBack 
      AutoRedraw      =   -1  'True
      Height          =   1335
      Index           =   2
      Left            =   5160
      ScaleHeight     =   1275
      ScaleWidth      =   3195
      TabIndex        =   16
      Top             =   3240
      Width           =   3255
      Begin MSFlexGridLib.MSFlexGrid flxOptionen1 
         Height          =   780
         Index           =   1
         Left            =   480
         TabIndex        =   17
         Top             =   240
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   1376
         _Version        =   393216
         Rows            =   0
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         BackColorBkg    =   -2147483633
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
      End
   End
   Begin VB.PictureBox picStammdatenBack 
      AutoRedraw      =   -1  'True
      Height          =   1335
      Index           =   1
      Left            =   5160
      ScaleHeight     =   1275
      ScaleWidth      =   3195
      TabIndex        =   14
      Top             =   1800
      Width           =   3255
      Begin MSFlexGridLib.MSFlexGrid flxOptionen1 
         Height          =   780
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Top             =   240
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   1376
         _Version        =   393216
         Rows            =   0
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         BackColorBkg    =   -2147483633
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
      End
   End
   Begin VB.PictureBox picStammdatenBack 
      AutoRedraw      =   -1  'True
      Height          =   5175
      Index           =   0
      Left            =   240
      ScaleHeight     =   5115
      ScaleWidth      =   4635
      TabIndex        =   1
      Top             =   1800
      Width           =   4695
      Begin VB.TextBox txtOptionen0 
         Height          =   495
         Index           =   0
         Left            =   2760
         MaxLength       =   7
         TabIndex        =   3
         Text            =   "WWW9999"
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtOptionen0 
         Height          =   495
         Index           =   3
         Left            =   2760
         TabIndex        =   9
         Text            =   "999,999"
         Top             =   1920
         Width           =   975
      End
      Begin VB.CheckBox chkOptionen0 
         Caption         =   "&Rezepturen mit Kassenrabatt taxieren"
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   2520
         Width           =   6015
      End
      Begin VB.TextBox txtOptionen0 
         Height          =   495
         Index           =   1
         Left            =   2760
         MaxLength       =   38
         TabIndex        =   5
         Text            =   "WWW9999"
         Top             =   840
         Width           =   1575
      End
      Begin VB.CheckBox chkOptionen0 
         Caption         =   "&BTM-Gebühr als Rezeptzeile"
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   3120
         Width           =   6015
      End
      Begin VB.TextBox txtOptionen0 
         Height          =   495
         Index           =   2
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   7
         Text            =   "WWW9999"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox chkOptionen0 
         Caption         =   "Rezeptur auf Rezept &drucken"
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   3600
         Width           =   6015
      End
      Begin VB.CheckBox chkOptionen0 
         Caption         =   "Teilnahme an @Rezept von &AVP"
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   13
         Top             =   4320
         Width           =   3000
      End
      Begin VB.Label lblchkOptionen0 
         Caption         =   "AAAA"
         Height          =   375
         Index           =   0
         Left            =   3240
         TabIndex        =   32
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label lblOptionen0 
         Caption         =   "&Instituts-Kennzeichen"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblOptionen0 
         Caption         =   "&Kassen-Rabatt  (%)"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   3615
      End
      Begin VB.Label lblOptionen0 
         Caption         =   "Rezept-&Text"
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lblOptionen0 
         Caption         =   "Rezept-Text &BTM"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2520
      TabIndex        =   31
      Top             =   7800
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   840
      TabIndex        =   30
      Top             =   7800
      Width           =   1200
   End
   Begin TabDlg.SSTab tabOptionen 
      Height          =   1365
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   2408
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      Tab             =   7
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&1 - Allgemein"
      TabPicture(0)   =   "nlRezkOptionen.frx":0216
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "&2 - A+V Taxierung"
      TabPicture(1)   =   "nlRezkOptionen.frx":0232
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "&3 - Tätigkeiten"
      TabPicture(2)   =   "nlRezkOptionen.frx":024E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "&4 - Abrechnungsdaten"
      TabPicture(3)   =   "nlRezkOptionen.frx":026A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "&5 - Sonderbelege"
      TabPicture(4)   =   "nlRezkOptionen.frx":0286
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "&6 - Parenteral"
      TabPicture(5)   =   "nlRezkOptionen.frx":02A2
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "&7 - Gesamt-Brutto"
      TabPicture(6)   =   "nlRezkOptionen.frx":02BE
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "&8 - Sonderfälle Fiverx"
      TabPicture(7)   =   "nlRezkOptionen.frx":02DA
      Tab(7).ControlEnabled=   -1  'True
      Tab(7).ControlCount=   0
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   375
      Left            =   2520
      TabIndex        =   37
      Top             =   8400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   375
      Left            =   1080
      TabIndex        =   38
      Top             =   8400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
End
Attribute VB_Name = "frmnlRezkOptionen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OrgRezApoDruckName$, OrgBtmRezDruckName$

Const PI = 3.14159265358979

Dim TabNamen$(8)
Dim TabEnabled%(8)
Dim AktTab%
Dim TabsPerRow%
Dim AnzTabs%

Dim iEditModus%

Dim ydiff%


Private Const DefErrModul = "NLREZKOPTIONEN.FRM"

Sub AbDatumEingeben()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AbDatumEingeben")
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
Dim i%, ret%
Dim j%, l%, row%, ind%, aRow%
Dim s$, h$

EditModus% = 4
            
Load frmEdit

With frmEdit
'    .Left = tabOptionen.Left + fmeOptionen(3).Left + flxOptionen1(2).Left + flxOptionen1(2).ColPos(1) + 45
    .Left = picStammdatenBack(0).Left + flxOptionen1(2).Left + flxOptionen1(2).ColPos(1)
    .Left = .Left + Me.Left + wpara.FrmBorderHeight
'    .Top = tabOptionen.Top + fmeOptionen(3).Top + flxOptionen1(2).Top + (flxOptionen1(2).row * flxOptionen1(2).RowHeight(0))
    .Top = picStammdatenBack(0).Top + flxOptionen1(2).Top + (flxOptionen1(2).row - flxOptionen1(2).TopRow + 1) * flxOptionen1(2).RowHeight(0)
    .Top = .Top + Me.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
    .Width = flxOptionen1(2).ColWidth(1)
    .Height = flxOptionen1(2).RowHeight(0)
End With
With frmEdit.txtEdit
    .Width = flxOptionen1(2).ColWidth(1)
    .Left = 0
    .Top = 0
    h$ = flxOptionen1(2).TextMatrix(flxOptionen1(2).row, 1)
    h$ = Left$(h$, 2) + Mid$(h$, 4, 2) + Mid$(h$, 7, 2)
    .text = h$
    .BackColor = vbWhite
    .Visible = True
End With

frmEdit.Show 1

If (EditErg%) Then
    h$ = Trim(EditTxt$)
    h$ = Left$(h$, 2) + "." + Mid$(h$, 3, 2) + "." + Mid$(h$, 5, 2)
    If IsDate(h$) Then
        flxOptionen1(2).TextMatrix(flxOptionen1(2).row, 1) = Format(CDate(h$), "dd.mm.yy")
        With AbrechDatenRec
            .MoveFirst
            Do While Not .EOF
                If AbrechDatenRec!Unique = flxOptionen1(2).row Then
                    .Edit
                    AbrechDatenRec!Datum = Trim(EditTxt$)
                    .Update
                    Exit Do
                End If
                .MoveNext
            Loop
        End With
    End If
End If

Call DefErrPop

End Sub

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

Private Sub cmdOk_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdOk_Click")
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
Dim i%, VERBAND%
Dim l&
Dim h$

If (ActiveControl.Name = cmdOk.Name) Or (ActiveControl.Name = nlcmdOk.Name) Then
    Call AuslesenFlexTaetigkeiten
    
    h$ = Right$(Space$(7) + Trim(txtOptionen0(0).text), 7)
    If (h$ <> OrgRezApoNr$) Then
        OrgRezApoNr$ = h$
        RezApoNr$ = h$
        l& = WritePrivateProfileString("Rezeptkontrolle", "InstitutsKz", h$, INI_DATEI)
    End If
    
    h$ = Trim(txtOptionen0(1).text)
    If (h$ <> OrgRezApoDruckName$) Then
        RezApoDruckName$ = h$
        l& = WritePrivateProfileString("Rezeptkontrolle", "RezeptText", h$, INI_DATEI)
    End If

    h$ = Trim(txtOptionen0(2).text)
    If (h$ <> OrgBtmRezDruckName$) Then
        BtmRezDruckName$ = h$
        l& = WritePrivateProfileString("Rezeptkontrolle", "BtmRezeptText", h$, INI_DATEI)
    End If

'    VmRabattFaktor# = Val(txtOptionen0(1).text)
    VmRabattFaktor# = 100# / (100# - Val(txtOptionen0(3).text))
   
    
    OrgBundesland% = Val(flxOptionen1(0).TextMatrix(flxOptionen1(0).row, 1))
    
    VERBAND% = FileOpen("verbandm.dat", "RW", "B")
    
    h$ = MKI(OrgBundesland%)
    Seek #VERBAND%, 7
    Put #VERBAND%, , h$
    
    h$ = Format(VmRabattFaktor#, "0.0000")
    For i% = 1 To Len(h$)
        If (Mid$(h$, i%, 1) = ",") Then Mid$(h$, i%, 1) = "."
    Next i%
    Seek #VERBAND%, 9
    Put #VERBAND%, , h$
    Close #VERBAND%
    
    h$ = "N"
    RezepturMitFaktor% = False
    If (chkOptionen0(0).Value) Then
        RezepturMitFaktor% = True
        h$ = "J"
    End If
    l& = WritePrivateProfileString("Rezeptkontrolle", "RezepturMitFaktor", h$, INI_DATEI)
    
    h$ = "N"
    BtmAlsZeile% = False
    If (chkOptionen0(1).Value) Then
        BtmAlsZeile% = True
        h$ = "J"
    End If
    l& = WritePrivateProfileString("Rezeptkontrolle", "BtmAlsZeile", h$, INI_DATEI)
    
    h$ = "N"
    RezepturDruck% = False
    If (chkOptionen0(2).Value) Then
        RezepturDruck% = True
        h$ = "J"
    End If
    l& = WritePrivateProfileString("Rezeptkontrolle", "RezepturDruck", h$, INI_DATEI)
    
    h$ = "N"
    AvpTeilnahme% = False
    If (chkOptionen0(3).Value) Then
        AvpTeilnahme% = True
        h$ = "J"
    End If
    l& = WritePrivateProfileString("Rezeptkontrolle", "AvpTeilnahme", h$, INI_DATEI)
    
    For i% = 0 To UBound(ParenteralPzn)
        h$ = ParenteralPzn(i) + ";" + Format(ParenteralPreis(i), "0.00")
        l& = WritePrivateProfileString("Parenteral", "SonderPzn" + CStr(i), h$, INI_DATEI)
    Next i%

    For i% = 0 To UBound(ParEnteralAufschlag)
        h$ = Format(ParEnteralAufschlag(i), "0.00")
        l& = WritePrivateProfileString("Parenteral", "Aufschlag" + CStr(i), h$, INI_DATEI)
    Next i%

    frmAction.mnuDateiInd(6).Enabled = AvpTeilnahme%
    frmAction.cmdDatei(6).Enabled = AvpTeilnahme%
    
    OptionenNeu% = True
    Call AbrechMonatErmitteln
    Unload Me
ElseIf (ActiveControl.Name = flxOptionen1(0).Name) Then
    If (ActiveControl.index = 1) Then
        Call EditOptionenLstMulti
    ElseIf (ActiveControl.index = 2) Then
        Call AbDatumEingeben
    ElseIf (ActiveControl.index = 3) Then
        Call EditOptionenTxt(ActiveControl.index)
    ElseIf (ActiveControl.index = 4) Then
        Call EditOptionenTxt(ActiveControl.index)
    ElseIf (ActiveControl.index = 5) Then
        Call EditOptionenTxt(ActiveControl.index)
    ElseIf (ActiveControl.index = 6) Then
        Call EditOptionenTxt(ActiveControl.index)
    End If
End If

Call DefErrPop
End Sub

Private Sub flxOptionen1_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxOptionen1_GotFocus")
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

With flxOptionen1(index)
    If (index = 0) Then
        .col = 0
        .ColSel = .Cols - 1
    End If
    .HighLight = flexHighlightAlways
End With

Call DefErrPop
End Sub

Private Sub flxOptionen1_LostFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxOptionen1_LostFocus")
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

With flxOptionen1(index)
    .HighLight = flexHighlightNever
End With

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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%
Dim iAdd%, iAdd2%
Dim h$, h2$, h3$, FormStr$
Dim c As Control


iEditModus = 1

Call wpara.InitFont(Me)

Call RefreshTabControls


With cmdOk
    .Width = wpara.ButtonX
    .Height = wpara.ButtonY
'    .Top = tabOptionen.Top + tabOptionen.Height + 300
    .Top = picStammdatenBack(0).Top + picStammdatenBack(0).Height + 300
End With
With cmdEsc
    .Width = wpara.ButtonX
    .Height = wpara.ButtonY
    .Top = cmdOk.Top
End With


'Me.Width = tabOptionen.Left + tabOptionen.Width + 2 * wpara.LinksX
Me.Width = picStammdatenBack(0).Left + picStammdatenBack(0).Width + 2 * wpara.LinksX
Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

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
'    Width = Width + iAdd2 + 600
    
    With nlcmdOk
        .Init
'        .Left = (Me.ScaleWidth - 2 * .Width - 300)
'        .Top = tabProfil.Top + tabProfil.Height + iAdd + 600
        .Top = picStammdatenBack(0).Top + picStammdatenBack(0).Height + iAdd + 600
        .Caption = cmdOk.Caption
        .TabIndex = cmdOk.TabIndex
        .Enabled = cmdOk.Enabled
        .Default = cmdOk.Default
        .Cancel = cmdOk.Cancel
        .Visible = True
    End With
    cmdOk.Visible = False

    With nlcmdEsc
        .Init
'        .Left = Me.ScaleWidth - .Width - 150
        .Top = nlcmdOk.Top
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .Default = cmdEsc.Default
        .Cancel = cmdEsc.Cancel
        .Visible = True
    End With
    cmdEsc.Visible = False
    
'    Me.Width = nlcmdImport(0).Left + nlcmdImport(0).Width + 600

    nlcmdOk.Left = (Me.Width - (nlcmdOk.Width * 2 + 300)) / 2
    nlcmdEsc.Left = nlcmdOk.Left + nlcmdEsc.Width + 300

    Me.Height = nlcmdOk.Top + nlcmdOk.Height + wpara.FrmCaptionHeight + 450

    Call wpara.NewLineWindow(Me, nlcmdEsc.Top)
'    With flxAbglPartner
'        RoundRect hdc, (.Left - iAdd) / Screen.TwipsPerPixelX, (.Top - iAdd) / Screen.TwipsPerPixelY, (.Left + .Width + iAdd) / Screen.TwipsPerPixelX, (.Top + .Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
'    End With

    On Error Resume Next
    For Each c In Controls
'        If (c.Container Is Me) Then
            If (c.tag <> "0") Then
                If (TypeOf c Is Label) Then
                    c.BackStyle = 0 'duchsichtig
                ElseIf (TypeOf c Is TextBox) Or (TypeOf c Is ComboBox) Or (TypeOf c Is ListBox) Then
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
                ElseIf (TypeOf c Is CheckBox) Or (TypeOf c Is OptionButton) Then
                    With c
'                        .BackColor = GetPixel(.Container.hdc, .Left / Screen.TwipsPerPixelX - 2, .Top / Screen.TwipsPerPixelY)
                        .BackColor = GetPixel(picStammdatenBack(0).hdc, .Left / Screen.TwipsPerPixelX - 2, .Top / Screen.TwipsPerPixelY)
                        .Height = 0
                        .Width = .Height * 3 / 4
                    End With
                    If (c.Name = "chkOptionen0") Then
                        If (c.index > 0) Then
                            Load lblchkOptionen0(c.index)
                        End If
                        With lblchkOptionen0(c.index)
                            .BackStyle = 0 'duchsichtig
                            .Caption = c.Caption
                            .Left = c.Left + c.Width + 60
                            .Top = c.Top
                            .Width = TextWidth(.Caption) + 90
                            .TabIndex = c.TabIndex
                            .Visible = True
                        End With
                    End If
                ElseIf (TypeOf c Is MSFlexGrid) Then
                    With c
'                        .Left = (c.Container.Width - .Width) / 2
                        .Left = (picStammdatenBack(0).Width - .Width) / 2
                        With c.Container
                            .ForeColor = RGB(180, 180, 180) ' vbWhite
                            .FillStyle = vbSolid
                            .FillColor = c.BackColor
                        End With
                        RoundRect c.Container.hdc, (.Left - iAdd) / Screen.TwipsPerPixelX, (.Top - iAdd) / Screen.TwipsPerPixelY, (.Left + .Width + iAdd) / Screen.TwipsPerPixelX, (.Top + .Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
                    End With
                End If
'            End If
        End If
    Next
    On Error GoTo DefErr
    
    
    With Me
        ind = 0
        If (TabsPerRow < AnzTabs) Then
            ind = TabsPerRow
        End If
        
        .ForeColor = RGB(180, 180, 180) ' vbWhite
    
        .FillStyle = vbSolid
        .FillColor = RGB(232, 217, 172)
        .FillColor = RGB(200, 200, 200)
        RoundRect .hdc, picTab(0).Left / Screen.TwipsPerPixelX - 1, picTab(ind).Top / Screen.TwipsPerPixelY - 1, (picStammdatenBack(0).Left + picStammdatenBack(0).Width) / Screen.TwipsPerPixelX + 1, (picStammdatenBack(0).Top + picStammdatenBack(0).Height) / Screen.TwipsPerPixelY + 1 + 10, 20, 20
    
        .FillColor = RGB(200, 200, 200)
    '    RoundRect .hdc, picTab(0).Left / Screen.TwipsPerPixelX - 1, picTab(0).Top / Screen.TwipsPerPixelY - 1, (picStammdatenBack(0).Left + picStammdatenBack(0).Width) / Screen.TwipsPerPixelX + 1, picTab(0).Top / Screen.TwipsPerPixelY + 20, 10, 10
    '    Me.Line (picTab(0).Left + 15, picTab(0).Top + 150)-(picStammdatenBack(0).Left + picStammdatenBack(0).Width - 30, picTab(0).Top + 600), .FillColor, BF
    
        .FillColor = RGB(232, 217, 172)
        .ForeColor = .FillColor
        RoundRect .hdc, picTab(0).Left / Screen.TwipsPerPixelX, (picStammdatenBack(0).Top + picStammdatenBack(0).Height) / Screen.TwipsPerPixelY + 1 - 11, (picStammdatenBack(0).Left + picStammdatenBack(0).Width) / Screen.TwipsPerPixelX, (picStammdatenBack(0).Top + picStammdatenBack(0).Height) / Screen.TwipsPerPixelY + 1 + 9, 20, 20
    End With
    
Else
    nlcmdOk.Visible = False
    nlcmdEsc.Visible = False
End If

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

'tabOptionen.Tab = 0
TabEnabled(0) = True
Call picTab_Click(0)

'Call InitAnimation

Call DefErrPop
End Sub

'txtOptionen0(1).text = String(38, "A")
'txtOptionen0(2).text = String(38, "A")
'
'On Error Resume Next
'For Each c In Controls
'    If (TypeOf c Is TextBox) Then
'        c.Width = TextWidth(c.text) + 90
'        c.text = ""
'    End If
'Next
'On Error GoTo DefErr
'
''''''''''''''''''''''''''''''''''''
'tabOptionen.Tab = 0
'
'lblOptionen0(0).Left = wpara.LinksX
'lblOptionen0(0).Top = 2 * wpara.TitelY
'txtOptionen0(0).Left = lblOptionen0(1).Left + lblOptionen0(1).Width + 300
'txtOptionen0(0).Top = lblOptionen0(0).Top + (lblOptionen0(0).Height - txtOptionen0(0).Height) / 2
'
'For i% = 1 To 3
'    lblOptionen0(i%).Left = lblOptionen0(0).Left
'    lblOptionen0(i%).Top = lblOptionen0(i% - 1).Top + lblOptionen0(i% - 1).Height + 300
'    txtOptionen0(i%).Left = txtOptionen0(0).Left
'    txtOptionen0(i%).Top = lblOptionen0(i%).Top + (lblOptionen0(i%).Height - txtOptionen0(i%).Height) / 2
'Next i%
'
'For i% = 0 To 3
'    With chkOptionen0(i%)
'        .Left = lblOptionen0(0).Left
'        If (i% = 0) Then
'            .Top = lblOptionen0(3).Top + lblOptionen0(3).Height + 600
'        Else
'            .Top = chkOptionen0(i% - 1).Top + chkOptionen0(i% - 1).Height + 150
'        End If
'    End With
'Next i%
'
'txtOptionen0(0).text = OrgRezApoNr$
'OrgRezApoDruckName$ = RezApoDruckName$
'txtOptionen0(1).text = Trim(Left$(RezApoDruckName$ + Space$(38), 38))
'OrgBtmRezDruckName$ = BtmRezDruckName$
'txtOptionen0(2).text = Trim(Left$(BtmRezDruckName$ + Space$(50), 50))
'
'txtOptionen0(3).text = Format(100# - ((1# / VmRabattFaktor#) * 100#), "0.00")
'
'chkOptionen0(0).Value = Abs(RezepturMitFaktor%)
'chkOptionen0(1).Value = Abs(BtmAlsZeile%)
'chkOptionen0(2).Value = Abs(RezepturDruck%)
'chkOptionen0(3).Value = Abs(AvpTeilnahme%)
'
''''''''''''''''''''''''''''''''''''
'tabOptionen.Tab = 1
'Call ActProgram.FlxOptionenBefuellen
'With flxOptionen1(0)
'    Breite1% = 0
'    For i% = 0 To (.Rows - 1)
'        Breite2% = TextWidth(.TextMatrix(i%, 0))
'        If (Breite2% > Breite1%) Then Breite1% = Breite2%
'    Next i%
'    .ColWidth(0) = Breite1% + 150
'    .ColWidth(1) = TextWidth("00000")
'    .ColWidth(2) = wpara.FrmScrollHeight
'
'    .Width = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + 90
'    .Height = .RowHeight(0) * 11 + 90
'
'    .Left = wpara.LinksX
'    .Top = 2 * wpara.TitelY
'
'    .SelectionMode = flexSelectionByRow
'    .col = 0
'    .ColSel = .Cols - 1
'End With
'
''''''''''''''''''''''''''''''''''''
'tabOptionen.Tab = 2
'With flxOptionen1(1)
'    .Cols = 3
'    .Rows = 2
'    .FixedRows = 1
'    .Rows = 1
'
'    .FormatString = "^Tätigkeit|^Personal|^ "
'
'    .ColWidth(0) = TextWidth(String(20, "A")) + 150
'    .ColWidth(1) = TextWidth(String(30, "A"))
'    .ColWidth(2) = 0
'
'    .Width = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + 90
'    .Height = .RowHeight(0) * 11 + 90
'
'    .Left = wpara.LinksX
'    .Top = 2 * wpara.TitelY
'
'    .SelectionMode = flexSelectionFree
'
'    h$ = "Rezeptspeicher" + vbTab
'
'    For i% = 0 To (AnzTaetigkeiten% - 1)
''        h$ = RTrim$(Taetigkeiten(i%).Taetigkeit)
''        h$ = h$ + vbTab
'        h2$ = ""
'        For k% = 0 To 79
'            If (Taetigkeiten(i%).pers(k%) > 0) Then
'                h2$ = h2$ + Mid$(Str$(Taetigkeiten(i%).pers(k%)), 2) + ","
'            Else
'                Exit For
'            End If
'        Next k%
'        If (k% = 1) Then
'            ind% = Taetigkeiten(i%).pers(0)
'            h3$ = RTrim$(para.Personal(ind%))
'        ElseIf (k% > 1) Then
'            h3$ = "mehrere (" + Mid$(Str$(k%), 2) + ")"
'        End If
'        h$ = h$ + h3$ + vbTab + h2$
'    Next i%
'
'    .AddItem h$
'    .row = 1
'End With
''''''''''''''''''''''''''''''''''''
'tabOptionen.Tab = 3
'With flxOptionen1(2)
'    .Rows = 13
'    .Cols = 2
'    .FixedRows = 1
'    .FixedCols = 1
'    .TextMatrix(0, 0) = "Monat"
'    .TextMatrix(0, 1) = "Datum"
'    .ColWidth(0) = TextWidth(String$(10, "A"))
'    .ColWidth(1) = TextWidth(String$(10, "0"))
'    AbrechDatenRec.MoveFirst
'    For i% = 1 To 12
'        .TextMatrix(i%, 0) = Format(CDate("01." + CStr(i%) + ".2002"), "MMMM")
'        .TextMatrix(i%, 1) = Left(AbrechDatenRec!Datum, 2) + "." + Mid(AbrechDatenRec!Datum, 3, 2) + "." + Mid(AbrechDatenRec!Datum, 5, 2)
'        AbrechDatenRec.MoveNext
'    Next i%
'
'    .Width = .ColWidth(0) + .ColWidth(1) + 90
'    .Height = .RowHeight(0) * 13 + 90
'
'    .Left = wpara.LinksX
'    .Top = 2 * wpara.TitelY
'
'    .SelectionMode = flexSelectionByRow
'    .col = 0
'    .ColSel = .Cols - 1
'End With
''''''''''''''''''''''''''''''''''''
'tabOptionen.Tab = 4
'With flxOptionen1(3)
'    .Rows = 11
'    .Cols = 5
'    .FixedRows = 1
'    .FixedCols = 0
'
'    .FormatString = "<SonderPzn|<KK-Bez|<KassenNr|<Status|<GültigBis"
'
'    .ColWidth(0) = TextWidth(String$(10, "9"))
'    .ColWidth(1) = TextWidth(String$(13, "X"))
'    .ColWidth(2) = TextWidth(String$(10, "9"))
'    .ColWidth(3) = TextWidth(String$(10, "9"))
'    .ColWidth(4) = TextWidth(String$(10, "9"))
''    .ColWidth(5) = wpara.FrmScrollHeight
'
'    wi% = 0
'    For i% = 0 To (.Cols - 1)
'        wi% = wi% + .ColWidth(i%)
'    Next i%
'    .Width = wi% + 90
'    .Height = .RowHeight(0) * .Rows + 90
'
'
'    .Left = wpara.LinksX
'    .Top = 2 * wpara.TitelY
'
'    .SelectionMode = flexSelectionFree
'    .GridLines = flexGridFlat
'    .BackColor = vbWhite
'
'    For i% = 1 To AnzSonderBelege%
'        .TextMatrix(i%, 0) = SonderBelege(i% - 1).pzn
'        .TextMatrix(i%, 1) = SonderBelege(i% - 1).KkBez
'        .TextMatrix(i%, 2) = SonderBelege(i% - 1).KassenId
'        .TextMatrix(i%, 3) = SonderBelege(i% - 1).Status
'        .TextMatrix(i%, 4) = SonderBelege(i% - 1).GültigBis
'    Next i%
'
'    .row = .FixedRows
'    .col = 0
'    .ColSel = .col
'End With
''''''''''''''''''''''''''''''''''''
'tabOptionen.Tab = 5
'With flxOptionen1(4)
'    .Rows = 9
'    .Cols = 3
'    .FixedRows = 1
'    .FixedCols = 2
'
'    .FormatString = "<SonderPzn|<Bezeichnung|>ArbeitsPreis (EUR)"
'
'    .ColWidth(0) = TextWidth(String$(10, "9"))
'    .ColWidth(1) = TextWidth(String$(42, "X"))
'    .ColWidth(2) = TextWidth(String$(17, "9"))
'
'    wi% = 0
'    For i% = 0 To (.Cols - 1)
'        wi% = wi% + .ColWidth(i%)
'    Next i%
'    .Width = wi% + 90
'    .Height = .RowHeight(0) * .Rows + 90
'
'    .Left = wpara.LinksX
'    .Top = 2 * wpara.TitelY
'
'    .SelectionMode = flexSelectionFree
'    .GridLines = flexGridFlat
'    .BackColor = vbWhite
'
'    .Rows = .FixedRows
'    For i% = 0 To UBound(ParenteralPzn)
'        h = ParenteralPzn(i) + vbTab + ParenteralTxt(i) + vbTab + Format(ParenteralPreis(i), "0.00")
'        .AddItem h$
'    Next i%
'
'    .row = .FixedRows
'    .col = .Cols - 1
'    .ColSel = .col
'End With
'
'With lblOptionen5(0)
'    .Left = wpara.LinksX
'    .Top = flxOptionen1(4).Top + flxOptionen1(4).Height + 450
'    txtOptionen5(0).Left = .Left + .Width + 150
'    txtOptionen5(0).Top = .Top + (.Height - txtOptionen0(0).Height) / 2
'End With
'For i% = 1 To 1
'    lblOptionen5(i%).Left = lblOptionen5(0).Left
'    lblOptionen5(i%).Top = lblOptionen5(i% - 1).Top + lblOptionen0(i% - 1).Height + 150
'    txtOptionen5(i%).Left = txtOptionen5(0).Left
'    txtOptionen5(i%).Top = lblOptionen5(i%).Top + (lblOptionen5(i%).Height - txtOptionen5(i%).Height) / 2
'Next i%
'For i = 0 To 1
'    txtOptionen5(i).text = Format(ParEnteralAufschlag(i), "0.00")
'Next i
''''''''''''''''''''''''''''''''''''
'tabOptionen.Tab = 6
'With flxOptionen1(5)
'    .Rows = 11
'    .Cols = 2
'    .FixedRows = 1
'    .FixedCols = 0
'
'    .FormatString = "<AbgabeArt|>Preis"
'
'    .ColWidth(0) = TextWidth(String$(30, "X"))
'    .ColWidth(1) = TextWidth(String$(10, "9"))
'
'    wi% = 0
'    For i% = 0 To (.Cols - 1)
'        wi% = wi% + .ColWidth(i%)
'    Next i%
'    .Width = wi% + 90
'    .Height = .RowHeight(0) * .Rows + 90
'
'
'    .Left = wpara.LinksX
'    .Top = 2 * wpara.TitelY
'
'    .SelectionMode = flexSelectionFree
'    .GridLines = flexGridFlat
'    .BackColor = vbWhite
'
'    Call ActProgram.LadeOptionenAbgabeKosten
'
'    .row = .FixedRows
'    .col = 0
'    .ColSel = .col
'End With
''''''''''''''''''''''''''''''''''''
'
'
'Font.Name = wpara.FontName(1)
'Font.Size = wpara.FontSize(1)
'
'With fmeOptionen(0)
'    .Left = wpara.LinksX
'    .Top = 3 * wpara.TitelY
''    .Width = txtOptionen0(1).Left + txtOptionen0(1).Width + 900
''    .Width = flxOptionen1(1).Left + flxOptionen1(1).Width + 1800
'    .Width = flxOptionen1(4).Left + flxOptionen1(4).Width + 900
'
'    Hoehe1% = chkOptionen0(3).Top + chkOptionen0(3).Height
'    Hoehe2% = flxOptionen1(2).Top + flxOptionen1(2).Height
'    If (Hoehe2% > Hoehe1%) Then
'        Hoehe1% = Hoehe2%
'    End If
'    .Height = Hoehe1% + 300
'End With
'For i% = 1 To 6
'    With fmeOptionen(i%)
'        .Left = fmeOptionen(0).Left
'        .Top = fmeOptionen(0).Top
'        .Width = fmeOptionen(0).Width
'        .Height = fmeOptionen(0).Height
'    End With
'Next i%
'
'With tabOptionen
'    .Left = wpara.LinksX
'    .Top = wpara.TitelY
'    .Width = fmeOptionen(0).Left + fmeOptionen(0).Width + wpara.LinksX
'    .Height = fmeOptionen(0).Top + fmeOptionen(0).Height + wpara.TitelY
'End With
'
'
'
'cmdOk.Top = tabOptionen.Top + tabOptionen.Height + 150
'cmdEsc.Top = cmdOk.Top
'
'Me.Width = tabOptionen.Width + 2 * wpara.LinksX
'
'cmdOk.Width = wpara.ButtonX
'cmdOk.Height = wpara.ButtonY
'cmdEsc.Width = wpara.ButtonX
'cmdEsc.Height = wpara.ButtonY
'cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
'cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300
'
'Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight
'
'Breite1% = frmAction.Left + (frmAction.Width - Me.Width) / 2
'If (Breite1% < 0) Then Breite1% = 0
'Me.Left = Breite1%
'Hoehe1% = frmAction.Top + (frmAction.Height - Me.Height) / 2
'If (Hoehe1% < 0) Then Hoehe1% = 0
'Me.Top = Hoehe1%
'
'tabOptionen.Tab = 0
'Call TabDisable
'Call TabEnable(tabOptionen.Tab)
'
'Call DefErrPop
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_KeyDown")
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

If (Shift And vbAltMask) And (KeyCode >= 49) And (KeyCode <= 57) Then
    Call picTab_Click(KeyCode - 49)
End If

'If (KeyCode = vbKeyF2) Then
'    cmdF2.Value = True
'End If

Call DefErrPop
End Sub

'Private Sub tabOptionen_Click(PreviousTab As Integer)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("tabOptionen_Click")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'If (tabOptionen.Visible = False) Then Call DefErrPop: Exit Sub
'
'Call TabDisable
'Call TabEnable(tabOptionen.Tab)
'
'Call DefErrPop
'End Sub

Private Sub flxOptionen1_DblClick(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxOptionen1_DblClick")
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
Call cmdOk_Click
Call DefErrPop
End Sub

'Sub TabDisable()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("TabDisable")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim i%
'
'For i% = 0 To 6
'    fmeOptionen(i%).Visible = False
'Next i%
'
'Call DefErrPop
'End Sub
'
'Sub TabEnable(hTab%)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("TabEnable")
'Call DefErrMod(DefErrModul)
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim i%
'
'fmeOptionen(hTab%).Visible = True
'
'If (hTab% = 0) Then
'    If (txtOptionen0(0).Visible) Then txtOptionen0(0).SetFocus
'ElseIf (hTab% = 1) Then
'    flxOptionen1(0).SetFocus
'ElseIf (hTab% = 2) Then
'    flxOptionen1(1).col = 1
'    flxOptionen1(1).SetFocus
'ElseIf (hTab% = 3) Then
'    flxOptionen1(2).col = 1
'    flxOptionen1(2).SetFocus
'ElseIf (hTab% = 4) Then
'    flxOptionen1(3).col = 0
'    flxOptionen1(3).SetFocus
'ElseIf (hTab% = 5) Then
'    flxOptionen1(4).col = 2
'    flxOptionen1(4).SetFocus
'ElseIf (hTab% = 6) Then
'    flxOptionen1(5).col = 0
'    flxOptionen1(5).SetFocus
'End If
'
'Call DefErrPop
'End Sub

Private Sub txtOptionen0_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtOptionen0_GotFocus")
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
Dim i%
Dim h$

With txtOptionen0(index)
    h$ = .text
    For i% = 1 To Len(h$)
        If (Mid$(h$, i%, 1) = ",") Then Mid$(h$, i%, 1) = "."
    Next i%
    .text = h$
    .SelStart = 0
    .SelLength = Len(.text)
End With

Call DefErrPop
End Sub

Private Sub txtOptionen0_KeyPress(index As Integer, KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtOptionen0_KeyPress")
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

If (index <> 1) And (index <> 2) Then
    If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
    If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And ((index <> 2) Or (Chr$(KeyAscii) <> ".")) Then
        Beep
        KeyAscii = 0
    End If
End If

Call DefErrPop
End Sub

Private Sub txtOptionen5_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtOptionen5_GotFocus")
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
Dim i%
Dim h$

With txtOptionen5(index)
    h$ = .text
    For i% = 1 To Len(h$)
        If (Mid$(h$, i%, 1) = ",") Then Mid$(h$, i%, 1) = "."
    Next i%
    .text = h$
    .SelStart = 0
    .SelLength = Len(.text)
End With

Call DefErrPop
End Sub

Private Sub txtOptionen5_KeyPress(index As Integer, KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtOptionen5_KeyPress")
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

If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And (Chr$(KeyAscii) <> ".") Then
    Beep
    KeyAscii = 0
End If

Call DefErrPop
End Sub

Sub EditOptionenLstMulti()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditOptionenLstMulti")
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
Dim i%, j%, l%, hTab%, row%, col%, ind%, aRow%
Dim s$, h$, BetrLief$, Lief2$

hTab% = tabOptionen.Tab
'row% = 1
col% = 1
                
With flxOptionen1(1)
    row = .row
    aRow% = .row
    .row = 0
    .CellFontBold = True
    .row = aRow%
End With
            
With frmEdit.lstMultiEdit
    .Clear
    .AddItem "(keiner)"
    For i% = 1 To 80
        h$ = para.Personal(i%)
        .AddItem h$
    Next i%

    For i% = 0 To (.ListCount - 1)
        .Selected(i%) = False
    Next i%

    
    Load frmEdit
    
     .ListIndex = 0
     
     BetrLief$ = LTrim$(RTrim$(flxOptionen1(1).TextMatrix(row%, 2)))
     
     For i% = 0 To 19
         If (BetrLief$ = "") Then Exit For
         
         ind% = InStr(BetrLief$, ",")
         If (ind% > 0) Then
             Lief2$ = RTrim$(Left$(BetrLief$, ind% - 1))
             BetrLief$ = LTrim$(Mid$(BetrLief$, ind% + 1))
         Else
             Lief2$ = BetrLief$
             BetrLief$ = ""
         End If
         
         If (Lief2$ <> "") Then
            ind% = Val(Lief2$)
            .Selected(ind%) = True
         End If
     Next i%
     
    With frmEdit
'        .Left = tabOptionen.Left + fmeOptionen(2).Left + flxOptionen1(1).Left + flxOptionen1(1).ColPos(col%) + 45
        .Left = picStammdatenBack(0).Left + flxOptionen1(1).Left + flxOptionen1(1).ColPos(col)
        .Left = .Left + Me.Left + wpara.FrmBorderHeight
'        .Top = tabOptionen.Top + fmeOptionen(2).Top + flxOptionen1(1).Top + flxOptionen1(1).RowHeight(0)
        .Top = picStammdatenBack(0).Top + flxOptionen1(1).Top + (row% - flxOptionen1(1).TopRow + 1) * flxOptionen1(1).RowHeight(0)
        .Top = .Top + Me.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
        .Width = flxOptionen1(1).ColWidth(col%)
        .Height = flxOptionen1(1).Height - flxOptionen1(1).RowHeight(0)
    End With
    With frmEdit.lstMultiEdit
        .Height = frmEdit.ScaleHeight
        frmEdit.Height = .Height
        .Width = frmEdit.ScaleWidth
        .Left = 0
        .Top = 0
        
        .Visible = True
    End With


    frmEdit.Show 1

    With flxOptionen1(1)
        aRow% = .row
        .row = 0
        .CellFontBold = False
        .row = aRow%
    End With
            

    If (EditErg%) Then
            
        flxOptionen1(1).TextMatrix(row%, col% + 1) = EditTxt$
        
        h$ = ""
        If (EditAnzGefunden% = 0) Then
            h$ = ""
        ElseIf (EditAnzGefunden% = 1) Then
            ind% = EditGef%(0)
            h$ = RTrim$(para.Personal(ind%))
        Else
            h$ = "mehrere (" + Mid$(Str$(EditAnzGefunden%), 2) + ")"
        End If

        With flxOptionen1(1)
            .TextMatrix(row%, col%) = h$
            If (.col < .Cols - 2) Then .col = .col + 1
            If (.col < .Cols - 2) And (.ColWidth(.col) = 0) Then .col = .col + 1
        End With
        
    End If

End With

Call DefErrPop
End Sub
                            
Sub AuslesenFlexTaetigkeiten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuslesenFlexTaetigkeiten")
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
Dim i%, j%, k%, ind%, ABG_HANDLE%
Dim l&
Dim h$, BetrLief$, Lief2$, Key$

With flxOptionen1(1)
    j% = 0
    For i% = 1 To (.Rows - 1)
        h$ = RTrim$(.TextMatrix(i%, 0))
        If (h$ <> "") Then  'And (RTrim$(.TextMatrix(i%, 1)) <> "") Then
            Taetigkeiten(j%).Taetigkeit = h$
            BetrLief$ = LTrim$(RTrim$(.TextMatrix(i%, 2)))
            For k% = 0 To 79
                If (BetrLief$ = "") Then Exit For
                
                ind% = InStr(BetrLief$, ",")
                If (ind% > 0) Then
                    Lief2$ = RTrim$(Left$(BetrLief$, ind% - 1))
                    BetrLief$ = LTrim$(Mid$(BetrLief$, ind% + 1))
                Else
                    Lief2$ = BetrLief$
                    BetrLief$ = ""
                End If
                If (Lief2$ <> "") Then
                    Taetigkeiten(j%).pers(k%) = Val(Lief2$)
                End If
            Next k%
            Do
                If (k% > 79) Then Exit Do
                Taetigkeiten(j%).pers(k%) = 0
                k% = k% + 1
            Loop
            j% = j% + 1
        End If
    Next i%
End With
AnzTaetigkeiten = j%

For i = 1 To AnzTaetigkeiten
    h$ = RTrim$(Taetigkeiten(i% - 1).Taetigkeit)
    For j% = 0 To 79
        If (Taetigkeiten(i% - 1).pers(j%) > 0) Then
            h$ = h$ + "," + Mid$(Str$(Taetigkeiten(i% - 1).pers(j%)), 2)
        Else
            Exit For
        End If
    Next j%

    Key$ = "Taetigkeit" + Format(i%, "00")
    l& = WritePrivateProfileString("Rezeptkontrolle", Key$, h$, INI_DATEI)
Next i


With flxOptionen1(3)
    j% = 0
    For i% = 1 To (.Rows - 1)
        h$ = Trim$(.TextMatrix(i%, 0))
'        If (h$ <> "") And (RTrim$(.TextMatrix(i%, 1)) <> "") Then
        If (h$ <> "") Then
            SonderBelege(j%).pzn = Trim$(.TextMatrix(i%, 0))
            SonderBelege(j%).KkBez = Trim$(.TextMatrix(i%, 1))
            SonderBelege(j%).KassenId = Trim$(.TextMatrix(i%, 2))
            SonderBelege(j%).Status = Trim$(.TextMatrix(i%, 3))
            SonderBelege(j%).GültigBis = Trim$(.TextMatrix(i%, 4))
            
            With SonderBelege(j%)
                h$ = .pzn + "," + Trim(.KkBez) + "," + .KassenId + "," + Trim(.Status) + "," + Trim(.GültigBis) + ","
            End With
            Key$ = "SonderBeleg" + Format(j% + 1, "00")
            l& = WritePrivateProfileString("Rezeptkontrolle", Key$, h$, INI_DATEI)
            
            j% = j% + 1
        End If
    Next i%
End With
AnzSonderBelege% = j%
For j% = AnzSonderBelege% To 9
    h$ = ""
    Key$ = "SonderBeleg" + Format(j% + 1, "00")
    l& = WritePrivateProfileString("Rezeptkontrolle", Key$, h$, INI_DATEI)
Next j%

With flxOptionen1(4)
'    For i% = 1 To (.Rows - 1)
'        ParenteralPreis(i - 1) = xVal(.TextMatrix(i, 2))
'    Next i%
    For i% = 1 To 8
        ParenteralPreis(i - 1) = xVal(.TextMatrix(i, 2))
        ParenteralPreis(i + 8 - 1) = xVal(.TextMatrix(i, 3))
    Next i%
End With
For i% = 0 To 1
    ParEnteralAufschlag(i) = xVal(txtOptionen5(i))
Next i%

With flxOptionen1(5)
    ABG_HANDLE% = FileOpen("abgpr.dat", "RW", "B")
    If (ABG_HANDLE% > 0) Then
        j% = 1
        For i% = 1 To (.Rows - 1)
            h$ = Trim$(.TextMatrix(i%, 0))
            If (h$ <> "") Then
                h$ = Left$(h$ + Space(12), 12)
                h$ = h$ + Format(xVal(.TextMatrix(i, 1)) * 100, "0000")
                Put #ABG_HANDLE%, (j% * 16) + 1, h$
                j% = j% + 1
            End If
        Next i
    End If
    h = Space(16)
    For i = j To 10
        Put #ABG_HANDLE%, (i% * 16) + 1, h$
    Next i
    Close #ABG_HANDLE%
End With
Call ActProgram.LadeAbgabePreise

With flxOptionen1(6)
    For i% = 1 To (.Rows - 1)
        h$ = Trim$(.TextMatrix(i%, 0)) + "," + Trim$(.TextMatrix(i%, 1)) + "," + Trim$(.TextMatrix(i%, 2))
        If (h$ <> "") Then
            FiveRxPzns(i - 1) = h
            Key$ = "Pzn" + Format(i%, "0")
            l& = WritePrivateProfileString("Sonderfaelle", Key$, h$, CurDir + "\FiveRxPzn.ini")
        End If
    Next i
End With

Call DefErrPop
End Sub

Function EditOptionenTxt%(FlexInd%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditOptionenTxt%")
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
Dim EditRow%, EditCol%, MaxLen%
Dim h2$

With flxOptionen1(FlexInd)
    EditRow% = .row
    EditCol% = .col
    h2$ = .text
End With

If (FlexInd = 3) Then
    If (EditCol% = 0) Then
        EditModus% = 0
        MaxLen% = 8
    ElseIf (EditCol% = 1) Then
        EditModus% = 1
        MaxLen% = 20
    ElseIf (EditCol% = 2) Then
        EditModus% = 0
        MaxLen% = 7
    ElseIf (EditCol% = 3) Then
        EditModus% = 1
        MaxLen% = 10
    ElseIf (EditCol% = 4) Then
        EditModus% = 1
        MaxLen% = 5
    End If
ElseIf (FlexInd = 5) Then
    If (EditCol% = 0) Then
        EditModus% = 1
        MaxLen% = 12
    Else
        EditModus% = 4
        MaxLen% = 7
    End If
Else
    EditModus% = 4
    MaxLen% = 7
End If

Load frmEdit

With frmEdit
'    .Left = tabOptionen.Left + fmeOptionen(FlexInd + 1).Left + flxOptionen1(FlexInd).Left + flxOptionen1(FlexInd).ColPos(EditCol%) + 45
    .Left = picStammdatenBack(0).Left + flxOptionen1(FlexInd).Left + flxOptionen1(FlexInd).ColPos(EditCol)
    .Left = .Left + Me.Left + wpara.FrmBorderHeight
'    .Top = tabOptionen.Top + fmeOptionen(FlexInd + 1).Top + flxOptionen1(FlexInd).Top + flxOptionen1(FlexInd).RowHeight(0) * EditRow%
    .Top = picStammdatenBack(0).Top + flxOptionen1(FlexInd).Top + (EditRow% - flxOptionen1(FlexInd).TopRow + 1) * flxOptionen1(FlexInd).RowHeight(0)
    .Top = .Top + Me.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
    .Width = flxOptionen1(FlexInd).ColWidth(EditCol%)
    .Height = frmEdit.txtEdit.Height 'flxarbeit(0).RowHeight(1)
End With
With frmEdit.txtEdit
    .Width = frmEdit.ScaleWidth
    .Left = 0
    .Top = 0
    .text = h2$
    .MaxLength = MaxLen%
    .BackColor = vbWhite
    .Visible = True
End With

frmEdit.Show 1
           
If (EditErg%) Then
    With flxOptionen1(FlexInd)
        If (FlexInd = 3) Then
            .text = EditTxt$
            If (EditCol% < (.Cols - 1)) Then
                .col = .col + 1
            ElseIf (EditRow% < (.Rows - 1)) Then
                .row = .row + 1
                .col = 0
            End If
        ElseIf (FlexInd = 5) Then
            If (EditCol = 0) Then
                .text = EditTxt$
            Else
                .text = Format(xVal(EditTxt$), "0.00")
            End If
            If (EditCol% < (.Cols - 1)) Then
                .col = .col + 1
            ElseIf (EditRow% < (.Rows - 1)) Then
                .row = .row + 1
                .col = 0
            End If
        Else
            If (FlexInd = 6) Then
                .text = uFormat(xVal(EditTxt$), "0.00")
            Else
                .text = Format(xVal(EditTxt$), "0.00")
            End If
            If (EditRow% < (.Rows - 1)) Then
                .row = .row + 1
                .col = 2
            End If
        End If
    End With
End If

EditOptionenTxt% = EditErg%

Call DefErrPop
End Function


Private Sub picTab_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("picTab_Click")
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
'If (tabStammdaten.Visible = False) Then Call DefErrPop: Exit Sub

If (TabEnabled(0)) And (TabEnabled(index + 1)) Then
    Call TabDisable
    Call TabEnable(index)
End If

Call DefErrPop
End Sub

Sub TabDisable()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TabDisable")
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
Dim i%

AktTab = -1
For i% = 0 To 7
    Call PaintTab(i)
Next i
'cmdImport(0).Visible = False
'cmdImport(1).Visible = False
'nlcmdImport(0).Visible = False
'nlcmdImport(1).Visible = False

Call DefErrPop
End Sub

Sub TabEnable(hTab%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TabEnable")
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
Dim i%

AktTab = hTab
Call PaintTab(hTab)

If (Me.Visible) Then
    If (hTab% = 0) Then
        If (txtOptionen0(0).Visible) Then txtOptionen0(0).SetFocus
    ElseIf (hTab% = 1) Then
        flxOptionen1(0).SetFocus
    ElseIf (hTab% = 2) Then
        flxOptionen1(1).col = 1
        flxOptionen1(1).SetFocus
    ElseIf (hTab% = 3) Then
        flxOptionen1(2).col = 1
        flxOptionen1(2).SetFocus
    ElseIf (hTab% = 4) Then
        flxOptionen1(3).col = 0
        flxOptionen1(3).SetFocus
    ElseIf (hTab% = 5) Then
        flxOptionen1(4).col = 2
        flxOptionen1(4).SetFocus
    ElseIf (hTab% = 6) Then
        flxOptionen1(5).col = 0
        flxOptionen1(5).SetFocus
    ElseIf (hTab% = 7) Then
        flxOptionen1(6).col = 2
        flxOptionen1(6).SetFocus
    End If
End If

Call DefErrPop
End Sub

Private Sub lblchkOptionen0_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("lblchkOptionen0_Click")
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

With chkOptionen0(index)
    If (.Enabled) Then
        If (.Value) Then
            .Value = 0
        Else
            .Value = 1
        End If
        .SetFocus
    End If
End With

Call DefErrPop
End Sub

Private Sub chkOptionen0_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("chkOptionen0_GotFocus")
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

Call nlCheckBox(chkOptionen0(index).Name, index)

Call DefErrPop
End Sub

Private Sub chkOptionen0_LostFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("chkOptionen0_LostFocus")
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

Call nlCheckBox(chkOptionen0(index).Name, index, 0)

Call DefErrPop
End Sub

Sub nlCheckBox(sCheckBox$, index As Integer, Optional GotFocus% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("nlCheckBox")
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
Dim ok%
Dim Such$
Dim c As Object

Such = "lbl" + sCheckBox

On Error Resume Next
For Each c In Controls
    If (c.Name = Such) Then
        ok = True
        If (index >= 0) Then
            ok = (c.index = index)
        End If
        If (ok) Then
            If (GotFocus) Then
'                c.Font.underline = True
'                c.ForeColor = vbHighlight
                c.BackStyle = 1
                c.BackColor = vbHighlight
                c.ForeColor = vbWhite
            Else
'                c.Font.underline = 0
'                c.ForeColor = vbBlack
                c.BackStyle = 0
                c.BackColor = vbHighlight
                c.ForeColor = vbBlack
            End If
        End If
    End If
Next
On Error GoTo DefErr

Call DefErrPop
End Sub


Sub RefreshTabControls()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RefreshTabControls")
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
Dim i%, j%, k%, TxtHe%, LblHe%, wi%, MaxWi%, iAdd%, x%, spBreite%, iRows%, RowsNeeded%, Breite1%, Breite2%, Hoehe1%, Hoehe2%, xpos%
Dim OrgTab%, RowHe%, AnzZe%, ind%
Dim von&
Dim h$, h2$, h3$, FormStr$, PreisStr$
Dim c As Control

Font.Name = wpara.FontName(0)
Font.Size = wpara.FontSize(0)

AnzTabs = 8
TabsPerRow = 5
For i% = 1 To 7
    Load picTab(i%)
Next i%

With picTab(0)
    .Height = .TextHeight("Äg") * 2 - 60
    
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
    If (TabsPerRow < AnzTabs) Then
        .Top = .Top + .Height
    End If
End With
With picStammdatenBack(0)
'    .Top = picTab(0).Top + .TextHeight("Äg") * 2 - 60 '(300 * wpara.BildFaktor)
    .Top = picTab(0).Top + picTab(0).Height  ' .TextHeight("Äg") * 2 - 60 '(300 * wpara.BildFaktor)
End With
For i = 0 To 7
    TabNamen(i + 1) = Trim(Mid$(tabOptionen.TabCaption(i), 5))
    TabEnabled(i + 1) = True
    Call PaintTab(i)
Next i
Breite1% = picTab(TabsPerRow - 1).Left + picTab(TabsPerRow - 1).Width + 150

For i% = 0 To 0
    With picStammdatenBack(0)
        .Left = picTab(0).Left
        .Top = picTab(0).Top + (.TextHeight("Äg") * 2 - 60) '(300 * wpara.BildFaktor)
        .Width = Breite1%
        .Height = 3000  'tabOptionen.Height - (.Top - tabOptionen.Top) - 210
        .BorderStyle = 0
    End With
Next i
For i% = 1 To 7
    With picStammdatenBack(i)
        .Left = picStammdatenBack(0).Left
        .Top = picStammdatenBack(0).Top
'        .Width = picStammdatenBack(0).Width
'        .Height = picStammdatenBack(0).Height
        .Width = 900
        .Height = 900
        .BorderStyle = picStammdatenBack(0).BorderStyle
    End With
Next i

On Error Resume Next
For Each c In Controls
    If (TypeOf c Is Label) Then
        c.BackStyle = 0 'duchsichtig
    End If
    
    If (c.Name = picStammdatenBack(0).Name) Then
    ElseIf (c.Container.Name = picStammdatenBack(0).Name) Then
        If (Left(c.Name, 10) = "lblSection") Then
            c.Font.Size = c.Font.Size + 4
        End If
        
        If (TypeOf c Is TextBox) Or (TypeOf c Is ComboBox) Then
            c.BackColor = vbWhite
        Else
            c.BackColor = RGB(232, 217, 172)
        End If
        
'        If (TypeOf c Is TextBox) Or (TypeOf c Is Label) Or (TypeOf c Is CheckBox) Or (TypeOf c Is OptionButton) Or (TypeOf c Is ComboBox) Then
        If (TypeOf c Is TextBox) Or (TypeOf c Is Label) Or (TypeOf c Is ComboBox) Then
            Font.Name = c.Font.Name
            Font.Size = c.Font.Size
            Font.Bold = c.Font.Bold
            c.Height = TextHeight("Äg") + 60
            If (TypeOf c Is TextBox) Then
            ElseIf (TypeOf c Is Label) Then
                c.Width = TextWidth(c.Caption) + 90
            Else
                c.Width = TextWidth(c.Caption) + 600
            End If
'        ElseIf (TypeOf c Is CheckBox) Then
'            c.Height = 0
'            c.Width = c.Height
'            If (c.Name = "chkStammdaten2") Then
'                With lblChkCaption2(c.Index)
'                    .Caption = c.Caption
'                    .Left = c.Left + 300
'                    .Top = c.Top
'                    .TabIndex = c.TabIndex
'                End With
'            End If
'        ElseIf (TypeOf c Is OptionButton) Then
        ElseIf (TypeOf c Is MSFlexGrid) Then
            With c
'                .FillStyle = flexFillRepeat
'                .row = 0
'                .col = 0
'                .RowSel = .Rows - 1
'                .ColSel = .Cols - 1
'                .CellFontSize = i%
'                .FillStyle = flexFillSingle
'                .row = .FixedRows
'                .col = .FixedCols
                
                .ScrollBars = flexScrollBarNone
                .BorderStyle = 0
                .GridLines = flexGridFlat
                .GridLinesFixed = .GridLines
                .GridColorFixed = .GridColor
                .BackColor = vbWhite
                .BackColorBkg = vbWhite
                .BackColorFixed = RGB(199, 176, 123)
                If (.SelectionMode = flexSelectionFree) Then
                    .BackColorSel = RGB(135, 61, 52)
                    .ForeColorSel = vbWhite '.ForeColor
                Else
                    .BackColorSel = RGB(232, 217, 172)
                    .ForeColorSel = .ForeColor
                End If
                .Appearance = 0
            End With
        End If
    End If
Next
On Error GoTo DefErr




'txtStammdaten(2).Width = TextWidth(String(7, "X")) + 90
'txtStammdaten2(0).Width = TextWidth(String(8, "9")) + 90
'txtStammdaten7(0).Width = txtStammdaten2(0).Width
'txtStammdaten7(12).Width = TextWidth(String(4, "X")) + 90

LblHe% = lblOptionen0(0).Height
TxtHe% = txtOptionen0(0).Height
ydiff% = (TxtHe% - LblHe) / Screen.TwipsPerPixelY
ydiff% = (ydiff% \ 2) * Screen.TwipsPerPixelY
'yDiff% = TxtHe% - lblStammdaten(0).Height
        
'xadd = 120
'yAdd = yDiff + 30 + 15

Font.Bold = False   ' True

'''''''''''''''''
txtOptionen0(1).text = String(38, "A")
txtOptionen0(2).text = String(38, "A")

On Error Resume Next
For Each c In Controls
    If (TypeOf c Is TextBox) Then
        c.Width = TextWidth(c.text) + 90
        c.text = ""
    End If
Next
On Error GoTo DefErr

'''''''''''''''''''''''''''''''''''
tabOptionen.Tab = 0

With lblOptionen0(0)
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
End With
With txtOptionen0(0)
    .Left = lblOptionen0(1).Left + lblOptionen0(1).Width + 750
    .Top = lblOptionen0(0).Top - ydiff
End With

For i% = 1 To 3
    With lblOptionen0(i%)
        .Left = lblOptionen0(0).Left
        .Top = lblOptionen0(i% - 1).Top + lblOptionen0(i% - 1).Height + 300
    End With
    With txtOptionen0(i%)
        .Left = txtOptionen0(0).Left
        .Top = lblOptionen0(i%).Top - ydiff
    End With
Next i%

For i% = 0 To 3
    With chkOptionen0(i%)
        .Left = lblOptionen0(0).Left
        If (i% = 0) Then
            .Top = lblOptionen0(3).Top + lblOptionen0(3).Height + 600
        Else
            .Top = chkOptionen0(i% - 1).Top + chkOptionen0(i% - 1).Height + 150
        End If
    End With
Next i%

txtOptionen0(0).text = OrgRezApoNr$
OrgRezApoDruckName$ = RezApoDruckName$
txtOptionen0(1).text = Trim(Left$(RezApoDruckName$ + Space$(38), 38))
OrgBtmRezDruckName$ = BtmRezDruckName$
txtOptionen0(2).text = Trim(Left$(BtmRezDruckName$ + Space$(50), 50))

txtOptionen0(3).text = Format(100# - ((1# / VmRabattFaktor#) * 100#), "0.00")

chkOptionen0(0).Value = Abs(RezepturMitFaktor%)
chkOptionen0(1).Value = Abs(BtmAlsZeile%)
chkOptionen0(2).Value = Abs(RezepturDruck%)
chkOptionen0(3).Value = Abs(AvpTeilnahme%)

'''''''''''''''''''''''''''''''''''
tabOptionen.Tab = 1
Call ActProgram.FlxOptionenBefuellen(flxOptionen1(0))
With flxOptionen1(0)
    Breite1% = 0
    For i% = 0 To (.Rows - 1)
        Breite2% = TextWidth(.TextMatrix(i%, 0))
        If (Breite2% > Breite1%) Then Breite1% = Breite2%
    Next i%
    .ColWidth(0) = Breite1% + 150
    .ColWidth(1) = TextWidth("00000")
    .ColWidth(2) = wpara.FrmScrollHeight

    .Width = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) '+ 90
    .Height = .RowHeight(0) * 11 '+ 90
    
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
    
    .ScrollBars = flexScrollBarVertical
    
    .SelectionMode = flexSelectionByRow
    .col = 0
    .ColSel = .Cols - 1
End With

'''''''''''''''''''''''''''''''''''
tabOptionen.Tab = 2
With flxOptionen1(1)
    .Cols = 3
    .Rows = 2
    .FixedRows = 1
    .Rows = 1
    
    .FormatString = "^Tätigkeit|^Personal|^ "
   
    .ColWidth(0) = TextWidth(String(20, "A")) + 150
    .ColWidth(1) = TextWidth(String(30, "A"))
    .ColWidth(2) = 0

    .Width = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) '+ 90
    .Height = .RowHeight(0) * 11 '+ 90
    
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
    
    .SelectionMode = flexSelectionFree
    
    For j = 0 To 1
        If (j = 0) Then
            h$ = "Rezeptspeicher"
        Else
            h$ = "Importkontrolle"
        End If
        
        For i% = 0 To (AnzTaetigkeiten% - 1)
            h2$ = RTrim$(Taetigkeiten(i%).Taetigkeit)
            If (UCase(h) = UCase(h2)) Then
                h2$ = ""
                h3 = ""
                For k% = 0 To 79
                    If (Taetigkeiten(i%).pers(k%) > 0) Then
                        h2$ = h2$ + Mid$(Str$(Taetigkeiten(i%).pers(k%)), 2) + ","
                    Else
                        Exit For
                    End If
                Next k%
                If (k% = 1) Then
                    ind% = Taetigkeiten(i%).pers(0)
                    h3$ = RTrim$(para.Personal(ind%))
                ElseIf (k% > 1) Then
                    h3$ = "mehrere (" + Mid$(Str$(k%), 2) + ")"
                End If
                h3$ = h3$ + vbTab + h2$
            End If
        Next i%
        
        .AddItem h$ + vbTab + h3
    Next j
    .row = 1
End With
'''''''''''''''''''''''''''''''''''
tabOptionen.Tab = 3
With flxOptionen1(2)
    .Rows = 13
    .Cols = 2
    .FixedRows = 1
    .FixedCols = 1
    .TextMatrix(0, 0) = "Monat"
    .TextMatrix(0, 1) = "Datum"
    .ColWidth(0) = TextWidth(String$(10, "A"))
    .ColWidth(1) = TextWidth(String$(10, "0"))
    AbrechDatenRec.MoveFirst
    For i% = 1 To 12
        .TextMatrix(i%, 0) = Format(CDate("01." + CStr(i%) + ".2002"), "MMMM")
        .TextMatrix(i%, 1) = Left(AbrechDatenRec!Datum, 2) + "." + Mid(AbrechDatenRec!Datum, 3, 2) + "." + Mid(AbrechDatenRec!Datum, 5, 2)
        AbrechDatenRec.MoveNext
    Next i%
    
    .Width = .ColWidth(0) + .ColWidth(1) '+ 90
    .Height = .RowHeight(0) * 13 '+ 90
    
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
    
    .SelectionMode = flexSelectionByRow
    .col = 0
    .ColSel = .Cols - 1
End With
'''''''''''''''''''''''''''''''''''
tabOptionen.Tab = 4
With flxOptionen1(3)
    .Rows = 11
    .Cols = 5
    .FixedRows = 1
    .FixedCols = 0
    
    .FormatString = "<SonderPzn|<KK-Bez|<KassenNr|<Status|<GültigBis"
    
    .ColWidth(0) = TextWidth(String$(10, "9"))
    .ColWidth(1) = TextWidth(String$(13, "X"))
    .ColWidth(2) = TextWidth(String$(10, "9"))
    .ColWidth(3) = TextWidth(String$(10, "9"))
    .ColWidth(4) = TextWidth(String$(10, "9"))
'    .ColWidth(5) = wpara.FrmScrollHeight
    
    wi% = 0
    For i% = 0 To (.Cols - 1)
        wi% = wi% + .ColWidth(i%)
    Next i%
    .Width = wi% '+ 90
    .Height = .RowHeight(0) * .Rows '+ 90
    
    
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
    
    .SelectionMode = flexSelectionFree
    .GridLines = flexGridFlat
    .BackColor = vbWhite
    
    For i% = 1 To AnzSonderBelege%
        .TextMatrix(i%, 0) = SonderBelege(i% - 1).pzn
        .TextMatrix(i%, 1) = SonderBelege(i% - 1).KkBez
        .TextMatrix(i%, 2) = SonderBelege(i% - 1).KassenId
        .TextMatrix(i%, 3) = SonderBelege(i% - 1).Status
        .TextMatrix(i%, 4) = SonderBelege(i% - 1).GültigBis
    Next i%
    
    .row = .FixedRows
    .col = 0
    .ColSel = .col
End With
'''''''''''''''''''''''''''''''''''
tabOptionen.Tab = 5
With flxOptionen1(4)
    .Rows = 10
    .Cols = 4
    .FixedRows = 1
    .FixedCols = 2
    
    .FormatString = "<SonderPzn|<Bezeichnung|>ArbeitsPreis GKV|>Privat"
    
    .ColWidth(0) = TextWidth(String$(10, "9"))
    .ColWidth(1) = TextWidth(String$(42, "X"))
    .ColWidth(2) = TextWidth(String$(17, "9"))
    .ColWidth(3) = TextWidth(String$(17, "9"))
    
    wi% = 0
    For i% = 0 To (.Cols - 1)
        wi% = wi% + .ColWidth(i%)
    Next i%
    .Width = wi% '+ 90
    .Height = .RowHeight(0) * .Rows '+ 90
    
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
    
    .SelectionMode = flexSelectionFree
    .GridLines = flexGridFlat
    .BackColor = vbWhite
    
    .Rows = .FixedRows
'    For i% = 0 To UBound(ParenteralPzn)
    For i% = 0 To 7
        h = ParenteralPzn(i) + vbTab + ParenteralTxt(i) + vbTab + Format(ParenteralPreis(i), "0.00") + vbTab + Format(ParenteralPreis(i + 8), "0.00")
        .AddItem h$
    Next i%
    
    .row = .FixedRows
    .col = .Cols - 1
    .ColSel = .col
End With

With lblOptionen5(0)
    .Left = wpara.LinksX
    .Top = flxOptionen1(4).Top + flxOptionen1(4).Height + 450
    txtOptionen5(0).Left = .Left + .Width + 150
    txtOptionen5(0).Top = .Top - ydiff
End With
For i% = 1 To 1
    With lblOptionen5(i%)
        .Left = lblOptionen5(0).Left
        .Top = lblOptionen5(i% - 1).Top + lblOptionen0(i% - 1).Height + 150
    End With
    With txtOptionen5(i%)
        .Left = txtOptionen5(0).Left
        .Top = lblOptionen5(i%).Top - ydiff
    End With
Next i%
For i = 0 To 1
    txtOptionen5(i).text = Format(ParEnteralAufschlag(i), "0.00")
Next i
'''''''''''''''''''''''''''''''''''
tabOptionen.Tab = 6
With flxOptionen1(5)
    .Rows = 11
    .Cols = 2
    .FixedRows = 1
    .FixedCols = 0
    
    .FormatString = "<AbgabeArt|>Preis"
    
    .ColWidth(0) = TextWidth(String$(30, "X"))
    .ColWidth(1) = TextWidth(String$(10, "9"))
    
    wi% = 0
    For i% = 0 To (.Cols - 1)
        wi% = wi% + .ColWidth(i%)
    Next i%
    .Width = wi% '+ 90
    .Height = .RowHeight(0) * .Rows '+ 90
    
    
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
    
    .SelectionMode = flexSelectionFree
    .GridLines = flexGridFlat
    .BackColor = vbWhite
    
    Call ActProgram.LadeOptionenAbgabeKosten(flxOptionen1(5))
        
    .row = .FixedRows
    .col = 0
    .ColSel = .col
End With
'''''''''''''''''''''''''''''''''''
tabOptionen.Tab = 7
With flxOptionen1(6)
    .Rows = 10
    .Cols = 3
    .FixedRows = 1
    .FixedCols = 2
    
    .FormatString = "<SonderPzn|<Bezeichnung|>Gebühr (brutto)|"
    
    .ColWidth(0) = TextWidth(String$(12, "9"))
    .ColWidth(1) = TextWidth(String$(42, "X"))
    .ColWidth(2) = TextWidth(String$(20, "9"))
    .ColWidth(3) = TextWidth(String$(2, "X"))
    
    wi% = 0
    For i% = 0 To (.Cols - 1)
        wi% = wi% + .ColWidth(i%)
    Next i%
    .Width = wi% '+ 90
    .Height = .RowHeight(0) * .Rows '+ 90
    
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
    
    .SelectionMode = flexSelectionFree
    .GridLines = flexGridFlat
    .BackColor = vbWhite
    .ScrollBars = flexScrollBarVertical
    
    .Rows = .FixedRows
    
    For i = 0 To MaxFiveRxPzns
         h = FiveRxPzns(i)
         ind = InStr(h, ",")
         If (ind > 0) Then
             h2 = Trim(Left(h, ind - 1))
             h = Mid(h, ind + 1)
             ind = InStr(h, ",")
             If (ind > 0) Then
                 PreisStr = Trim(Mid(h, ind + 1))
                 h = Trim(Left(h, ind - 1))
                 .AddItem h2 + vbTab + h + vbTab + PreisStr
             End If
         End If
     Next
    
'    For i% = 0 To 7
'        h = ParenteralPzn(i) + vbTab + ParenteralTxt(i) + vbTab + Format(ParenteralPreis(i), "0.00") + vbTab + Format(ParenteralPreis(i + 8), "0.00")
'        .AddItem h$
'    Next i%
    
    .row = .FixedRows
    .col = .Cols - 1
    .ColSel = .col
End With
'''''''''''''''''''''''''''''''''''


Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

Breite1 = flxOptionen1(4).Left + flxOptionen1(4).Width + 900

Hoehe1% = chkOptionen0(3).Top + chkOptionen0(3).Height
Hoehe2% = flxOptionen1(2).Top + flxOptionen1(2).Height
If (Hoehe2% > Hoehe1%) Then
    Hoehe1% = Hoehe2%
End If

With picStammdatenBack(0)
'    .Width = Breite1 ' lblOptionenZusatz0(1).Left + lblOptionenZusatz0(1).Width + 2 * wpara.LinksX
    .Height = Hoehe1 + 2 * wpara.TitelY 'chkOptionen0(2).Top + chkOptionen0(2).Height + 2 * wpara.TitelY
            
    .BackColor = RGB(232, 217, 172)
    Call wpara.FillGradient(picStammdatenBack(0), 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, RGB(232, 217, 172), RGB(252, 247, 202))
    picStammdatenBack(0).Line (0, .ScaleHeight - (30 * wpara.BildFaktor) * 15)-(.ScaleWidth, .ScaleHeight), .BackColor, BF
End With

For i% = 1 To 6
    With picStammdatenBack(i)
        .Left = picStammdatenBack(0).Left
        .Top = picStammdatenBack(0).Top
'        .Width = picStammdatenBack(0).Width
'        .Height = picStammdatenBack(0).Height
        .Width = 900
        .Height = 900
        
        .BorderStyle = picStammdatenBack(0).BorderStyle
        
        BitBlt .hdc, 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, picStammdatenBack(0).hdc, 0, 0, SRCCOPY
    End With
Next i

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

Call DefErrPop
End Sub

Sub PaintTab(index)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PaintTab")
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
Dim i%, TextHe%
Dim RetVal&, lColor&, bColor&(1)
Dim xStart#, x#, y#
Dim h$
Dim c As Control

'Call DefErrPop: Exit Sub
With picTab(index)
    .Visible = False
    If (index = 0) Then
'        .Left = picStammdatenBack(0).Left + 15
    Else
        .Left = picTab(index - 1).Left + picTab(index - 1).Width '+ 90
    End If
    .Top = picTab(0).Top
    
    .Width = 3000   ' tabOptionen.Width - 3 * wpara.LinksX
    .Height = picTab(0).Height
    
    If (index >= TabsPerRow) Then
        .Top = .Top - .Height
        
        If (index = TabsPerRow) Then
            .Left = picTab(0).Left + 210
        Else
            .Left = picTab(index - 1).Left + picTab(index - 1).Width '+ 90
        End If
    End If
    
    .BorderStyle = 0
    
    .Enabled = False
    If (TabEnabled(0) = 0) Then
        bColor(0) = RGB(150, 150, 150)
        bColor(1) = RGB(165, 165, 165)
    ElseIf (index = AktTab) Then
        bColor(0) = picStammdatenBack(0).BackColor
        bColor(1) = RGB(242, 237, 192)  'bColor(0)
        .Enabled = True
    ElseIf (TabEnabled(index + 1)) Then
        bColor(0) = RGB(199, 176, 123)
        bColor(1) = RGB(214, 191, 138)
        .Enabled = True
    Else
        bColor(0) = RGB(150, 150, 150)
        bColor(1) = RGB(165, 165, 165)
    End If
    
    .BackColor = RGB(200, 200, 200)
    
    .FillStyle = vbSolid
    .FillColor = bColor(0)
    .ForeColor = bColor(0)
    RoundRect .hdc, 0, 0, .Width, .Height, 10, 10
    
    .FillColor = bColor(1)
    .ForeColor = .FillColor
    RoundRect .hdc, 2, 2, .Width, 10, 10, 10
    picTab(index).Line (30, 90)-(.Width, .Height / 2), .FillColor, BF
    
    TextHe = .TextHeight("Äg")
    
    .CurrentX = 90
    .CurrentY = (.Height - TextHe) / 2
    
    If (TabEnabled(0) = 0) Then
        .ForeColor = vbWhite
    ElseIf (index = AktTab) Then
        .ForeColor = RGB(135, 61, 52) ' vbWhite
    Else
        .ForeColor = vbWhite
    End If
    .FillStyle = vbSolid
    .FillColor = .ForeColor
    RoundRect .hdc, .CurrentX / Screen.TwipsPerPixelX, (.CurrentY) / Screen.TwipsPerPixelY, (.CurrentX + TextHe) / Screen.TwipsPerPixelX, (.CurrentY + TextHe) / Screen.TwipsPerPixelY, 5, 5
    
    If (TabEnabled(0) = 0) Then
        .ForeColor = .BackColor
    ElseIf (index = AktTab) Then
        .ForeColor = vbWhite
    Else
        .ForeColor = .BackColor
    End If
    h$ = CStr(index + 1)
    .CurrentX = 90 + (TextHe - .TextWidth(h$)) / 2
    picTab(index).Print h$;
    
    If (TabEnabled(0) = 0) Then
        .ForeColor = vbWhite
    ElseIf (index = AktTab) Then
        .ForeColor = RGB(135, 61, 52) ' vbWhite
    Else
        .ForeColor = vbWhite
    End If
    h = TabNamen(index + 1)
    .CurrentX = 60 + TextHe + 150
    picTab(index).Print h$;
    
    .Width = .CurrentX + 1800

    .ForeColor = RGB(200, 200, 200) ' vbWhite
    
    xStart = .CurrentX + 150
    For i = 0 To 180
        y = -Sin((i + 90) * PI / 180) + 1
        y = y * .ScaleHeight / 2
        x = i * 510# / 180
        picTab(index).Line (xStart + x, 0)-(xStart + x, y), .ForeColor
        SetPixel .hdc, (xStart + x) / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY + 1, bColor(0)
    Next i
    .Width = xStart + 480   ' 600
    .Visible = True
    
    .Refresh
End With

With picStammdatenBack(index)
'    If (IstVorlage) Then
'        Call wpara.FillGradient(picStammdatenBack(0), 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, RGB(232, 217, 172), RGB(236, 215, 53))
'    Else
'        Call wpara.FillGradient(picStammdatenBack(0), 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, RGB(232, 217, 172), RGB(252, 247, 202))
'    End If

    If (index = 0) Then
    ElseIf (index = AktTab) Then
'        .Width = tabOptionen.Width - 30 '- 3 * wpara.LinksX
'        .Height = tabOptionen.Height - (.Top - tabOptionen.Top) - 210
        .Left = picStammdatenBack(0).Left
        .Width = picStammdatenBack(0).Width
        .Height = picStammdatenBack(0).Height

        .BackColor = RGB(232, 217, 172)
        Call wpara.FillGradient(picStammdatenBack(index), 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, RGB(232, 217, 172), RGB(252, 247, 202))
        picStammdatenBack(index).Line (0, .ScaleHeight - (50 * wpara.BildFaktor) * 15)-(.ScaleWidth, .ScaleHeight), .BackColor, BF
'        .BackColor = picStammdatenBack(0).BackColor
'        BitBlt .hdc, 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, picStammdatenBack(0).hdc, 0, 0, SRCCOPY
    Else
        .Width = 900
        .Height = 900
    End If
    
    .TabStop = False
    .Visible = (index = AktTab)
End With


''''''''
If (index = AktTab) Then
    On Error Resume Next
    For Each c In Controls
        If (c.tag <> "0") Then
            If (c.Container.Name = picStammdatenBack(0).Name) Then
                If (c.Container.index = index) Then
                    If (TypeOf c Is TextBox) Or (TypeOf c Is ComboBox) Then
        '                If (TypeOf c Is ComboBox) Then
        '                    Call wpara.ControlBorderless(c)
        '                ElseIf (c.Appearance = 1) Then
        '                    Call wpara.ControlBorderless(c, 2, 2)
        '                Else
        '                    Call wpara.ControlBorderless(c, 1, 1)
        '                End If
        
                        If (c.Enabled) And (c.Locked = 0) Then
                            c.BackColor = vbWhite
                        Else
                            c.BackColor = lblOptionen0(0).BackColor
                        End If
                        
                        If (c.Visible) Then
                            With c.Container
                                .ForeColor = RGB(180, 180, 180) ' vbWhite
                                .FillStyle = vbSolid
                                .FillColor = c.BackColor
                
                                RoundRect .hdc, (c.Left - 60) / Screen.TwipsPerPixelX, (c.Top - 30) / Screen.TwipsPerPixelY, (c.Left + c.Width + 60) / Screen.TwipsPerPixelX, (c.Top + c.Height + 15) / Screen.TwipsPerPixelY, 10, 10
                            End With
                        End If
                    End If
                End If
            End If
        End If
    Next
    picStammdatenBack(index).Refresh
    On Error GoTo DefErr
End If

Call DefErrPop: Exit Sub
'''''''''''''''''''''''''''''''''

picStammdatenBack(index).Visible = (index = AktTab)
If (index = AktTab) Then
    On Error Resume Next
    For Each c In Controls
        If (c.tag <> "0") Then
            If (c.Container.Name = picStammdatenBack(0).Name) Then
                If (c.Container.index = index) Then
                    If (TypeOf c Is TextBox) Or (TypeOf c Is ComboBox) Then
        '                If (TypeOf c Is ComboBox) Then
        '                    Call wpara.ControlBorderless(c)
        '                ElseIf (c.Appearance = 1) Then
        '                    Call wpara.ControlBorderless(c, 2, 2)
        '                Else
        '                    Call wpara.ControlBorderless(c, 1, 1)
        '                End If
        
                        If (c.Enabled) And (c.Locked = 0) Then
                            c.BackColor = vbWhite
                        Else
                            c.BackColor = lblOptionen0(0).BackColor
                        End If
                        
                        If (c.Visible) Then
                            With c.Container
                                .ForeColor = RGB(180, 180, 180) ' vbWhite
                                .FillStyle = vbSolid
                                .FillColor = c.BackColor
                
                                RoundRect .hdc, (c.Left - 60) / Screen.TwipsPerPixelX, (c.Top - 30) / Screen.TwipsPerPixelY, (c.Left + c.Width + 60) / Screen.TwipsPerPixelX, (c.Top + c.Height + 15) / Screen.TwipsPerPixelY, 10, 10
                            End With
                        End If
                    End If
                End If
            End If
        End If
    Next
    picStammdatenBack(index).Refresh
    On Error GoTo DefErr
End If

'tabOptionen.Visible = False

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

Private Sub nlcmdOk_Click()
Call cmdOk_Click
End Sub

Private Sub nlcmdEsc_Click()
Call cmdEsc_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If (para.Newline) Then
    If (KeyAscii = 13) Then
        Call nlcmdOk_Click
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
