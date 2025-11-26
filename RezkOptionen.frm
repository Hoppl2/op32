VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRezkOptionen 
   Caption         =   "Optionen"
   ClientHeight    =   7395
   ClientLeft      =   540
   ClientTop       =   735
   ClientWidth     =   13245
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   13245
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2400
      TabIndex        =   13
      Top             =   6840
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   720
      TabIndex        =   12
      Top             =   6840
      Width           =   1200
   End
   Begin TabDlg.SSTab tabOptionen 
      Height          =   6285
      Left            =   240
      TabIndex        =   14
      Top             =   360
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   11086
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   6
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
      TabPicture(0)   =   "RezkOptionen.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fmeOptionen(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&2 - A+V Taxierung"
      TabPicture(1)   =   "RezkOptionen.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fmeOptionen(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&3 - Tätigkeiten"
      TabPicture(2)   =   "RezkOptionen.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fmeOptionen(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&4 - Abrechnungsdaten"
      TabPicture(3)   =   "RezkOptionen.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fmeOptionen(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "&5 - Sonderbelege"
      TabPicture(4)   =   "RezkOptionen.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fmeOptionen(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "&6 - Parenteral"
      TabPicture(5)   =   "RezkOptionen.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fmeOptionen(5)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "&7 - Gesamt-Brutto"
      TabPicture(6)   =   "RezkOptionen.frx":00A8
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "fmeOptionen(6)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      Begin VB.Frame fmeOptionen 
         Height          =   4215
         Index           =   6
         Left            =   480
         TabIndex        =   30
         Top             =   1200
         Width           =   8895
         Begin MSFlexGridLib.MSFlexGrid flxOptionen1 
            Height          =   2700
            Index           =   5
            Left            =   840
            TabIndex        =   31
            Top             =   840
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   4763
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
      Begin VB.Frame fmeOptionen 
         Height          =   4215
         Index           =   5
         Left            =   -74400
         TabIndex        =   24
         Top             =   960
         Width           =   8895
         Begin VB.TextBox txtOptionen5 
            Height          =   495
            Index           =   1
            Left            =   5640
            TabIndex        =   29
            Text            =   "999,999"
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox txtOptionen5 
            Height          =   495
            Index           =   0
            Left            =   5640
            TabIndex        =   27
            Text            =   "999,999"
            Top             =   1920
            Width           =   975
         End
         Begin MSFlexGridLib.MSFlexGrid flxOptionen1 
            Height          =   780
            Index           =   4
            Left            =   840
            TabIndex        =   25
            Top             =   840
            Width           =   5520
            _ExtentX        =   9737
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
            Caption         =   "Aufschlag &Gefässe  (%)"
            Height          =   375
            Index           =   1
            Left            =   600
            TabIndex        =   28
            Top             =   2640
            Width           =   3615
         End
         Begin VB.Label lblOptionen5 
            Caption         =   "Aufschlag &Spezialitäten  (%)"
            Height          =   375
            Index           =   0
            Left            =   600
            TabIndex        =   26
            Top             =   1920
            Width           =   3615
         End
      End
      Begin VB.Frame fmeOptionen 
         Height          =   4215
         Index           =   4
         Left            =   -75000
         TabIndex        =   22
         Top             =   720
         Width           =   8895
         Begin MSFlexGridLib.MSFlexGrid flxOptionen1 
            Height          =   2700
            Index           =   3
            Left            =   840
            TabIndex        =   23
            Top             =   840
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   4763
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
      Begin VB.Frame fmeOptionen 
         Height          =   4215
         Index           =   3
         Left            =   -74880
         TabIndex        =   20
         Top             =   480
         Width           =   8895
         Begin MSFlexGridLib.MSFlexGrid flxOptionen1 
            Height          =   2700
            Index           =   2
            Left            =   840
            TabIndex        =   21
            Top             =   840
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   4763
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
      Begin VB.Frame fmeOptionen 
         Height          =   4215
         Index           =   2
         Left            =   -75000
         TabIndex        =   18
         Top             =   840
         Width           =   8895
         Begin MSFlexGridLib.MSFlexGrid flxOptionen1 
            Height          =   2700
            Index           =   1
            Left            =   840
            TabIndex        =   19
            Top             =   840
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   4763
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
      Begin VB.Frame fmeOptionen 
         Height          =   4215
         Index           =   1
         Left            =   -74520
         TabIndex        =   16
         Top             =   480
         Width           =   8895
         Begin MSFlexGridLib.MSFlexGrid flxOptionen1 
            Height          =   2700
            Index           =   0
            Left            =   840
            TabIndex        =   17
            Top             =   840
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   4763
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
      Begin VB.Frame fmeOptionen 
         Height          =   5655
         Index           =   0
         Left            =   -74280
         TabIndex        =   15
         Top             =   480
         Width           =   8895
         Begin VB.CheckBox chkOptionen0 
            Caption         =   "Teilnahme an @Rezept von &AVP"
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   11
            Top             =   4680
            Width           =   6015
         End
         Begin VB.CheckBox chkOptionen0 
            Caption         =   "Rezeptur auf Rezept &drucken"
            Height          =   495
            Index           =   2
            Left            =   240
            TabIndex        =   10
            Top             =   3960
            Width           =   6015
         End
         Begin VB.TextBox txtOptionen0 
            Height          =   495
            Index           =   2
            Left            =   5280
            MaxLength       =   50
            TabIndex        =   5
            Text            =   "WWW9999"
            Top             =   1680
            Width           =   1575
         End
         Begin VB.CheckBox chkOptionen0 
            Caption         =   "&BTM-Gebühr als Rezeptzeile"
            Height          =   495
            Index           =   1
            Left            =   360
            TabIndex        =   9
            Top             =   3480
            Width           =   6015
         End
         Begin VB.TextBox txtOptionen0 
            Height          =   495
            Index           =   1
            Left            =   5280
            MaxLength       =   38
            TabIndex        =   3
            Text            =   "WWW9999"
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CheckBox chkOptionen0 
            Caption         =   "&Rezepturen mit Kassenrabatt taxieren"
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   2880
            Width           =   6015
         End
         Begin VB.TextBox txtOptionen0 
            Height          =   495
            Index           =   3
            Left            =   5280
            TabIndex        =   7
            Text            =   "999,999"
            Top             =   2280
            Width           =   975
         End
         Begin VB.TextBox txtOptionen0 
            Height          =   495
            Index           =   0
            Left            =   5280
            MaxLength       =   7
            TabIndex        =   1
            Text            =   "WWW9999"
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblOptionen0 
            Caption         =   "Rezept-Text &BTM"
            Height          =   375
            Index           =   2
            Left            =   480
            TabIndex        =   4
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label lblOptionen0 
            Caption         =   "Rezept-&Text"
            Height          =   375
            Index           =   1
            Left            =   480
            TabIndex        =   2
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label lblOptionen0 
            Caption         =   "&Kassen-Rabatt  (%)"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   6
            Top             =   2280
            Width           =   3615
         End
         Begin VB.Label lblOptionen0 
            Caption         =   "&Instituts-Kennzeichen"
            Height          =   375
            Index           =   0
            Left            =   480
            TabIndex        =   0
            Top             =   480
            Width           =   2175
         End
      End
   End
End
Attribute VB_Name = "frmRezkOptionen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OrgRezApoDruckName$, OrgBtmRezDruckName$


Private Const DefErrModul = "REZKOPTIONEN.FRM"

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
    .Left = tabOptionen.Left + fmeOptionen(3).Left + flxOptionen1(2).Left + flxOptionen1(2).ColPos(1) + 45
    .Left = .Left + Me.Left + wpara.FrmBorderHeight
    .Top = tabOptionen.Top + fmeOptionen(3).Top + flxOptionen1(2).Top + (flxOptionen1(2).row * flxOptionen1(2).RowHeight(0))
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

If (ActiveControl.Name = cmdOk.Name) Then
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
Dim h$, h2$, h3$, FormStr$
Dim c As Object


Call wpara.InitFont(Me)

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

lblOptionen0(0).Left = wpara.LinksX
lblOptionen0(0).Top = 2 * wpara.TitelY
txtOptionen0(0).Left = lblOptionen0(1).Left + lblOptionen0(1).Width + 300
txtOptionen0(0).Top = lblOptionen0(0).Top + (lblOptionen0(0).Height - txtOptionen0(0).Height) / 2

For i% = 1 To 3
    lblOptionen0(i%).Left = lblOptionen0(0).Left
    lblOptionen0(i%).Top = lblOptionen0(i% - 1).Top + lblOptionen0(i% - 1).Height + 300
    txtOptionen0(i%).Left = txtOptionen0(0).Left
    txtOptionen0(i%).Top = lblOptionen0(i%).Top + (lblOptionen0(i%).Height - txtOptionen0(i%).Height) / 2
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

    .Width = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + 90
    .Height = .RowHeight(0) * 11 + 90
    
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
    
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

    .Width = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + 90
    .Height = .RowHeight(0) * 11 + 90
    
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
    
    .SelectionMode = flexSelectionFree
    
    h$ = "Rezeptspeicher" + vbTab
    
    For i% = 0 To (AnzTaetigkeiten% - 1)
'        h$ = RTrim$(Taetigkeiten(i%).Taetigkeit)
'        h$ = h$ + vbTab
        h2$ = ""
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
        h$ = h$ + h3$ + vbTab + h2$
    Next i%
    
    .AddItem h$
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
    
    .Width = .ColWidth(0) + .ColWidth(1) + 90
    .Height = .RowHeight(0) * 13 + 90
    
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
    .Width = wi% + 90
    .Height = .RowHeight(0) * .Rows + 90
    
    
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
    .Width = wi% + 90
    .Height = .RowHeight(0) * .Rows + 90
    
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
    txtOptionen5(0).Top = .Top + (.Height - txtOptionen0(0).Height) / 2
End With
For i% = 1 To 1
    lblOptionen5(i%).Left = lblOptionen5(0).Left
    lblOptionen5(i%).Top = lblOptionen5(i% - 1).Top + lblOptionen0(i% - 1).Height + 150
    txtOptionen5(i%).Left = txtOptionen5(0).Left
    txtOptionen5(i%).Top = lblOptionen5(i%).Top + (lblOptionen5(i%).Height - txtOptionen5(i%).Height) / 2
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
    .Width = wi% + 90
    .Height = .RowHeight(0) * .Rows + 90
    
    
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


Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

With fmeOptionen(0)
    .Left = wpara.LinksX
    .Top = 3 * wpara.TitelY
'    .Width = txtOptionen0(1).Left + txtOptionen0(1).Width + 900
'    .Width = flxOptionen1(1).Left + flxOptionen1(1).Width + 1800
    .Width = flxOptionen1(4).Left + flxOptionen1(4).Width + 900
    
    Hoehe1% = chkOptionen0(3).Top + chkOptionen0(3).Height
    Hoehe2% = flxOptionen1(2).Top + flxOptionen1(2).Height
    If (Hoehe2% > Hoehe1%) Then
        Hoehe1% = Hoehe2%
    End If
    .Height = Hoehe1% + 300
End With
For i% = 1 To 6
    With fmeOptionen(i%)
        .Left = fmeOptionen(0).Left
        .Top = fmeOptionen(0).Top
        .Width = fmeOptionen(0).Width
        .Height = fmeOptionen(0).Height
    End With
Next i%

With tabOptionen
    .Left = wpara.LinksX
    .Top = wpara.TitelY
    .Width = fmeOptionen(0).Left + fmeOptionen(0).Width + wpara.LinksX
    .Height = fmeOptionen(0).Top + fmeOptionen(0).Height + wpara.TitelY
End With



cmdOk.Top = tabOptionen.Top + tabOptionen.Height + 150
cmdEsc.Top = cmdOk.Top

Me.Width = tabOptionen.Width + 2 * wpara.LinksX

cmdOk.Width = wpara.ButtonX
cmdOk.Height = wpara.ButtonY
cmdEsc.Width = wpara.ButtonX
cmdEsc.Height = wpara.ButtonY
cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

Breite1% = frmAction.Left + (frmAction.Width - Me.Width) / 2
If (Breite1% < 0) Then Breite1% = 0
Me.Left = Breite1%
Hoehe1% = frmAction.Top + (frmAction.Height - Me.Height) / 2
If (Hoehe1% < 0) Then Hoehe1% = 0
Me.Top = Hoehe1%

tabOptionen.Tab = 0
Call TabDisable
Call TabEnable(tabOptionen.Tab)

Call DefErrPop
End Sub

Private Sub tabOptionen_Click(PreviousTab As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("tabOptionen_Click")
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
If (tabOptionen.Visible = False) Then Call DefErrPop: Exit Sub

Call TabDisable
Call TabEnable(tabOptionen.Tab)

Call DefErrPop
End Sub

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

For i% = 0 To 6
    fmeOptionen(i%).Visible = False
Next i%

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

fmeOptionen(hTab%).Visible = True

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
End If

Call DefErrPop
End Sub

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
row% = 1
col% = 1
                
With flxOptionen1(1)
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
        .Left = tabOptionen.Left + fmeOptionen(2).Left + flxOptionen1(1).Left + flxOptionen1(1).ColPos(col%) + 45
        .Left = .Left + Me.Left + wpara.FrmBorderHeight
        .Top = tabOptionen.Top + fmeOptionen(2).Top + flxOptionen1(1).Top + flxOptionen1(1).RowHeight(0)
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
        If (h$ <> "") And (RTrim$(.TextMatrix(i%, 1)) <> "") Then
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

i% = 1
If (i% <= AnzTaetigkeiten) Then
    h$ = RTrim$(Taetigkeiten(i% - 1).Taetigkeit)
    For j% = 0 To 79
        If (Taetigkeiten(i% - 1).pers(j%) > 0) Then
            h$ = h$ + "," + Mid$(Str$(Taetigkeiten(i% - 1).pers(j%)), 2)
        Else
            Exit For
        End If
    Next j%
Else
    h$ = ""
End If

Key$ = "Taetigkeit" + Format(i%, "00")
l& = WritePrivateProfileString("Rezeptkontrolle", Key$, h$, INI_DATEI)


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
    .Left = tabOptionen.Left + fmeOptionen(FlexInd + 1).Left + flxOptionen1(FlexInd).Left + flxOptionen1(FlexInd).ColPos(EditCol%) + 45
    .Left = .Left + Me.Left + wpara.FrmBorderHeight
    .Top = tabOptionen.Top + fmeOptionen(FlexInd + 1).Top + flxOptionen1(FlexInd).Top + flxOptionen1(FlexInd).RowHeight(0) * EditRow%
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
            .text = Format(xVal(EditTxt$), "0.00")
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


