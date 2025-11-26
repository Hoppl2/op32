VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlbutton.ocx"
Begin VB.Form frmAngebote 
   Caption         =   "GH-Angebote"
   ClientHeight    =   5850
   ClientLeft      =   1395
   ClientTop       =   1020
   ClientWidth     =   9120
   Icon            =   "Angebote.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9120
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   8760
      Picture         =   "Angebote.frx":014A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   8520
      Picture         =   "Angebote.frx":0203
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   8280
      Picture         =   "Angebote.frx":02B7
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin nlCommandButton.nlCommand nlcmdAngebotEdit 
      Height          =   495
      Left            =   6480
      TabIndex        =   20
      Top             =   3720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdAngebotLoeschen 
      Height          =   495
      Left            =   6360
      TabIndex        =   19
      Top             =   2880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdAngebotNeu 
      Height          =   495
      Left            =   6360
      TabIndex        =   18
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdF5 
      Height          =   495
      Left            =   6360
      TabIndex        =   17
      Top             =   1560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdF2 
      Height          =   495
      Left            =   6360
      TabIndex        =   16
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   495
      Left            =   4200
      TabIndex        =   15
      Top             =   4320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   495
      Left            =   2520
      TabIndex        =   14
      Top             =   4320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin VB.CommandButton cmdAngebotEdit 
      Caption         =   "E&dit"
      Height          =   450
      Left            =   4920
      TabIndex        =   5
      Top             =   3720
      Width           =   1200
   End
   Begin VB.CommandButton cmdAngebotLoeschen 
      Caption         =   "&Entfernen"
      Height          =   450
      Left            =   4920
      TabIndex        =   4
      Top             =   3000
      Width           =   1200
   End
   Begin VB.CommandButton cmdAngebotNeu 
      Caption         =   "&Neu anlegen"
      Height          =   450
      Left            =   4920
      TabIndex        =   3
      Top             =   2280
      Width           =   1200
   End
   Begin VB.CommandButton cmdF2 
      Caption         =   "&Temporär (F2)"
      Height          =   450
      Left            =   4800
      TabIndex        =   1
      Top             =   960
      Width           =   1200
   End
   Begin VB.CommandButton cmdF5 
      Caption         =   "&Ablehnen (F5)"
      Height          =   450
      Left            =   4800
      TabIndex        =   2
      Top             =   1560
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "ESC"
      Height          =   450
      Left            =   3600
      TabIndex        =   7
      Top             =   3480
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Zuordnen"
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   3480
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid flxAngebote 
      Height          =   2280
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4022
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483633
      BackColorBkg    =   -2147483633
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLines       =   0
      ScrollBars      =   1
      SelectionMode   =   1
   End
   Begin VB.Label lblAngeboteWert 
      BackStyle       =   0  'Transparent
      Caption         =   "999999.99"
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
      Index           =   2
      Left            =   5040
      TabIndex        =   13
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblAngebote 
      BackStyle       =   0  'Transparent
      Caption         =   "TaxeAEP"
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
      Left            =   4320
      TabIndex        =   12
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblAngeboteWert 
      BackStyle       =   0  'Transparent
      Caption         =   "999999.99"
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
      Index           =   1
      Left            =   3000
      TabIndex        =   11
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblAngebote 
      BackStyle       =   0  'Transparent
      Caption         =   "AEP"
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
      Left            =   2280
      TabIndex        =   10
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblAngeboteWert 
      BackStyle       =   0  'Transparent
      Caption         =   "999999"
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
      Index           =   0
      Left            =   960
      TabIndex        =   9
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblAngebote 
      BackStyle       =   0  'Transparent
      Caption         =   "BMopt"
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
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmAngebote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "ANGEBOTE.FRM"

Dim NoChange%

Dim F5Gedrueckt%

Private Sub cmdEsc_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdEsc_Click")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

If (AngebotDirektEingabe% Or F5Gedrueckt%) Then
    AngebotInd% = 0
Else
    AngebotInd% = -1
End If

AngebotTemporaer = 0
Unload Me

Call clsError.DefErrPop
End Sub

Private Sub cmdF2_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdF2_Click")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim iAng%, iTemp As Byte

iAng% = flxAngebote.col - 1

If (LocalAngebotRec(iAng%).LaufNr = 2000) Then
    AngebotInd% = flxAngebote.col
    AngebotTemporaer = 1
    Unload Me
    Call clsError.DefErrPop: Exit Sub
End If


flxAngebote.HighLight = flexHighlightNever
cmdOk.Visible = False
cmdEsc.Visible = False
cmdF2.Visible = False
cmdF5.Visible = False
cmdAngebotNeu.Visible = False
cmdAngebotLoeschen.Visible = False
cmdAngebotEdit.Visible = False

nlcmdOk.Visible = False
nlcmdEsc.Visible = False
nlcmdF2.Visible = False
nlcmdF5.Visible = False
nlcmdAngebotNeu.Visible = False
nlcmdAngebotLoeschen.Visible = False
nlcmdAngebotEdit.Visible = False


AngebotBm% = LocalAngebotRec(iAng%).bmOrg
If (LocalAngebotRec(iAng%).st <> "P") Then
    AngebotNm% = LocalAngebotRec(iAng%).mpOrg
Else
    AngebotNm% = 0
End If

AngebotNeuLief% = LocalAngebotRec(iAng%).gh
AngebotGhBest% = LocalAngebotRec(iAng%).ghBest
iTemp = AngebotTemporaer
AngebotTemporaer = 1

If (AnzLocalAngebote% = 0) Then
    AngebotMitZr = 1
Else
    AngebotMitZr = LocalAngebotRec(iAng%).IstManuell
End If
AngebotNeu% = True

frmEditAngebot.Show 1

If (EditErg%) Then
    AngebotInd% = AnzLocalAngebote%
    LocalAngebotRec(AnzLocalAngebote% - 1).IstManuell = LocalAngebotRec(iAng%).IstManuell
    LocalAngebotRec(AnzLocalAngebote% - 1).LaufNr = LocalAngebotRec(iAng%).LaufNr
    If (LocalAngebotRec(AnzLocalAngebote% - 1).bm < LocalAngebotRec(iAng%).bm) Then LocalAngebotRec(AnzLocalAngebote% - 1).LaufNr = 200
Else
    AngebotInd% = -1
    AngebotTemporaer = 0
End If
Unload Me

''AngebotInd% = flxAngebote.col + 100
'AngebotInd% = flxAngebote.col
'AngebotTemporaer = 1
'Unload Me

Call clsError.DefErrPop
End Sub

Private Sub cmdF5_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdF5_Click")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, aCol%


AngebotInd% = 0
AngebotTemporaer = 0

With flxAngebote
    .Redraw = False
    aCol% = .col
    .row = 0
    For i% = 1 To (.Cols - 1)
        .col = i%
        If (.CellBackColor = vbWhite) Then
            .CellBackColor = .BackColor
            Exit For
        End If
    Next i%
    .row = 1
    .col = aCol%
    .RowSel = .Rows - 1
    .Redraw = True
End With

F5Gedrueckt% = True
'Unload Me

Call clsError.DefErrPop
End Sub

Private Sub cmdAngebotNeu_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdAngebotNeu_Click")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
flxAngebote.HighLight = flexHighlightNever
cmdOk.Visible = False
cmdEsc.Visible = False
cmdF2.Visible = False
cmdF5.Visible = False
cmdAngebotNeu.Visible = False
cmdAngebotLoeschen.Visible = False
cmdAngebotEdit.Visible = False

nlcmdOk.Visible = False
nlcmdEsc.Visible = False
nlcmdF2.Visible = False
nlcmdF5.Visible = False
nlcmdAngebotNeu.Visible = False
nlcmdAngebotLoeschen.Visible = False
nlcmdAngebotEdit.Visible = False

AngebotNeuLief% = AngebotActLief%
AngebotGhBest% = AngebotActLief%
AngebotMitZr = 1
AngebotNeu% = True
AngebotEditBm% = AngebotBm%
AngebotEditNm% = AngebotNm%
AngebotEditZr! = -123.45!   '0!
frmEditAngebot.Show 1

If (EditErg%) Then
    AngebotInd% = AnzLocalAngebote%
Else
    AngebotInd% = -1
End If
AngebotTemporaer = 0
Unload Me

'flxAngebote.HighLight = flexHighlightAlways
'cmdOk.Visible = True
'cmdEsc.Visible = True
'cmdF2.Visible = True
'cmdF5.Visible = True
'If (EditErg%) Then Call Form_Load

Call clsError.DefErrPop
End Sub

Private Sub cmdAngebotEdit_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdAngebotEdit_Click")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim iAng%, OrgF5Val%

iAng% = flxAngebote.col - 1

OrgF5Val% = cmdF5.Visible
flxAngebote.HighLight = flexHighlightNever
cmdOk.Visible = False
cmdEsc.Visible = False
cmdF2.Visible = False
cmdF5.Visible = False
cmdAngebotNeu.Visible = False
cmdAngebotLoeschen.Visible = False
cmdAngebotEdit.Visible = False
nlcmdOk.Visible = False
nlcmdEsc.Visible = False
nlcmdF2.Visible = False
nlcmdF5.Visible = False
nlcmdAngebotNeu.Visible = False
nlcmdAngebotLoeschen.Visible = False
nlcmdAngebotEdit.Visible = False

AngebotNeuLief% = LocalAngebotRec(iAng%).gh
AngebotGhBest% = LocalAngebotRec(iAng%).ghBest
AngebotMitZr = 1
AngebotNeu% = False
AngebotEditBm% = LocalAngebotRec(iAng%).bm
AngebotEditNm% = LocalAngebotRec(iAng%).mp
AngebotEditZr! = LocalAngebotRec(iAng%).zr
AngebotEditRecno& = LocalAngebotRec(iAng%).recno
frmEditAngebot.Show 1

If (EditErg%) Then Call Form_Load
If (iNewLine) Then
    nlcmdOk.Visible = True
    nlcmdEsc.Visible = True
    nlcmdF5.Visible = OrgF5Val%
    nlcmdAngebotNeu.Visible = True
Else
    cmdOk.Visible = True
    cmdEsc.Visible = True
    cmdF5.Visible = OrgF5Val%
    cmdAngebotNeu.Visible = True
End If
Call flxAngebote_RowColChange
With flxAngebote
    .row = 1
    .RowSel = .Rows - 1
    .HighLight = flexHighlightAlways
End With

Call clsError.DefErrPop
End Sub

Private Sub cmdAngebotLoeschen_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdAngebotLoeschen_Click")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim FabsErrf%
Dim FabsRecno&, lRecs&
                        
With flxAngebote
    FabsRecno& = Val(.TextMatrix(18, .col))
End With

If (AngeboteDbOk) Then
    SQLStr = "DELETE * FROM Angebote WHERE Id=" + CStr(FabsRecno)
    ManuellAngeboteDB1.ActiveConn.CommandTimeout = 120
    Call ManuellAngeboteDB1.ActiveConn.Execute(SQLStr, lRecs, adExecuteNoRecords)
Else
    With clsManuelleAngebote1
    '    FabsErrf% = .IndexDelete(0, FabsRecno&, AngebotPzn$, FabsRecno&)
    '    If (FabsErrf% = 0) Then
            .GetRecord (FabsRecno& + 1)
            .st = "X"
            .PutRecord (FabsRecno& + 1)
    '    End If
    End With
End If
Call Form_Load
flxAngebote.SetFocus

Call clsError.DefErrPop
End Sub

Private Sub cmdOk_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("cmdOk_Click")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim col%

col% = flxAngebote.col
'If ((AngebotInd% >= 100) And (col% = 1)) Then
If ((AngebotTemporaer) And (col% = 1)) Then
    AngebotInd% = -1
Else
    AngebotInd% = col%
End If
AngebotTemporaer = 0
Unload Me

Call clsError.DefErrPop
End Sub

Private Sub flxAngebote_KeyDown(KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("flxAngebote_KeyDown")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%

If (iNewLine) Then
    If (KeyCode = vbKeyF2) And (nlcmdF2.Enabled) Then
        nlcmdF2.Value = True
    ElseIf (KeyCode = vbKeyF5) And (nlcmdF5.Enabled) Then
        nlcmdF5.Value = True
    End If
Else
    If (KeyCode = vbKeyF2) And (cmdF2.Enabled) Then
        cmdF2.Value = True
    ElseIf (KeyCode = vbKeyF5) And (cmdF5.Enabled) Then
        cmdF5.Value = True
    End If
End If

Call clsError.DefErrPop
End Sub

Private Sub flxAngebote_RowColChange()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("flxAngebote_RowColChange")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, AltRow%, iLaufNr%

If (NoChange%) Then Call clsError.DefErrPop: Exit Sub
NoChange% = True

If (iNewLine) Then
    With flxAngebote
        If (Val(.TextMatrix(16, .col)) > 0) Then
            nlcmdAngebotLoeschen.Visible = True
            nlcmdAngebotEdit.Visible = True
        Else
            nlcmdAngebotLoeschen.Visible = False
            nlcmdAngebotEdit.Visible = False
        End If
            
        AltRow% = .row
        .row = 0
        If (.CellBackColor = vbWhite) Then
            nlcmdAngebotLoeschen.Enabled = False
            nlcmdAngebotEdit.Enabled = False
        Else
            nlcmdAngebotLoeschen.Enabled = True
            nlcmdAngebotEdit.Enabled = True
        End If
        .row = AltRow%
        
        iLaufNr% = Val(.TextMatrix(17, .col))
        If ((AngebotTemporaer) And (.col = 1)) Or ((iLaufNr% > 1000) And (iLaufNr% <> 2000)) Then
            nlcmdF2.Enabled = False
        Else
            nlcmdF2.Enabled = True
        End If
    End With
    cmdF2.Enabled = nlcmdF2.Enabled
    cmdF5.Enabled = nlcmdF5.Enabled
    cmdAngebotNeu.Enabled = nlcmdAngebotNeu.Enabled
    cmdAngebotLoeschen.Enabled = nlcmdAngebotLoeschen.Enabled
    cmdAngebotEdit.Enabled = nlcmdAngebotEdit.Enabled
Else
    With flxAngebote
        If (Val(.TextMatrix(16, .col)) > 0) Then
            cmdAngebotLoeschen.Visible = True
            cmdAngebotEdit.Visible = True
        Else
            cmdAngebotLoeschen.Visible = False
            cmdAngebotEdit.Visible = False
        End If
            
        AltRow% = .row
        .row = 0
        If (.CellBackColor = vbWhite) Then
            cmdAngebotLoeschen.Enabled = False
            cmdAngebotEdit.Enabled = False
        Else
            cmdAngebotLoeschen.Enabled = True
            cmdAngebotEdit.Enabled = True
        End If
        .row = AltRow%
        
        iLaufNr% = Val(.TextMatrix(17, .col))
        If ((AngebotTemporaer) And (.col = 1)) Or ((iLaufNr% > 1000) And (iLaufNr% <> 2000)) Then
            cmdF2.Enabled = False
        Else
            cmdF2.Enabled = True
        End If
    End With
End If
NoChange% = False

Call clsError.DefErrPop
End Sub

Private Sub Form_Load()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_Load")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%, iAdd%, iAdd2%, x%, y%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%, maxSp%
Dim MenuHeight&, ScrollHeight&
Dim h$, h2$, FormStr$

Call wPara1.InitFont(Me)

Font.Bold = False   ' True

For i% = 0 To 2
    lblAngebote(i%).Top = wPara1.TitelY
    lblAngeboteWert(i%).Caption = ""
    lblAngeboteWert(i%).Top = wPara1.TitelY
Next i%

lblAngebote(0).Left = wPara1.LinksX
lblAngeboteWert(0).Left = lblAngebote(0).Left + lblAngebote(0).Width + 60
lblAngebote(1).Left = lblAngeboteWert(0).Left + lblAngeboteWert(0).Width + 180
lblAngeboteWert(1).Left = lblAngebote(1).Left + lblAngebote(1).Width + 60
lblAngebote(2).Left = lblAngeboteWert(1).Left + lblAngeboteWert(1).Width + 180
lblAngeboteWert(2).Left = lblAngebote(2).Left + lblAngebote(2).Width + 60



With flxAngebote
    .Cols = 2
    .Rows = 15
    .FixedRows = 1
    .FixedCols = 1
    
    .Top = lblAngebote(0).Top + lblAngebote(0).Height + 300
    .Left = wPara1.LinksX
    .Height = .RowHeight(0) * 15 + 90   '16
    
    FormStr$ = ""
    For i% = 1 To 30
        FormStr$ = FormStr$ + "|^" + Mid$(Str$(i%), 2)
    Next i%
    FormStr$ = FormStr$ + ";|Lieferant|Original-Angebot|Angebot|AEP|Angebots-Rabatt|Angebots-Preis|Rabatt|Preis|"
    FormStr$ = FormStr$ + "NR (zu" + Str$(Para1.FakNatu * 100) + "%)|"
    FormStr$ = FormStr$ + "aliquote Lager-Kosten|aliquote Bestell-Kosten|Staffel|kalkulierter AEP|Gespart|%|||"
'    FormStr$ = FormStr$ + ";|Lieferant|Orig.Angebot|Angebot|AEP|Ang.Rabatt|Ang.Preis|Rabatt|Preis|NR|aliq.Lager|aliq.Best|Staffel|kalkAEP|Gespart|%|||"
    .FormatString = FormStr$
    .SelectionMode = flexSelectionByColumn
    
    .RowHeight(12) = 0
    
    .FillStyle = flexFillRepeat
    
    If (iNewLine = 0) Then
        .row = 13
        .col = 0
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1
        .CellBackColor = vbWhite
    End If
    
    .row = 2
    .col = 0
    .RowSel = 3
    .ColSel = .Cols - 1
    .CellFontBold = True
    
    .row = 6
    .col = 0
    .RowSel = .row
    .ColSel = .Cols - 1
    .CellFontBold = True
    
    .row = 8
    .col = 0
    .RowSel = 9
    .ColSel = .Cols - 1
    .CellFontBold = True
    
    .FillStyle = flexFillSingle
    
    If (IstBewertung% = False) Then
        Call Angebote
    Else
        .Cols = 2
        For i% = 0 To 2
            lblAngebote(i%).Visible = False
            lblAngeboteWert(i%).Visible = False
        Next i%
    End If



    For i% = 0 To .Cols - 1
        .ColWidth(i%) = TextWidth(String(9, "W"))
        .ColAlignment(i%) = flexAlignRightCenter
    Next i%
    .ColWidth(0) = TextWidth(String(13, "W"))

'    maxSp% = (frmAction.ScaleWidth - (2 * wPara1.LinksX) - 900) \ .ColWidth(0)
    maxSp% = (Screen.Width - (2 * wPara1.LinksX) - (TextWidth(cmdF2.Caption) + 150 + 150) - 900) \ .ColWidth(0)
    maxSp% = maxSp% - 1
    If (iNewLine) Then
        maxSp = 6
    End If
    If (.Cols <= maxSp%) Then
        maxSp% = .Cols
    Else
'        .Height = .Height + .RowHeight(0)
        .Height = .Height + wPara1.FrmScrollHeight
    End If
    
    spBreite% = 0
    For i% = 0 To maxSp% - 1
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90    '90
    
    Do
        If (.col >= (.LeftCol + maxSp% - 1)) Then
            .LeftCol = .LeftCol + 1
        Else
            Exit Do
        End If
    Loop
        
    If (.Cols > maxSp%) Then
        .ScrollBars = flexScrollBarHorizontal
    Else
        .ScrollBars = flexScrollBarNone
    End If
End With

'

Font.Name = wPara1.FontName(1)
Font.Size = wPara1.FontSize(1)


With cmdOk
    .Width = wPara1.ButtonX
    .Height = wPara1.ButtonY
End With
With cmdEsc
    .Width = cmdOk.Width
    .Height = cmdOk.Height
End With
With cmdF2
    .Width = TextWidth(cmdF2.Caption) + 150
    .Height = cmdOk.Height
End With
With cmdF5
    .Width = cmdF2.Width
    .Height = cmdOk.Height
End With
With cmdAngebotNeu
    .Width = cmdF2.Width
    .Height = cmdOk.Height
End With
With cmdAngebotLoeschen
    .Width = cmdF2.Width
    .Height = cmdOk.Height
End With
With cmdAngebotEdit
    .Width = cmdF2.Width
    .Height = cmdOk.Height
End With

cmdOk.Top = flxAngebote.Top + flxAngebote.Height + 150
cmdEsc.Top = cmdOk.Top

cmdF2.Top = flxAngebote.Top
cmdF2.Left = flxAngebote.Left + flxAngebote.Width + 150
cmdF5.Top = cmdF2.Top + cmdF2.Height + 270
cmdF5.Left = cmdF2.Left

cmdAngebotNeu.Top = cmdF5.Top + cmdF5.Height + 270
cmdAngebotNeu.Left = cmdF5.Left
cmdAngebotLoeschen.Top = cmdAngebotNeu.Top + cmdAngebotNeu.Height + 270
cmdAngebotLoeschen.Left = cmdF5.Left
cmdAngebotEdit.Top = cmdAngebotLoeschen.Top + cmdAngebotLoeschen.Height + 270
cmdAngebotEdit.Left = cmdF5.Left
    
Breite1% = flxAngebote.Width + 2 * wPara1.LinksX
If (AngebotModus% = 1) Then
'    Breite2% = cmdF5.Width + 900 + cmdOk.Width + 300 + cmdEsc.Width + 2 * wPara1.LinksX
    Breite2% = cmdF5.Left + cmdF5.Width + 2 * wPara1.LinksX
Else
    Breite2% = 0
End If

If (Breite2% > Breite1%) Then
    Me.Width = Breite2%
Else
    Me.Width = Breite1%
End If


If (AngebotModus% = 1) Then
    cmdOk.Caption = "Übernahme"    '"Binden"
    cmdOk.default = True
    cmdOk.Cancel = False
    cmdOk.Visible = True
    cmdEsc.Cancel = True
    cmdEsc.Visible = True
'    cmdF5.Visible = True
    cmdF5.Visible = (AngebotInd% >= 0)
    cmdAngebotNeu.Visible = (AngebotActLief% >= 0)
    
'    cmdF5.Left = (Me.ScaleWidth - (cmdF5.Width + 900 + cmdOk.Width + 300 + cmdEsc.Width)) / 2
'    cmdOk.Left = cmdF5.Left + cmdF5.Width + 900
'    cmdEsc.Left = cmdOk.Left + cmdOk.Width + 300
    cmdOk.Left = (Me.ScaleWidth - (cmdOk.Width * 2 + 300)) / 2
    cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300
Else
    cmdOk.Caption = "OK"
    cmdOk.default = True
    cmdOk.Cancel = True
    cmdOk.Visible = True
    cmdEsc.Visible = False
    cmdF5.Visible = False
    cmdAngebotNeu.Visible = False
    cmdAngebotLoeschen.Visible = False
    cmdAngebotEdit.Visible = False
    
    cmdOk.Left = (Me.ScaleWidth - cmdOk.Width) / 2
End If


Me.Height = cmdOk.Top + cmdOk.Height + wPara1.TitelY + 90 + wPara1.FrmCaptionHeight

'Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
'Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

    
If (iNewLine) Then
    If (PixelX = 1680) Then
        iAdd = 210
    Else
        iAdd = 90
    End If
    
    With flxAngebote
        If (.Cols > maxSp%) Then
            .ScrollBars = flexScrollBarHorizontal
        Else
            .ScrollBars = flexScrollBarNone
        End If
'        .ScrollBars = flexScrollBarNone
        .BorderStyle = 0
        .Width = .Width - 90
        .Height = .Height - 90
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridFlat
        .GridColorFixed = .GridColor
        .BackColor = wPara1.nlFlexBackColor    'vbWhite
        .BackColorBkg = wPara1.nlFlexBackColor    'vbWhite
        .BackColorFixed = wPara1.nlFlexBackColorFixed   ' RGB(199, 176, 123)
        .BackColorSel = wPara1.nlFlexBackColorSel  ' RGB(232, 217, 172)
        .ForeColorSel = vbBlack
        
        .Left = .Left + iAdd
        .Top = .Top + iAdd

'        ForeColor = RGB(180, 180, 180) ' vbWhite
'        FillStyle = vbSolid
'        FillColor = vbWhite
'        RoundRect hdc, (.Left - iAdd) / Screen.TwipsPerPixelX, (.Top - iAdd) / Screen.TwipsPerPixelY, (.Left + .Width + iAdd) / Screen.TwipsPerPixelX, (.Top + .Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
    End With
    
    cmdOk.Top = cmdOk.Top + 2 * iAdd
    cmdEsc.Top = cmdEsc.Top + 2 * iAdd
    
    cmdF2.Left = cmdF2.Left + 2 * iAdd
    cmdF5.Left = cmdF5.Left + 2 * iAdd
    
    cmdAngebotNeu.Left = cmdAngebotNeu.Left + 2 * iAdd
    cmdAngebotLoeschen.Left = cmdAngebotLoeschen.Left + 2 * iAdd
    cmdAngebotEdit.Left = cmdAngebotEdit.Left + 2 * iAdd
    
    Width = Width + 2 * iAdd
    Height = Height + 2 * iAdd

    iAdd2 = 450
    For i% = 0 To 2
        lblAngebote(i%).Top = lblAngebote(i%).Top + iAdd2
        lblAngeboteWert(i%).Top = lblAngeboteWert(i%).Top + iAdd2
    Next i%
    flxAngebote.Top = flxAngebote.Top + iAdd2
    cmdOk.Top = cmdOk.Top + iAdd2
    cmdEsc.Top = cmdEsc.Top + iAdd2
    cmdF2.Top = cmdF2.Top + iAdd2
    cmdF5.Top = cmdF5.Top + iAdd2
    cmdAngebotNeu.Top = cmdAngebotNeu.Top + iAdd2
    cmdAngebotLoeschen.Top = cmdAngebotLoeschen.Top + iAdd2
    cmdAngebotEdit.Top = cmdAngebotEdit.Top + iAdd2
    Height = Height + iAdd2
    
    With nlcmdOk
        .Init
'        .Width = 3000
'        .Height = 600
        .Left = (Me.ScaleWidth - (.Width * 2 + 300)) / 2
        .Top = flxAngebote.Top + flxAngebote.Height + iAdd + 600
        .Caption = cmdOk.Caption
        .TabIndex = cmdOk.TabIndex
        .Enabled = cmdOk.Enabled
        .Visible = True
    End With
    cmdOk.Visible = False

    With nlcmdEsc
        .Init
'        .Width = nlcmdOk.Width
'        .Height = nlcmdOk.Height
        .Left = nlcmdOk.Left + .Width + 300
        .Top = nlcmdOk.Top
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .Visible = True
    End With
    cmdEsc.Visible = False
    
    With nlcmdF2
        .Init
'        .Width = 1800
'        .Height = 510
        .Left = cmdF2.Left
        .Top = cmdF2.Top
        .Caption = cmdF2.Caption
        .TabIndex = cmdF2.TabIndex
        .Enabled = cmdF2.Enabled
        .Visible = True 'cmdF2.Visible
    End With
    cmdF2.Visible = False

    With nlcmdF5
        .Init
        .Width = nlcmdF2.Width
'        .Height = nlcmdF2.Height
        .Left = nlcmdF2.Left
        .Top = nlcmdF2.Top + nlcmdF2.Height + 210
        .Caption = cmdF5.Caption
        .TabIndex = cmdF5.TabIndex
        .Enabled = cmdF5.Enabled
        .Visible = True 'cmdF5.Visible
    End With
    cmdF5.Visible = False

    With nlcmdAngebotNeu
        .Init
        .Width = nlcmdF2.Width
'        .Height = nlcmdF2.Height
        .Left = nlcmdF2.Left
        .Top = nlcmdF5.Top + nlcmdF5.Height + 210
        .Caption = cmdAngebotNeu.Caption
        .TabIndex = cmdAngebotNeu.TabIndex
        .Enabled = cmdAngebotNeu.Enabled
        .Visible = False
    End With
    cmdAngebotNeu.Visible = False

    With nlcmdAngebotLoeschen
        .Init
        .Width = nlcmdF2.Width
'        .Height = nlcmdF2.Height
        .Left = nlcmdF2.Left
        .Top = nlcmdAngebotNeu.Top + nlcmdAngebotNeu.Height + 210
        .Caption = cmdAngebotLoeschen.Caption
        .TabIndex = cmdAngebotLoeschen.TabIndex
        .Enabled = cmdAngebotLoeschen.Enabled
        .Visible = False
    End With
    cmdAngebotLoeschen.Visible = False

    With nlcmdAngebotEdit
        .Init
        .Width = nlcmdF2.Width
'        .Height = nlcmdF2.Height
        .Left = nlcmdF2.Left
        .Top = nlcmdAngebotLoeschen.Top + nlcmdAngebotLoeschen.Height + 210
        .Caption = cmdAngebotEdit.Caption
        .TabIndex = cmdAngebotEdit.TabIndex
        .Enabled = cmdAngebotEdit.Enabled
        .Visible = False
    End With
    cmdAngebotEdit.Visible = False

    If (AngebotModus% = 1) Then
'        cmdOk.Caption = "Binden"
        nlcmdOk.default = True
        nlcmdOk.Cancel = False
        nlcmdOk.Visible = True
        nlcmdEsc.Cancel = True
        nlcmdEsc.Visible = True
        nlcmdF5.Visible = (AngebotInd% >= 0)
        nlcmdAngebotNeu.Visible = (AngebotActLief% >= 0)
    Else
'        cmdOk.Caption = "OK"
        nlcmdOk.default = True
        nlcmdOk.Cancel = True
        nlcmdOk.Visible = True
        nlcmdEsc.Visible = False
        nlcmdF5.Visible = False
        nlcmdAngebotNeu.Visible = False
        nlcmdAngebotLoeschen.Visible = False
        nlcmdAngebotEdit.Visible = False
    End If
    
    Me.Width = nlcmdF2.Left + nlcmdF2.Width + 600
    Me.Height = nlcmdOk.Top + nlcmdOk.Height + wPara1.FrmCaptionHeight + 450

    Call wPara1.NewLineWindow(Me, nlcmdOk.Top)
'    RoundRect hdc, (flxAngebote.Left - iAdd) / Screen.TwipsPerPixelX, (flxAngebote.Top - iAdd) / Screen.TwipsPerPixelY, (flxAngebote.Left + flxAngebote.Width + iAdd) / Screen.TwipsPerPixelX, (flxAngebote.Top + flxAngebote.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
    
    Me.Left = ProjektForm.Left + (ProjektForm.Width - Me.Width) / 2
    Me.Top = ProjektForm.Top + (ProjektForm.Height - Me.Height) / 2
' Kopfhöhe: 30 Pixel
'  RGB 247 ... 12 Pixel
'   230-191

'   177-225
'unten 177, 45 Pixel hoch
Else
    nlcmdOk.Visible = False
    nlcmdEsc.Visible = False
    nlcmdF2.Visible = False
    nlcmdF5.Visible = False
    nlcmdAngebotNeu.Visible = False
    nlcmdAngebotLoeschen.Visible = False
    nlcmdAngebotEdit.Visible = False
End If

F5Gedrueckt% = False

Call clsError.DefErrPop
End Sub

Private Sub Form_Paint()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_Paint")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, spBreite%, ind%, iAnzZeilen%, RowHe%, bis%, bis2%
Dim sp&
Dim h$, h2$
Dim iAdd%, iAdd2%, wi%
Dim c As Control

If (Para1.Newline) Then
    iAdd = wPara1.NlFlexBackY
    iAdd2 = wPara1.NlCaptionY
    
    Call wPara1.NewLineWindow(Me, nlcmdOk.Top, False)
    RoundRect hdc, (flxAngebote.Left - iAdd) / Screen.TwipsPerPixelX, (flxAngebote.Top - iAdd) / Screen.TwipsPerPixelY, (flxAngebote.Left + flxAngebote.Width + iAdd) / Screen.TwipsPerPixelX, (flxAngebote.Top + flxAngebote.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20

    Call Form_Resize
End If

Call clsError.DefErrPop
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_QueryUnload")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
AngebotY% = Me.Top
Call clsError.DefErrPop
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_MouseDown")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
If (y <= 450) Then
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

Call clsError.DefErrPop
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_MouseMove")
Call clsError.DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case clsError.DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call ProjektForm.EndeDll
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

Call clsError.DefErrPop
End Sub

Private Sub Form_Resize()
If (iNewLine) And (Me.Visible) Then
    CurrentX = 210
    CurrentY = (450 - TextHeight(Caption)) / 2
    ForeColor = vbBlack
    Me.Print Caption
End If
End Sub

Private Sub nlcmdOk_Click()
Call cmdOk_Click
End Sub

Private Sub nlcmdEsc_Click()
Call cmdEsc_Click
End Sub

Private Sub nlcmdF2_Click()
Call cmdF2_Click
End Sub

Private Sub nlcmdF5_Click()
Call cmdF5_Click
End Sub

Private Sub nlcmdAngebotNeu_Click()
Call cmdAngebotNeu_Click
End Sub

Private Sub nlcmdAngebotLoeschen_Click()
Call cmdAngebotLoeschen_Click
End Sub

Private Sub nlcmdAngebotEdit_Click()
Call cmdAngebotEdit_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If (iNewLine) Then
    If (KeyAscii = 13) Then
        Call nlcmdOk_Click
    ElseIf (KeyAscii = 27) Then
        Call nlcmdEsc_Click
    End If
End If

End Sub

Private Sub picControlBox_Click(Index As Integer)

If (Index = 0) Then
    Me.WindowState = vbMinimized
ElseIf (Index = 1) Then
    Me.WindowState = vbNormal
Else
    Unload Me
End If

End Sub
