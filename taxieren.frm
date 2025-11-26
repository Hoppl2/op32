VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlbutton.ocx"
Begin VB.Form frmTaxieren 
   AutoRedraw      =   -1  'True
   Caption         =   "Individuelle Rezeptur"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   11640
   Begin VB.CommandButton cmdSonderfälle 
      Caption         =   "+"
      Height          =   450
      Left            =   9000
      TabIndex        =   37
      Top             =   4800
      Width           =   480
   End
   Begin VB.ComboBox cboSonderfälle 
      Height          =   315
      Left            =   7440
      Sorted          =   -1  'True
      Style           =   2  'Dropdown-Liste
      TabIndex        =   36
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   4440
      TabIndex        =   13
      Top             =   4320
      Width           =   1200
   End
   Begin VB.CommandButton cmdVerwurf 
      Caption         =   "&Verwurf"
      Height          =   450
      Left            =   7560
      TabIndex        =   33
      Top             =   4200
      Width           =   1200
   End
   Begin VB.CommandButton cmdTrägerLösung 
      Caption         =   "Träger&Lösung"
      Height          =   450
      Left            =   7560
      TabIndex        =   31
      Top             =   3600
      Width           =   1200
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   3120
      Picture         =   "taxieren.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   3360
      Picture         =   "taxieren.frx":00A9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   3600
      Picture         =   "taxieren.frx":015D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdAuswahl 
      Caption         =   "&FAM=Substanz"
      Height          =   450
      Index           =   4
      Left            =   7320
      TabIndex        =   11
      Top             =   2520
      Width           =   1200
   End
   Begin VB.CheckBox chkKassenRabatt 
      Caption         =   "&Kassen-Rabatt"
      Height          =   495
      Index           =   0
      Left            =   2280
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdF7 
      Caption         =   "Inklusive (F7)"
      Enabled         =   0   'False
      Height          =   450
      Left            =   4320
      TabIndex        =   4
      Top             =   1440
      Width           =   1200
   End
   Begin VB.CommandButton cmdF5 
      Caption         =   "Löschen (F5)"
      Height          =   450
      Left            =   4320
      TabIndex        =   3
      Top             =   840
      Width           =   1200
   End
   Begin VB.CommandButton cmdF2 
      Caption         =   "Einfügen (F2)"
      Height          =   450
      Left            =   4320
      TabIndex        =   2
      Top             =   240
      Width           =   1200
   End
   Begin VB.TextBox txtTaxieren 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Tag             =   "0"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdAuswahl 
      Caption         =   "S&onstige"
      Height          =   450
      Index           =   3
      Left            =   7320
      TabIndex        =   10
      Top             =   1920
      Width           =   1200
   End
   Begin VB.CommandButton cmdAuswahl 
      Caption         =   "&Spezialität"
      Height          =   450
      Index           =   2
      Left            =   7320
      TabIndex        =   9
      Top             =   1320
      Width           =   1200
   End
   Begin VB.CommandButton cmdAuswahl 
      Caption         =   "&Arbeit"
      Height          =   450
      Index           =   1
      Left            =   7320
      TabIndex        =   8
      Top             =   720
      Width           =   1200
   End
   Begin VB.CommandButton cmdAuswahl 
      Caption         =   "&Gefäß"
      Height          =   450
      Index           =   0
      Left            =   7320
      TabIndex        =   7
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton cmdDarstellung 
      Caption         =   "&Darstellung ..."
      Height          =   450
      Left            =   600
      TabIndex        =   6
      Top             =   4920
      Width           =   1200
   End
   Begin VB.CommandButton cmdTaxmuster 
      Caption         =   "&Taxmuster ..."
      Height          =   450
      Left            =   600
      TabIndex        =   5
      Top             =   4320
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   5880
      TabIndex        =   14
      Top             =   4320
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxTaxieren 
      Height          =   1620
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   2858
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
   End
   Begin MSFlexGridLib.MSFlexGrid flxTaxSumme 
      Height          =   1380
      Left            =   480
      TabIndex        =   15
      Top             =   2040
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   2434
      _Version        =   393216
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      BackColorBkg    =   -2147483633
      GridColor       =   -2147483626
      Enabled         =   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   375
      Left            =   5880
      TabIndex        =   19
      Top             =   4920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdF2 
      Height          =   495
      Left            =   5760
      TabIndex        =   20
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdDarstellung 
      Height          =   375
      Left            =   1920
      TabIndex        =   21
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdTaxmuster 
      Height          =   375
      Left            =   1920
      TabIndex        =   22
      Top             =   4320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdF5 
      Height          =   495
      Left            =   5760
      TabIndex        =   23
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdF7 
      Height          =   495
      Left            =   5760
      TabIndex        =   24
      Top             =   1560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdAuswahl 
      Height          =   495
      Index           =   0
      Left            =   8880
      TabIndex        =   25
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdAuswahl 
      Height          =   495
      Index           =   1
      Left            =   8880
      TabIndex        =   26
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdAuswahl 
      Height          =   495
      Index           =   2
      Left            =   8880
      TabIndex        =   27
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdAuswahl 
      Height          =   495
      Index           =   3
      Left            =   8880
      TabIndex        =   28
      Top             =   1920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdAuswahl 
      Height          =   495
      Index           =   4
      Left            =   8880
      TabIndex        =   29
      Top             =   2520
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdTrägerLösung 
      Height          =   375
      Left            =   8880
      TabIndex        =   32
      Top             =   3600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdVerwurf 
      Height          =   375
      Left            =   8880
      TabIndex        =   34
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   375
      Left            =   4560
      TabIndex        =   35
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdSonderfälle 
      Height          =   375
      Left            =   9600
      TabIndex        =   38
      Top             =   4800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin VB.Label lblchkKassenRabatt 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   30
      Top             =   3480
      Width           =   2055
   End
End
Attribute VB_Name = "frmTaxieren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TxtCol%
Dim TaxierTyp As Byte

Dim OrgRezepturMitFaktor%

'Dim ParenteralSpez(10) As TaxierungStruct
'Dim ParenteralEk#(10)
'Dim ParenteralPreisProEinheit#(10)
'Dim ParEnteralAnzEinheiten#(10)

Dim AnzParenteralSpez%

Dim iArbeitAnzzeilen%

Private Const DefErrModul = "TAXIEREN.FRM"

Private Sub chkKassenRabatt_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("chkKassenRabatt_Click")
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

If (chkKassenRabatt(0).Value) Then
    RezepturMitFaktor% = True
Else
    RezepturMitFaktor% = False
End If
Call TaxSumme

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
Dim ok%

ok = True
If (ActiveControl.Name = txtTaxieren.Name) Then
    With flxTaxieren
        ok = (TxtCol% = 3) And (.row = (.Rows - 1)) And (Trim(txtTaxieren.text) = "")
    End With
End If
If (ok) Then
    If (TaxmusterDBok) Then
        TaxmusterDB.Close
    Else
        If (TM_NAMEN% > 0) Then Close (TM_NAMEN%)
        If (TM_DATEN% > 0) Then Close (TM_DATEN%)
    End If
    
    If (AnfMagIndex& > 0) Then
'            Call SpeicherAnfMag
    Else
        If (MagSpeicherIndex% = 0) Then
            If (ParenteralRezept >= 0) Then
                MagSpeicherIndex = 1
            Else
                MagSpeicherIndex% = LOF(MAG_SPEICHER%) / Len(TaxierRec) + 1
            End If
        End If
        If (MagSpeicherIndex% > 0) Then
            Call SpeicherMagSpeicher
        End If
    End If
    
    If (MAG_SPEICHER% > 0) Then Close (MAG_SPEICHER%)
    
''        If (ANF_MAG% > 0) Then Close (ANF_MAG%)
'        AbholerDB.Close
    
    Unload Me
Else
    Call txtTaxieren_KeyPress(13)
End If

RezepturMitFaktor% = OrgRezepturMitFaktor%

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
Dim ok%

ok = True
On Error Resume Next
If (ActiveControl.Name = txtTaxieren.Name) Then
    With flxTaxieren
        ok = (TxtCol% = 3) And (.row = (.Rows - 1)) 'And (Trim(txtTaxieren.text) = "")
    End With
End If
On Error GoTo DefErr
If (ok) Then
    SumPreis = 0
    
    If (TaxmusterDBok) Then
        TaxmusterDB.Close
    Else
        If (TM_NAMEN% > 0) Then Close (TM_NAMEN%)
        If (TM_DATEN% > 0) Then Close (TM_DATEN%)
    End If
    
    If (MAG_SPEICHER% > 0) Then Close (MAG_SPEICHER%)
    
    Unload Me
Else
    With flxTaxieren
        If (.row = (.Rows - 1)) Then .AddItem " "
        TxtCol% = 3
        .row = .Rows - 1
        Call ShowEditBox
    End With
End If

'With flxTaxieren
'    If (TxtCol% = 3) And (.row = (.Rows - 1)) Then
'        If (TaxmusterDBok) Then
'            TaxmusterDB.Close
'        Else
'            If (TM_NAMEN% > 0) Then Close (TM_NAMEN%)
'            If (TM_DATEN% > 0) Then Close (TM_DATEN%)
'        End If
'
'        If (AnfMagIndex& > 0) Then
''            Call SpeicherAnfMag
'        Else
'            If (MagSpeicherIndex% = 0) Then
'                MagSpeicherIndex% = LOF(MAG_SPEICHER%) / Len(TaxierRec) + 1
'            End If
'            If (MagSpeicherIndex% > 0) Then
'                Call SpeicherMagSpeicher
'            End If
'        End If
'
'        If (MAG_SPEICHER% > 0) Then Close (MAG_SPEICHER%)
'
'''        If (ANF_MAG% > 0) Then Close (ANF_MAG%)
''        AbholerDB.Close
'
'        Unload Me
'    Else
'        If (.row = (.Rows - 1)) Then .AddItem " "
'        TxtCol% = 3
'        .row = .Rows - 1
'        Call ShowEditBox
'    End If
'End With

RezepturMitFaktor% = OrgRezepturMitFaktor%

Call DefErrPop
End Sub

Private Sub cmdAuswahl_Click(index As Integer)
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
Dim h$

If (TxtCol% = 3) Then
    txtTaxieren.text = Format(index, "0")
    Call txtTaxieren_KeyPress(13)
End If

Call DefErrPop
End Sub

Private Sub cmdDarstellung_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdDarstellung_Click")
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
Dim i%, OrgRow%, OrgCol%
Dim h$

frmMagDarstellung.Show 1
If (FormErg%) Then
    With flxTaxieren
        .Redraw = False
        OrgRow% = .row
        OrgCol% = .col
        For i% = 1 To (.Rows - 1)
            h$ = .TextMatrix(i%, 6)
            If (h$ <> "") Then
                .FillStyle = flexFillRepeat
                .row = i%
                .col = 0
                .RowSel = .row
                .ColSel = .Cols - 1
                .CellForeColor = MagDarstellung&(h$, 0)
                .CellBackColor = MagDarstellung&(h$, 1)
                .FillStyle = flexFillSingle
            End If
        Next i%
        .row = OrgRow%
        .col = OrgCol%
        .Redraw = True
    End With
End If
If (txtTaxieren.Visible) Then txtTaxieren.SetFocus

Call DefErrPop
End Sub

Private Sub cmdF2_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdF2_Click")
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

If (para.Newline) And (nlcmdF2.Enabled = False) Then Call DefErrPop: Exit Sub
If (para.Newline = 0) And (cmdF2.Enabled = False) Then Call DefErrPop: Exit Sub

With flxTaxieren
    If (.row < (.Rows - 1)) Then
        txtTaxieren.Visible = False
        .AddItem "", .row
        TxtCol% = 3
        Call ShowEditBox
        Call TaxSumme
    End If
End With

Call DefErrPop
End Sub

Private Sub cmdF5_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdF5_Click")
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

If (para.Newline) And (nlcmdF5.Enabled = False) Then Call DefErrPop: Exit Sub
If (para.Newline = 0) And (cmdF5.Enabled = False) Then Call DefErrPop: Exit Sub

With flxTaxieren
    If (.row < (.Rows - 1)) Then
        txtTaxieren.Visible = False
        .RemoveItem .row
        Call MakeEditCol
        Call TaxSumme
    End If
End With

Call DefErrPop
End Sub

Private Sub cmdF7_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdF7_Click")
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
Dim row%
Dim dVal#
Dim txt$

If (para.Newline) And (nlcmdF7.Enabled = False) Then Call DefErrPop: Exit Sub
If (para.Newline = 0) And (cmdF7.Enabled = False) Then Call DefErrPop: Exit Sub

With txtTaxieren
    txt$ = UCase(Trim(.text))
    If (txt$ <> "") Then
        dVal# = (Val(txt$) / (100# + para.Mwst(2))) * 100# / MalFaktor#
        .text = Format(dVal#, "0.000")
        
        If (TxtCol% = 0) Then
            Call txtTaxieren_KeyPress(13)
        Else
            Call UmspeichernPreisEingabe
            TaxierRec.ActPreis = dVal# * MalFaktor#
            Call ZeigeTaxierZeile(flxTaxieren.row)
            
            Call TaxSumme
            
            row% = flxTaxieren.row
            If (row% = flxTaxieren.Rows - 1) Then
                flxTaxieren.AddItem " "
                flxTaxieren.row = flxTaxieren.Rows - 1
                TxtCol% = 3
                Call ShowEditBox
            Else
                Call txtTaxieren_KeyDown(vbKeyDown, 0)
            End If
        End If
    End If
End With

Call DefErrPop
End Sub

Private Sub cmdSonderfälle_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdSonderfälle_Click")
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
Dim row%, ind%
Dim h$, PreisStr$, PznStr$

h = Trim(cboSonderfälle.text)
If (h <> "") Then
    With frmTaxieren.flxTaxieren
        If (TxtCol <> 3) Then
            .row = .Rows - 1
            If (.TextMatrix(.row, 0) = "") Then
                .AddItem " "
                .row = .Rows - 1
            End If
            row% = .row
        End If
        
        With TaxierRec
            ind = InStr(h, vbTab)
            If (ind > 0) Then
                PreisStr = Trim(Mid(h, ind + 1))
                h = Trim(Left(h, ind - 1))
                PznStr = Right(h, 8)
                h = Trim(Left(h, Len(h) - 20))
            End If
            
            .pzn = PznStr   ' "02567001"   ' Space$(Len(.pzn))
            .kurz = Left$(h + Space$(Len(.kurz)), Len(.kurz))
            .menge = Space$(Len(.menge))
            .Meh = Space$(Len(.Meh))
            .kp = 0
            .GStufe = 0
            
            .ActMenge = 1#
            .ActPreis = xVal(PreisStr) / (1# + (para.Mwst(2) / 100#)) ' 3.58
            .kp = .ActPreis
            
            .flag = MAG_PREISEINGABE
        End With
        Call frmTaxieren.ZeigeTaxierZeile(.row)
        
        Call TaxSumme
        
        row% = flxTaxieren.row
        If (row% = flxTaxieren.Rows - 1) Then
            flxTaxieren.AddItem " "
            flxTaxieren.row = flxTaxieren.Rows - 1
            TxtCol% = 3
            Call ShowEditBox
        Else
            Call txtTaxieren_KeyDown(vbKeyDown, 0)
        End If
    
    End With
End If
    
Call DefErrPop
End Sub

Private Sub cmdTaxmuster_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdTaxmuster_Click")
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

With flxTaxieren
'    If (.Rows <= 2) And (.TextMatrix(1, 3) = "") Then
    If ((ParenteralRezept >= 0) And (.Rows <= 3) And (TxtCol% = 3)) Or _
       ((ParenteralRezept < 0) And (.Rows <= 2) And (TxtCol% = 3)) Then
        TaxMusterModus% = 0
        TaxMusterSuch$ = UCase(Trim(txtTaxieren.text))
        frmTaxMuster.Show 1
        If (FormErg%) Then
            Call HoleTaxMuster
            .AddItem " "
            .row = .Rows - 1
        End If
    Else
        TaxMusterModus% = 1
        frmTaxMuster.Show 1
        If (FormErg%) Then Call SpeicherTaxMuster
    End If
    TxtCol% = 3
    Call ShowEditBox
End With

Call DefErrPop
End Sub

Private Sub cmdTrägerLösung_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdTrägerLösung_Click")
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
Dim row%, OrgRow%
Dim dVal#
Dim txt$

row% = flxTaxieren.row
OrgRow = row
If (Val(flxTaxieren.TextMatrix(row%, 6)) <> MAG_SPEZIALITAET) Then
    row = row - 1
End If

TaxierRec.ActMenge = iCDbl(flxTaxieren.TextMatrix(row%, 1))
TaxierRec.Meh = flxTaxieren.TextMatrix(row%, 2)
TaxierRec.kurz = flxTaxieren.TextMatrix(row%, 3)
TaxierRec.pzn = flxTaxieren.TextMatrix(row%, 4)
TaxierRec.flag = Val(flxTaxieren.TextMatrix(row%, 6))
TaxierRec.kp = iCDbl(flxTaxieren.TextMatrix(row%, 7))
TaxierRec.GStufe = iCDbl(flxTaxieren.TextMatrix(row%, 8))
TaxierRec.Verwurf = iCDbl(flxTaxieren.TextMatrix(row%, 5))

TaxierTyp = TaxierRec.flag

If (TaxierTyp = MAG_SPEZIALITAET) Then
    TaxierRec.ActPreis = 1.2
    Call ZeigeTaxierZeile(row%)
    Call TaxSumme
End If
flxTaxieren.row = OrgRow
Call ShowEditBox

Call DefErrPop
End Sub

Private Sub cmdVerwurf_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdVerwurf_Click")
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
Dim row%, OrgRow%
Dim dVal#
Dim txt$

row% = flxTaxieren.row
'OrgRow = row
'If (Val(flxTaxieren.TextMatrix(row%, 6)) <> MAG_SPEZIALITAET) Then
'    row = row - 1
'End If

TaxierRec.ActMenge = iCDbl(flxTaxieren.TextMatrix(row%, 1))
TaxierRec.Meh = flxTaxieren.TextMatrix(row%, 2)
TaxierRec.kurz = flxTaxieren.TextMatrix(row%, 3)
TaxierRec.pzn = flxTaxieren.TextMatrix(row%, 4)
TaxierRec.flag = Val(flxTaxieren.TextMatrix(row%, 6))
TaxierRec.kp = iCDbl(flxTaxieren.TextMatrix(row%, 7))
TaxierRec.GStufe = iCDbl(flxTaxieren.TextMatrix(row%, 8))
'TaxierRec.Verwurf = iCDbl(flxTaxieren.TextMatrix(row%, 5))
TaxierRec.Verwurf = 1   ' iCDbl(flxTaxieren.TextMatrix(row%, 12))
Call ZeigeTaxierZeile(row%)

If (row% < (flxTaxieren.Rows - 1)) Then
    flxTaxieren.row = row% + 1
    txtTaxieren.Visible = False
    Call MakeEditCol
Else
    Call ShowEditBox
End If

Call DefErrPop
End Sub

Private Sub flxTaxieren_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxTaxieren_GotFocus")
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

With flxTaxieren
    .col = 0
    .ColSel = .Cols - 1
    .HighLight = flexHighlightAlways
End With

Call DefErrPop
End Sub

Private Sub flxTaxieren_LostFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxTaxieren_LostFocus")
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

With flxTaxieren
    .HighLight = flexHighlightNever
End With

Call DefErrPop
End Sub

Private Sub flxTaxieren_KeyPress(KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxTaxieren_KeyPress")
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
Dim i%, row%, gef%, col%
Dim ch$, h$

ch$ = UCase$(Chr$(KeyAscii))

If (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890", ch$) > 0) Then
    gef% = False
    With flxTaxieren
        row% = .row
        For i% = (row% + 1) To (.Rows - 1)
            If (UCase(Left$(.TextMatrix(i%, 1), 1)) = ch$) Then
                .row = i%
                gef% = True
                Exit For
            End If
        Next i%
        If (gef% = False) Then
            For i% = 1 To (row% - 1)
                If (UCase(Left$(.TextMatrix(i%, 1), 1)) = ch$) Then
                    .row = i%
                    gef% = True
                    Exit For
                End If
            Next i%
        End If
        If (gef% = True) Then
'            If (.row < .TopRow) Then .TopRow = .row
            .TopRow = .row
            .col = 0
            .ColSel = .Cols - 1
        End If
    End With
End If

Call DefErrPop
End Sub

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
Dim ind%
Dim c As Control

If (Shift And vbAltMask) Then
    If (KeyCode <> 18) Then
        On Error Resume Next
        For Each c In Controls
            If (TypeOf c Is nlCommand) Then
                If (c.Accelerator <> "") And (c.Enabled) Then
                    If (Asc(c.Accelerator) = KeyCode) Then
                        c.SetFocus
                        c.Value = 1
'                        Exit For
                        Call DefErrPop: Exit Sub
                    End If
                End If
            End If
        Next
        On Error GoTo DefErr
    End If
End If

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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%, lief%, erg%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%, FeldInd%
Dim iAdd%, iAdd2%
Dim h$, h2$, FormStr$, PreisStr$
Dim c As Control

Call wpara.InitFont(Me)

'ParenteralRezept = -1

TaxmusterDBok = (Dir(TAXMUSTER_DB) <> "")
If (TaxmusterDBok) Then
    Set TaxmusterDB = OpenDatabase(TAXMUSTER_DB, False, False)
Else
    TM_NAMEN% = FileOpen("tmnamen.dat", "RW", "B")
    TM_DATEN% = FileOpen("tmdaten.dat", "RW", "B")
    
    If (LOF(TM_NAMEN%) = 0) Then
        h$ = String(Len(TmHeader), 0)
        Put #TM_NAMEN%, , h$
    End If
    If (LOF(TM_DATEN%) = 0) Then
        h$ = String(Len(TmInhalt), 0)
        Put #TM_DATEN%, , h$
    End If
End If


If (MagSpeicherIndex <= 0) Then
    On Error Resume Next
    Kill "MagTax" + Trim(para.User)
    On Error GoTo DefErr
End If
MAG_SPEICHER% = FileOpen("MagTax" + Trim(para.User), "RW", "B")

'erg% = OpenAbholer%

With flxTaxieren
    .Rows = 2
    .FixedRows = 1
    .FormatString = ">Preis|>Menge|Meh|<Kurzbezeichnung|<PZN||Flag|KP|Gstufe|PE_PM|PE_AI|PE_AnzEinheiten|Verwurf|>PreisUNgerundet"
    .Rows = 1
    
    .ColWidth(0) = TextWidth("9999999.99")
    .ColWidth(1) = TextWidth(String(8, "9"))
    .ColWidth(2) = TextWidth(String(4, "A"))
    .ColWidth(3) = TextWidth(String(28, "A"))
    .ColWidth(4) = TextWidth(String(9, "9"))
    .ColWidth(5) = wpara.FrmScrollHeight
    For i = 6 To 13
        .ColWidth(i) = 0
    Next i
    
    Breite1% = 0
    For i% = 0 To (.Cols - 1)
        Breite1% = Breite1% + .ColWidth(i%)
    Next i%
    .Width = Breite1% + 90
    
    iArbeitAnzzeilen% = 15
    .Height = .RowHeight(0) * iArbeitAnzzeilen% + 90
    
    .Top = wpara.TitelY
    .Left = wpara.LinksX
    
'    If (MagSpeicherIndex% > 0) Then
'        Call HoleMagSpeicher
'    End If
'
'    .AddItem " "
'    .row = .Rows - 1
End With

With flxTaxSumme
    .Rows = 3
    If (RezepturMitFaktor%) Then .Rows = .Rows + 1
    .FixedRows = 0
    .Cols = 3
    
    For i% = 0 To 1
        .ColWidth(i%) = flxTaxieren.ColWidth(i%)
    Next i%
    .ColWidth(2) = 2 * flxTaxieren.ColWidth(2)
    
    .ColAlignment(2) = flexAlignRightCenter
    
    Breite1% = 0
    For i% = 0 To (.Cols - 1)
        Breite1% = Breite1% + .ColWidth(i%)
    Next i%
    .Width = Breite1% + 90
    .Height = .RowHeight(0) * .Rows + 90
    
    .Top = flxTaxieren.Top + flxTaxieren.Height
    .Left = wpara.LinksX
    
    .ScrollBars = flexScrollBarNone
End With

OrgRezepturMitFaktor% = RezepturMitFaktor%
If (RezepturMitFaktor%) Then
    With chkKassenRabatt(0)
        .Left = flxTaxSumme.Left + flxTaxSumme.Width + 150
        .Top = flxTaxSumme.Top + flxTaxSumme.RowPos(2)
        .Value = 1
        .Visible = True
    End With
End If

'AnfMagIndex& = 0    'für Zwischenversion

With flxTaxieren
    If (MagSpeicherIndex% > 0) Then
        Call HoleMagSpeicher
    ElseIf (AnfMagIndex& > 0) Then
'        Call HoleAnfMag
    End If
    
    .AddItem " "
    .row = .Rows - 1
End With

TxtCol% = 3
Call ShowEditBox

Font.Bold = False   ' True

With cmdF2
    .Left = flxTaxieren.Left + flxTaxieren.Width + 300
    .Top = flxTaxieren.Top
    .Width = TextWidth(.Caption) + 300
'    .Height = wpara.ButtonY
    .Height = TextHeight("Äg") + 240
End With
With cmdF5
    .Left = cmdF2.Left
    .Top = cmdF2.Top + cmdF2.Height + 90
    .Width = cmdF2.Width
    .Height = cmdF2.Height
End With
With cmdF7
    .Left = cmdF2.Left
    .Top = cmdF5.Top + cmdF5.Height + 90
    .Width = cmdF2.Width
    .Height = cmdF2.Height
End With

With cmdTaxmuster
    .Left = cmdF2.Left
    .Top = cmdF7.Top + cmdF7.Height + 450
    .Width = cmdF2.Width
    .Height = cmdF2.Height
End With

With cmdDarstellung
    .Left = cmdF2.Left
    .Top = cmdTaxmuster.Top + cmdTaxmuster.Height + 90
    .Width = cmdF2.Width
    .Height = cmdF2.Height
End With

For i% = 0 To 4
    With cmdAuswahl(i%)
        .Left = cmdF2.Left
        If (i% = 0) Then
            .Top = cmdDarstellung.Top + cmdDarstellung.Height + 450
        Else
            .Top = cmdAuswahl(i% - 1).Top + cmdAuswahl(i% - 1).Height + 60
        End If
        .Width = cmdTaxmuster.Width
        .Height = cmdTaxmuster.Height
    End With
Next i%

With cmdTrägerLösung
    .Left = cmdF2.Left
    .Width = cmdF2.Width
    .Height = cmdF2.Height
    .Top = flxTaxieren.Top + flxTaxieren.Height - .Height
    .Visible = False
End With
With cmdVerwurf
    .Left = cmdF2.Left
    .Width = cmdF2.Width
    .Height = cmdF2.Height
    .Top = cmdTrägerLösung.Top + cmdTrägerLösung.Height + 60
    .Visible = False
End With

With cboSonderfälle
    .Left = cmdF2.Left
    .Width = cmdF2.Width
    .Top = cmdTrägerLösung.Top - .Height - 300
End With
With cmdSonderfälle
    .Left = cboSonderfälle.Left + cboSonderfälle.Width + 30
'    .Width = cmdF2.Width
    .Height = cmdF2.Height
    .Top = cboSonderfälle.Top
End With


Me.Width = cmdTaxmuster.Left + cmdTaxmuster.Width + 2 * wpara.LinksX

With cmdEsc
    .Width = wpara.ButtonX
    .Height = wpara.ButtonY
    .Top = flxTaxSumme.Top + flxTaxSumme.Height + wpara.ButtonY + 150
    .Left = flxTaxieren.Left + flxTaxieren.Width - .Width
End With
With cmdOk
    .Width = wpara.ButtonX
    .Height = wpara.ButtonY
    .Top = cmdEsc.Top
'    .Left = (Me.Width - .Width) / 2
    .Left = cmdEsc.Left - .Width - 300
End With


Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

If (para.Newline) Then
    iAdd = wpara.NlFlexBackY
    iAdd2 = wpara.NlCaptionY
    
    With flxTaxieren
'        .ScrollBars = flexScrollBarNone
        .BorderStyle = 0
        .Width = .Width - 90
        .Height = .Height - 90
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridFlat
        .GridColorFixed = .GridColor
        .BackColor = wpara.nlFlexBackColor    'vbWhite
        .BackColorBkg = wpara.nlFlexBackColor    'vbWhite
        .BackColorFixed = wpara.nlFlexBackColorFixed   ' RGB(199, 176, 123)
        .BackColorSel = wpara.nlFlexBackColorSel  ' RGB(232, 217, 172)
        .ForeColorSel = vbBlack
        
        .Left = .Left + iAdd
        .Top = .Top + iAdd
        
        iArbeitAnzzeilen = 20
        .Height = .RowHeight(0) * iArbeitAnzzeilen%
    End With
    With flxTaxSumme
'        .ScrollBars = flexScrollBarNone
        .BorderStyle = 0
        .Width = .Width - 90
        .Height = .Height - 90
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridFlat
        .GridColorFixed = .GridColor
        .BackColor = wpara.nlFlexBackColor    'vbWhite
        .BackColorBkg = wpara.nlFlexBackColor    'vbWhite
        .BackColorFixed = wpara.nlFlexBackColorFixed   ' RGB(199, 176, 123)
        .BackColorSel = wpara.nlFlexBackColorSel  ' RGB(232, 217, 172)
        .ForeColorSel = vbBlack
        
        .Left = .Left + iAdd
        .Top = flxTaxieren.Top + flxTaxieren.Height
    End With
    With chkKassenRabatt(0)
        .Left = flxTaxSumme.Left + flxTaxSumme.Width + 210
        .Top = flxTaxSumme.Top + flxTaxSumme.RowPos(2)
    End With
    
    cmdF2.Left = cmdF2.Left + 2 * iAdd
    
    cmdEsc.Top = cmdEsc.Top + 2 * iAdd
    cmdDarstellung.Top = cmdEsc.Top
    cmdTaxmuster.Top = cmdEsc.Top
    
    Width = Width + 2 * iAdd
    Height = Height + 2 * iAdd

    On Error Resume Next
    For Each c In Controls
        If (c.Container Is Me) Then
            c.Top = c.Top + iAdd2
        End If
    Next
    On Error GoTo DefErr
    
    Height = Height + iAdd2
    
    With nlcmdOk
        .Init
        .Left = (Me.ScaleWidth - 2 * .Width - 300)
        .Top = flxTaxSumme.Top + flxTaxSumme.Height + iAdd + 600
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
        .Left = Me.ScaleWidth - .Width - 150
        .Top = flxTaxSumme.Top + flxTaxSumme.Height + iAdd + 600
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .Default = cmdEsc.Default
        .Cancel = cmdEsc.Cancel
        .Visible = True
    End With
    cmdEsc.Visible = False

    With nlcmdTaxmuster
        .Init
        .AutoSize = True
        .Left = flxTaxieren.Left
        .Top = nlcmdEsc.Top
        .Caption = cmdTaxmuster.Caption
        .TabIndex = cmdTaxmuster.TabIndex
        .Enabled = cmdTaxmuster.Enabled
        .Default = cmdTaxmuster.Default
        .Cancel = cmdTaxmuster.Cancel
        .Visible = True
    End With
    cmdTaxmuster.Visible = False

    With nlcmdDarstellung
        .Init
        .AutoSize = True
        .Left = nlcmdTaxmuster.Left + nlcmdTaxmuster.Width + 90
        .Top = nlcmdEsc.Top
        .Caption = cmdDarstellung.Caption
        .TabIndex = cmdDarstellung.TabIndex
        .Enabled = cmdDarstellung.Enabled
        .Default = cmdDarstellung.Default
        .Cancel = cmdDarstellung.Cancel
        .Visible = True
    End With
    cmdDarstellung.Visible = False

    With nlcmdF2
        .Init
        .AutoSize = True
        .Caption = cmdAuswahl(4).Caption
        .AutoSize = 0
        .Left = cmdF2.Left
        .Top = cmdF2.Top
        .Caption = cmdF2.Caption
        .TabIndex = cmdF2.TabIndex
        .Enabled = cmdF2.Enabled
        .Default = cmdF2.Default
        .Cancel = cmdF2.Cancel
        .Visible = True
    End With
    cmdF2.Visible = False

    With nlcmdF5
        .Init
        .Width = nlcmdF2.Width
        .Left = nlcmdF2.Left
        .Top = nlcmdF2.Top + nlcmdF2.Height + 45
        .Caption = cmdF5.Caption
        .TabIndex = cmdF5.TabIndex
        .Enabled = cmdF5.Enabled
        .Default = cmdF5.Default
        .Cancel = cmdF5.Cancel
        .Visible = True
    End With
    cmdF5.Visible = False

    With nlcmdF7
        .Init
        .Width = nlcmdF2.Width
        .Left = nlcmdF2.Left
        .Top = nlcmdF5.Top + nlcmdF5.Height + 45
        .Caption = cmdF7.Caption
        .TabIndex = cmdF7.TabIndex
        .Enabled = cmdF7.Enabled
        .Default = cmdF7.Default
        .Cancel = cmdF7.Cancel
        .Visible = True
    End With
    cmdF7.Visible = False

    For i% = 0 To 4
        With nlcmdAuswahl(i%)
            .Width = nlcmdF2.Width
            .Left = nlcmdF2.Left
            If (i% = 0) Then
                .Top = nlcmdF7.Top + nlcmdF7.Height + 210
            Else
                .Top = nlcmdAuswahl(i% - 1).Top + nlcmdAuswahl(i% - 1).Height + 45
            End If
            .Caption = cmdAuswahl(i).Caption
            .TabIndex = cmdAuswahl(i).TabIndex
            .Enabled = cmdAuswahl(i).Enabled
            .Default = cmdAuswahl(i).Default
            .Cancel = cmdAuswahl(i).Cancel
            .Visible = True
        End With
        cmdAuswahl(i).Visible = False
    Next i%
    If (para.Land = "A") Then
        With nlcmdAuswahl(4)
            .Top = nlcmdAuswahl(3).Top
            .Visible = False
        End With
    End If

    With nlcmdTrägerLösung
        .Init
        .Width = nlcmdF2.Width
        .Left = nlcmdF2.Left
        .Top = flxTaxieren.Top + flxTaxieren.Height - .Height
        .Caption = cmdTrägerLösung.Caption
        .TabIndex = cmdTrägerLösung.TabIndex
        .Enabled = cmdTrägerLösung.Enabled
        .Default = cmdTrägerLösung.Default
        .Cancel = cmdTrägerLösung.Cancel
        .Visible = False
    End With
    cmdTrägerLösung.Visible = False

    With nlcmdVerwurf
        .Init
        .Width = nlcmdF2.Width
        .Left = nlcmdF2.Left
        .Top = nlcmdTrägerLösung.Top + nlcmdTrägerLösung.Height + 45
        .Caption = cmdVerwurf.Caption
        .TabIndex = cmdVerwurf.TabIndex
        .Enabled = cmdVerwurf.Enabled
        .Default = cmdVerwurf.Default
        .Cancel = cmdVerwurf.Cancel
        .Visible = False
    End With
    cmdVerwurf.Visible = False

    With cboSonderfälle
        .Left = nlcmdF2.Left
        .Width = nlcmdF2.Width
        .Top = nlcmdTrägerLösung.Top - .Height - 300
    End With
    With nlcmdSonderfälle
        .Init
        .AutoSize = True
        .Width = nlcmdF2.Width
        .Left = cboSonderfälle.Left + cboSonderfälle.Width + 90
        .Top = cboSonderfälle.Top
        .Caption = cmdSonderfälle.Caption
        .TabIndex = cmdSonderfälle.TabIndex
        .Enabled = cmdSonderfälle.Enabled
        .Default = cmdSonderfälle.Default
        .Cancel = cmdSonderfälle.Cancel
        .Visible = True
    End With
    cmdSonderfälle.Visible = False

'    Me.Width = nlcmdF2.Left + nlcmdF2.Width + 600
    Me.Width = nlcmdSonderfälle.Left + nlcmdSonderfälle.Width + 600
    Me.Height = nlcmdEsc.Top + nlcmdEsc.Height + wpara.FrmCaptionHeight + 450

'    nlcmdOk.Left = (Me.ScaleWidth - (nlcmdOk.Width + nlcmdEsc.Width + 300)) / 2
'    nlcmdEsc.Left = nlcmdOk.Left + nlcmdOk.Width + 300

    Call wpara.NewLineWindow(Me, nlcmdEsc.Top)
    With flxTaxieren
        RoundRect hdc, (.Left - iAdd) / Screen.TwipsPerPixelX, (.Top - iAdd) / Screen.TwipsPerPixelY, (.Left + .Width + iAdd) / Screen.TwipsPerPixelX, (flxTaxSumme.Top + flxTaxSumme.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
    End With
'    With flxBeleg(0)
'        RoundRect hdc, (.Left - iAdd) / Screen.TwipsPerPixelX, (.Top - iAdd) / Screen.TwipsPerPixelY, (.Left + .Width + iAdd) / Screen.TwipsPerPixelX, (.Top + .Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
'    End With

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
            ElseIf (TypeOf c Is CheckBox) Or (TypeOf c Is OptionButton) Then
                With c
                    .BackColor = GetPixel(.Container.hdc, .Left / Screen.TwipsPerPixelX - 2, .Top / Screen.TwipsPerPixelY)
                    .Height = 0
                    .Width = .Height * 3 / 4 '- 30
'                    If (.Width > 330) Then
'                        .Width = 330
'                    End If
'                    .Width = 240
                End With
                If (c.Name = "chkKassenRabatt") Then
                    If (c.index > 0) Then
                        Load lblchkKassenRabatt(c.index)
                    End If
                    With lblchkKassenRabatt(c.index)
                        .BackStyle = 0
                        .Caption = c.Caption
                        .Left = c.Left + c.Width + 60
                        .Top = c.Top
                        .Width = TextWidth(.Caption) + 90
                        .TabIndex = c.TabIndex
                        .Visible = True
                    End With
                End If
            End If
        End If
    Next
    On Error GoTo DefErr
Else
    nlcmdOk.Visible = False
    nlcmdEsc.Visible = False
    nlcmdF2.Visible = False
    nlcmdF5.Visible = False
    nlcmdF7.Visible = False
    For i = 0 To 4
        nlcmdAuswahl(i).Visible = False
    Next i
    nlcmdDarstellung.Visible = False
    nlcmdTaxmuster.Visible = False
    nlcmdTrägerLösung.Visible = False
    nlcmdVerwurf.Visible = False
End If

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

With cboSonderfälle
    .Clear
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
                .AddItem h + Space(100) + h2 + vbTab + PreisStr
            End If
        End If
    Next
'    .AddItem ("BTM-Gebühr")
'    .AddItem ("Noctu")
'    .AddItem ("Beschaffungskosten")
'    .AddItem ("Botendienst")
End With

MalFaktor# = 1.9

With frmAction.flxVerordnung
    If (.Visible) Then
        .Left = frmAction.txtRezeptNr.Left
        If (.Left + .Width > Me.Left) Then
            Me.Left = .Left + .Width + 90
        End If
    End If
End With

TxtCol% = 3
Call ShowEditBox

Call DefErrPop
End Sub

Private Sub HoleTaxMuster()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleTaxMuster")
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
Dim i%, iPrimärPackmittel%
Dim dPreis#
Dim SatzPtr&
Dim h$, h2$

With flxTaxieren
    .Visible = False
    .Rows = 1

    If (ParenteralRezept >= 0) Then
        SQLStr$ = "SELECT * FROM ParenteralTmZeilen WHERE TmId=" + CStr(FormErg)
        SQLStr$ = SQLStr$ + " ORDER BY AnfMagInd"
        Set ParenteralTmRec = RezSpeicherDB.OpenRecordset(SQLStr$)
        Do
            If (ParenteralTmRec.EOF) Then
                Exit Do
            End If
                    
            TmInhalt.pzn = PznString(ParenteralTmRec!pzn)
            
            TmInhalt.kurz = ParenteralTmRec!text
            
            TmInhalt.ActMenge = ParenteralTmRec!ActMenge
            TmInhalt.ActPreis = ParenteralTmRec!ActPreis
            TmInhalt.flag = ParenteralTmRec!flag
    
            ParEnteralPrimärPackmittel = ParenteralTmRec!packmittel
            ParEnteralAI = ParenteralTmRec!AI
            ParEnteralAnzEinheiten = ParenteralTmRec!WirkstoffMenge
            iPrimärPackmittel = ParEnteralPrimärPackmittel
            
            Call HoleTaxMusterZeile
            .AddItem " "
            Call ZeigeTaxierZeile(.Rows - 1)
            
            If (iPrimärPackmittel) Or (TmInhalt.flag = MAG_GEFAESS) Then
                If (ParEnteralAufschlag(1) > 0) Then
                    .AddItem " "
                    .row = .Rows - 1
                    
                    With TaxierRec
                        .pzn = Space$(Len(.pzn))
                        .kurz = Left$("Aufschlag " + CStr(ParEnteralAufschlag(1)) + "%" + Space$(Len(.kurz)), Len(.kurz))
                        .menge = Space$(Len(.menge))
                        .Meh = Space$(Len(.Meh))
                        .kp = 0
                        .GStufe = 0
                        
                        .ActMenge = 0#
                        .ActPreis = xVal(flxTaxieren.TextMatrix(flxTaxieren.row - 1, 0)) * ParEnteralAufschlag(1) / 100#
                        
                        .flag = MAG_PREISEINGABE
                    End With
                    Call ZeigeTaxierZeile(.Rows - 1)
                End If
            End If
            
            ParenteralTmRec.MoveNext
        Loop
        
        If (.Rows < 2) Then .AddItem " "
        .row = 1
        .Visible = True
    ElseIf (TaxmusterDBok) Then
        MalFaktor# = 1.9
        NeuMalFaktor# = MalFaktor#

        SQLStr$ = "SELECT * FROM Taxmuster WHERE Id=" + CStr(FormErg)
        Set TaxmusterRec = TaxmusterDB.OpenRecordset(SQLStr$)
        If Not (TaxmusterRec.EOF) Then
            TmHeader.ActMenge = CheckNullLong(TaxmusterRec!ActMenge)
        End If
        
        SQLStr$ = "SELECT * FROM TaxmusterZeilen WHERE TaxmusterId=" + CStr(FormErg)
        SQLStr$ = SQLStr$ + " ORDER BY LaufNr"
        Set TaxmusterRec = TaxmusterDB.OpenRecordset(SQLStr$)
        Do
            If (TaxmusterRec.EOF) Then
                Exit Do
            End If
                    
            TmInhalt.pzn = PznString(TaxmusterRec!pzn)
            
            h2 = TaxmusterRec!Name
'            Call CharToOem(h2$, h2$)
            TmInhalt.kurz = h2
                
            
            TmInhalt.ActMenge = TaxmusterRec!ActMenge
            TmInhalt.ActPreis = TaxmusterRec!ActPreis
            TmInhalt.flag = TaxmusterRec!flag
    
            Call HoleTaxMusterZeile
            .AddItem " "
            Call ZeigeTaxierZeile(.Rows - 1)
            
            TaxmusterRec.MoveNext
        Loop
        
        For i% = 1 To (.Rows - 1)
'            dPreis# = CDbl(.TextMatrix(i%, 0))
'            dPreis# = dPreis# * NeuMalFaktor# / MalFaktor#
'            .TextMatrix(i%, 0) = Format(dPreis#, "0.00")
            
            dPreis# = iCDbl(.TextMatrix(i%, 13))
            dPreis# = dPreis# * NeuMalFaktor# / MalFaktor#
            .TextMatrix(i%, 0) = Format(dPreis#, "0.00")
            .TextMatrix(i%, 13) = Format(dPreis#, "0.0000")
        Next i%
        
        MalFaktor# = NeuMalFaktor#
            
        If (.Rows < 2) Then .AddItem " "
        .row = 1
        .Visible = True
    Else
        MalFaktor# = 1.9
        NeuMalFaktor# = MalFaktor#

        Seek TM_NAMEN%, 1& * FormErg% * Len(TmHeader) + 1
        Get #TM_NAMEN%, , TmHeader
        
        If (TM_DATEN% > 0) Then
            Call HoleTaxMusterZeilen
            
            For i% = 1 To (.Rows - 1)
'                dPreis# = CDbl(.TextMatrix(i%, 0))
'                dPreis# = dPreis# * NeuMalFaktor# / MalFaktor#
'                .TextMatrix(i%, 0) = Format(dPreis#, "0.00")
                            
                dPreis# = iCDbl(.TextMatrix(i%, 13))
                dPreis# = dPreis# * NeuMalFaktor# / MalFaktor#
                .TextMatrix(i%, 0) = Format(dPreis#, "0.00")
                .TextMatrix(i%, 13) = Format(dPreis#, "0.0000")
            Next i%
            
            MalFaktor# = NeuMalFaktor#
            
            .row = 1
            .Visible = True
        End If
    End If
End With

Call TaxSumme

Call DefErrPop
End Sub

Sub ZeigeTaxierZeile(row%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeTaxierZeile")
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
Dim iFlag%
Dim h$, h2$

With TaxierRec
    iFlag% = .flag
    If (iFlag% >= MAG_NN) Then
        iFlag% = iFlag% - MAG_NN
    End If
    
    flxTaxieren.TextMatrix(row%, 0) = Format(.ActPreis, "0.00")
    flxTaxieren.TextMatrix(row%, 13) = Format(.ActPreis, "0.0000")
    
    h$ = Format(.ActMenge, "0.000")
    Do
        If (Right$(h$, 1) = ",") Then
            h$ = Left$(h$, Len(h$) - 1)
            Exit Do
        End If
        If (Right$(h$, 1) = "0") Then
            h$ = Left$(h$, Len(h$) - 1)
        Else
            Exit Do
        End If
    Loop
    flxTaxieren.TextMatrix(row%, 1) = h$
    
    flxTaxieren.TextMatrix(row%, 2) = .Meh
    
    h2$ = iTrim$(.kurz)
    If (TaxmusterDBok = 0) Then
        Call OemToChar(h2$, h2$)
    End If
    flxTaxieren.TextMatrix(row%, 3) = h2$
    
    h2$ = ""
    If (Val(.pzn) > 0) Then h2$ = .pzn
    flxTaxieren.TextMatrix(row%, 4) = h2$
    
    If (TaxierRec.Verwurf) Then
        flxTaxieren.TextMatrix(row%, 5) = "V"
    Else
        flxTaxieren.TextMatrix(row%, 5) = ""
    End If
    
    flxTaxieren.TextMatrix(row%, 6) = Format(iFlag%, "0")
    flxTaxieren.TextMatrix(row%, 7) = Format(.kp, "0.00")
    flxTaxieren.TextMatrix(row%, 8) = Format(.GStufe, "0.00")
End With
            
With flxTaxieren
    .TextMatrix(row%, 9) = CStr(Abs(ParEnteralPrimärPackmittel))
    .TextMatrix(row%, 10) = CStr(Abs(ParEnteralAI))
    .TextMatrix(row%, 11) = Format(ParEnteralAnzEinheiten, "0.00")
    
    .FillStyle = flexFillRepeat
    .row = row%
    .col = 0
    .RowSel = .row
    .ColSel = .Cols - 1
    .CellForeColor = MagDarstellung&(iFlag%, 0)
    .CellBackColor = MagDarstellung&(iFlag%, 1)
    
    If (TaxierRec.flag >= MAG_NN) Then
        .CellFontUnderline = True
    Else
        .CellFontUnderline = False
    End If
    .FillStyle = flexFillSingle
End With

ParEnteralPrimärPackmittel = 0
ParEnteralAI = 0
ParEnteralAnzEinheiten = 0

Call DefErrPop
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    Beep
'End Sub

Private Sub txtTaxieren_Change()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtTaxieren_Change")
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
Dim ind%
Dim txt$

txt$ = txtTaxieren.text
Do
    ind% = InStr(txt, ",")
    If (ind% > 0) Then
        Mid$(txt$, ind%, 1) = "."
    Else
        Exit Do
    End If
Loop
txtTaxieren.text = txt$
    
Call DefErrPop
End Sub

Private Sub txtTaxieren_KeyDown(KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtTaxieren_KeyDown")
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
Dim row%

With flxTaxieren
    row% = .row
    
#If (WINREZX = 1) Then
#Else
    If (Shift And vbAltMask) And (KeyCode = 73) Then
        With flxTaxieren
            SollPzn = .TextMatrix(.row, 4)
            SollTaxierTyp = .TextMatrix(.row, 6)
            frmMarktPzn.Show 1
        End With
    End If
#End If

    If (KeyCode = vbKeyUp) Then
        If (row% > 1) Then
            .row = row% - 1
        End If
    ElseIf (KeyCode = vbKeyDown) Then
        If (row% < (.Rows - 1)) Then
            .row = row% + 1
        End If
    ElseIf (KeyCode = vbKeyF2) Then
        cmdF2.Value = True
    ElseIf (KeyCode = vbKeyF5) Then
        cmdF5.Value = True
    ElseIf (KeyCode = vbKeyF7) Then
        cmdF7.Value = True
    End If
    
    If (row% <> .row) Then
        txtTaxieren.Visible = False
        Call MakeEditCol
    End If
End With

Call DefErrPop
End Sub

Private Sub MakeEditCol()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MakeEditCol")
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
Dim row%, iFlag%

With flxTaxieren
    row% = .row
    iFlag% = Val(.TextMatrix(row%, 6))
    
    If (iFlag% >= MAG_NN) Then
    Else
        If (.TextMatrix(row%, 3) = "") Then
            TxtCol% = 3
        ElseIf (iFlag% = MAG_PREISEINGABE) Then
            'If (.TextMatrix(row, 4) = "02567001") Then
            If (IstFiveRxPzn(.TextMatrix(row, 4))) Then
                TxtCol% = 1
            Else
                TxtCol% = 0
            End If
        Else
            TxtCol% = 1
        End If
        Call ShowEditBox
    End If
End With

Call DefErrPop
End Sub

Private Sub txtTaxieren_KeyPress(KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtTaxieren_KeyPress")
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
Dim i%, row%, ind%, erg%, OrgParenteral_AOK_LosGebiet%, OrgParenteral_AOK_NordOst%

Dim dMenge#, dPreis#
Dim pzn$, txt$, mErg$, SQLStr$, M2Nr$
Dim bHA As Boolean

If (KeyAscii = 13) Then
    With txtTaxieren
        txt$ = UCase(Trim(.text))
        If (txt$ = "") Then
            If (para.Newline) Then
                With flxTaxieren
                    If ((TxtCol% = 3) And (.row = (.Rows - 1))) Then
                        nlcmdOk.Value = True
                        Call DefErrPop: Exit Sub
                    End If
                End With
            End If
        ElseIf (txt$ <> "") Then
            Do
                ind% = InStr(txt, ".")
                If (ind% > 0) Then
                    Mid$(txt$, ind%, 1) = ","
                Else
                    Exit Do
                End If
            Loop
            
            .Visible = False
            flxTaxieren.Enabled = True
            If (TxtCol% = 0) Then
                With flxTaxieren
                    If (InStr(UCase(.TextMatrix(.row, 3)), "FIX-AUFSCHLAG") <= 0) Then
                        Call UmspeichernPreisEingabe
                        TaxierRec.ActPreis = CDbl(txt$) * MalFaktor#
                    End If
                End With
                Call ZeigeTaxierZeile(flxTaxieren.row)
                
                Call TaxSumme
                
                row% = flxTaxieren.row
                If (row% = flxTaxieren.Rows - 1) Then
                    flxTaxieren.AddItem " "
                    flxTaxieren.row = flxTaxieren.Rows - 1
                    TxtCol% = 3
                    Call ShowEditBox
                Else
                    Call txtTaxieren_KeyDown(vbKeyDown, 0)
                End If
            ElseIf (TxtCol% = 1) Then
                row% = flxTaxieren.row
                
                If (Left$(txt$, 2) = "AD") Then
                    flxTaxieren.TextMatrix(row%, 1) = "0"
                    Call TaxSumme
                    
                    dMenge# = iCDbl(Mid$(txt$, 3))
                    txt$ = "0"
                    dMenge# = dMenge# - (TeilMenge# - iCDbl(flxTaxieren.TextMatrix(row%, 1)))
                    If (dMenge# > 0) Then
                        txt$ = Format(dMenge#, "0.000")
                    End If
                    flxTaxieren.TextMatrix(row%, 1) = txt$
                    Call ShowEditBox
                    Call TaxSumme
                    Call DefErrPop: Exit Sub
                End If
                              
                flxTaxieren.TextMatrix(row%, 1) = txt$
                TaxierRec.ActMenge = iCDbl(flxTaxieren.TextMatrix(row%, 1))
                TaxierRec.Meh = flxTaxieren.TextMatrix(row%, 2)
                TaxierRec.kurz = flxTaxieren.TextMatrix(row%, 3)
                TaxierRec.pzn = flxTaxieren.TextMatrix(row%, 4)
                TaxierRec.flag = Val(flxTaxieren.TextMatrix(row%, 6))
                TaxierRec.kp = iCDbl(flxTaxieren.TextMatrix(row%, 7))
                TaxierRec.GStufe = iCDbl(flxTaxieren.TextMatrix(row%, 8))
                TaxierRec.Verwurf = iCDbl(flxTaxieren.TextMatrix(row%, 12))
                
                TaxierTyp = TaxierRec.flag
                
                
                ParEnteralPrimärPackmittel = 0
                ParEnteralAI = 0
                ParEnteralAnzEinheiten = 0
                If (TaxierTyp = MAG_SPEZIALITAET) Then
'                    If (AnzParenteralSpez > 0) Then
'                        Call ParenteralSpezAuswahl(Val(txt))
'                        Call DefErrPop: Exit Sub
'                    End If
    '                If (TaxierRec.ActMenge <= TaxierRec.gstufe) Then
                    Call PreisSpezialitaet
                
                    If (ParenteralRezept > 15) Then
                        If ((TaxierRec.ActMenge / TaxierRec.GStufe) >= 100) Then
                            Call MessageBox("Achtung: Wert '" + CStr(TaxierRec.ActMenge) + "' ist für die Hashcode-Ermittlung zu groß !" + vbCrLf + vbCrLf + "Er muss unter '" + CStr(TaxierRec.GStufe * 100) + "' liegen !", vbInformation)
                        End If
                    End If

                    
                    If (ParEnteralPrimärPackmittel) Then
                        If (ParEnteralAufschlag(1) > 0) Then
                            Call ZeigeTaxierZeile(row%)
                        
                            flxTaxieren.AddItem " "
                            flxTaxieren.row = flxTaxieren.Rows - 1
                            row% = flxTaxieren.row
                            
                            With TaxierRec
                                .pzn = Space$(Len(.pzn))
                                .kurz = Left$("Aufschlag " + CStr(ParEnteralAufschlag(1)) + "%" + Space$(Len(.kurz)), Len(.kurz))
                                .menge = Space$(Len(.menge))
                                .Meh = Space$(Len(.Meh))
                                .kp = 0
                                .GStufe = 0
                                
                                .ActMenge = 0#
                                .ActPreis = xVal(flxTaxieren.TextMatrix(row - 1, 0)) * ParEnteralAufschlag(1) / 100#
                                
                                .flag = MAG_PREISEINGABE
                            End With
                        End If
                    ElseIf (Parenteral_AOK_LosGebiet Or Parenteral_AOK_NordOst) And (TaxierRec.ActPreis = 0) Then
                        With flxTaxieren
                            For i = 0 To (.Cols - 1)
                                .TextMatrix(row, i) = ""
                            Next i
                        End With
                        TxtCol% = 3
                        Call MakeEditCol
                        Call TaxSumme
                        Call DefErrPop: Exit Sub
                    End If
                ElseIf (TaxierTyp = MAG_HILFSTAXE) Then
                    Call PreisHilfsTaxe
                ElseIf (TaxierTyp = MAG_SONSTIGES) Then
                    Call PreisSonstiges
                ElseIf (TaxierTyp = MAG_ANTEILIG) Then
                    Call PreisAnteilig
                ElseIf (TaxierTyp = MAG_PREISEINGABE) Then
'                    flxTaxieren.TextMatrix(row%, 0) = Format(3.58 * Val(txt), "0.00")
                    flxTaxieren.TextMatrix(row%, 0) = Format(xVal(flxTaxieren.TextMatrix(row%, 7)) * Val(txt), "0.00")
                    Call TaxSumme
                    
                    row% = flxTaxieren.row
                    If (row% = flxTaxieren.Rows - 1) Then
                        flxTaxieren.AddItem " "
                        flxTaxieren.row = flxTaxieren.Rows - 1
                        TxtCol% = 3
                        Call ShowEditBox
                    Else
                        Call txtTaxieren_KeyDown(vbKeyDown, 0)
                    End If
                    Call DefErrPop: Exit Sub
                Else
                    OrgParenteral_AOK_LosGebiet = Parenteral_AOK_LosGebiet
                    OrgParenteral_AOK_NordOst = Parenteral_AOK_NordOst
                    SollMenge# = TaxierRec.ActMenge
                    erg% = AuswahlArbEmb%
            
                    If Not (OrgParenteral_AOK_LosGebiet) And (Parenteral_AOK_LosGebiet) Then
                        erg% = AuswahlArbEmb%
                    ElseIf Not (OrgParenteral_AOK_NordOst) And (Parenteral_AOK_NordOst) Then
                        erg% = AuswahlArbEmb%
                    End If
        
                    If (erg% = False) Then
                        Call ShowEditBox
                        Call DefErrPop: Exit Sub
                    ElseIf (TaxierRec.flag = MAG_ARBEIT) Then
                        Call ZeigeTaxierZeile(row%)
                        
                        For i% = 1 To (flxTaxieren.Rows - 1)
'                            dPreis# = iCDbl(flxTaxieren.TextMatrix(i%, 0))
                            dPreis# = iCDbl(flxTaxieren.TextMatrix(i%, 13))
                            dPreis# = dPreis# * NeuMalFaktor# / MalFaktor#
                            flxTaxieren.TextMatrix(i%, 0) = Format(dPreis#, "0.00")
                            flxTaxieren.TextMatrix(i%, 13) = Format(dPreis#, "0.0000")
                        Next i%
                        MalFaktor# = NeuMalFaktor#
                    ElseIf (TaxierRec.flag = MAG_GEFAESS) Then
                    
                        bHA = False
#If (WINREZX = 1) Then
#Else
                        SQLStr = "Select count(*) as iAnz FROM GRP_HA "
                        SQLStr = SQLStr & " WHERE PZN_2=" + TaxierRec.pzn
                        On Error Resume Next
                        ABDA_Komplett_Rec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        ABDA_Komplett_Rec.Open SQLStr, ABDA_Komplett_Conn
                        If (ABDA_Komplett_Rec.EOF = False) Then
                            bHA = (CheckNullInt(ABDA_Komplett_Rec!iAnz) > 0)
                        End If
                        ABDA_Komplett_Rec.Close
                        
                        If (bHA) Then
                            SollTaxierTyp = MAG_SPEZIALITAET
                            SollPzn = TaxierRec.pzn
                            frmAuswahlHA.Show 1
                            
                            If (FormErgTxt <> "") Then
                                pzn = FormErgTxt
                                
                                If (ArtikelDbOk) Then
                                    SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + pzn
                '                    SQLStr = SQLStr + " AND LagerKz<>0"
                                    FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr)
                                Else
                                    FabsErrf% = ast.IndexSearch(0, pzn, FabsRecno&)
                                    If (FabsErrf% = 0) Then
                                        ast.GetRecord (FabsRecno& + 1)
                                    End If
                                End If
                                If (FabsErrf% = 0) Then
        '                            ast.GetRecord (FabsRecno& + 1)
                                Else
                                    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
                                    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                                    On Error Resume Next
                                    TaxeRec.Close
                                    Err.Clear
                                    On Error GoTo DefErr
                                    TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                                    If (TaxeRec.EOF = False) Then
                                        Call Taxe2ast(pzn$)
                                        FabsErrf% = 0
                                    End If
                                End If
                                    
                                If (FabsErrf% = 0) Then
                                    Call UmspeichernSpezialitaet(MAG_GEFAESS)
                                    TaxierRec.ActPreis = TaxierRec.kp
                                    TaxierRec.ActMenge = TaxierRec.GStufe
                                    If (ParenteralRezept = 24) Or (ParenteralRezept = 26) Or (ParenteralRezept = 28) Or (ParenteralRezept = 30) Then
                                        TaxierRec.ActPreis = TaxierRec.ActPreis * 1.9
                                    ElseIf (ParenteralRezept = 25) Or (ParenteralRezept = 27) Or (ParenteralRezept = 29) Then
                                        TaxierRec.ActPreis = TaxierRec.ActPreis * 2
                                    Else
                                        TaxierRec.ActPreis = TaxierRec.ActPreis * MalFaktor#
                                    End If
                                End If
                            End If
                        End If
#End If
                    
                        If (TaxierRec.flag = MAG_GEFAESS) And (ParenteralRezept >= 0) And (ParenteralRezept <= 15) Then
                            If (ParEnteralAufschlag(1) > 0) Then
                                Call ZeigeTaxierZeile(row%)
                            
                                flxTaxieren.AddItem " "
                                flxTaxieren.row = flxTaxieren.Rows - 1
                                row% = flxTaxieren.row
                                
                                With TaxierRec
                                    .pzn = Space$(Len(.pzn))
                                    .kurz = Left$("Aufschlag " + CStr(ParEnteralAufschlag(1)) + "%" + Space$(Len(.kurz)), Len(.kurz))
                                    .menge = Space$(Len(.menge))
                                    .Meh = Space$(Len(.Meh))
                                    .kp = 0
                                    .GStufe = 0
                                    
                                    .ActMenge = 0#
                                    .ActPreis = xVal(flxTaxieren.TextMatrix(row - 1, 0)) * ParEnteralAufschlag(1) / 100#
                                    
                                    .flag = MAG_PREISEINGABE
                                End With
                            End If
                        End If
                    End If
                End If
                
                Call ZeigeTaxierZeile(row%)
                If (ParenteralRezept > 15) Then
                    If (TaxierTyp = MAG_SPEZIALITAET) Or (TaxierTyp = MAG_HILFSTAXE) Then
                        If (ParenteralRezept > 23) Then
                            Call CannabisZuschlag
                        Else
#If (WINREZX = 1) Then
#Else
                            Call SubstitutionZuschlag
#End If
                            
                            
                        End If
                    End If
                End If
                Call TaxSumme
                
                row% = flxTaxieren.row
                If (row% = flxTaxieren.Rows - 1) Then
                    flxTaxieren.AddItem " "
                    flxTaxieren.row = flxTaxieren.Rows - 1
                    TxtCol% = 3
                    Call ShowEditBox
                Else
                    Call txtTaxieren_KeyDown(vbKeyDown, 0)
                End If
                
            ElseIf (TxtCol% = 3) Then
                Call PruefeEingabe(txt$)
                If (TaxierTyp = MAG_GEFAESS) Or (TaxierTyp = MAG_ARBEIT) Or (TaxierTyp = MAG_SONSTIGES) Then
                    TaxierRec.pzn = Space$(Len(TaxierRec.pzn))
                    TaxierRec.kurz = Space$(Len(TaxierRec.kurz))
                    TaxierRec.menge = Space$(Len(TaxierRec.menge))
                    TaxierRec.Meh = Space$(Len(TaxierRec.Meh))
                    TaxierRec.kp = 0
                    TaxierRec.GStufe = 0
                    TaxierRec.ActMenge = TeilMenge#
                    TaxierRec.ActPreis = 0
                    TaxierRec.flag = TaxierTyp
                    TaxierRec.Verwurf = 0
                    
                    Call ZeigeTaxierZeile(flxTaxieren.row)
                    
                    If (TaxierTyp = MAG_SONSTIGES) Then
                        erg% = AuswahlArbEmb%
                        If (erg% = False) Then
                            Call ShowEditBox
                            Call DefErrPop: Exit Sub
                        End If
                        Call ZeigeTaxierZeile(flxTaxieren.row)
                    End If
                    TxtCol% = 1
                    Call ShowEditBox
                ElseIf (TaxierTyp = MAG_SPEZIALITAET) Or (TaxierTyp = MAG_ANTEILIG) Then
                    mErg$ = Matchcode(0, pzn$, txt$, False, True)
                    If (mErg$ <> "") Then
                        ind% = InStr(mErg$, "@")
                        If (ind > 0) Then
                            mErg$ = Left$(mErg$, ind% - 1)
                        End If
                        pzn$ = mErg$
                        
                        
                        If (ArtikelDbOk) Then
                            SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + pzn
        '                    SQLStr = SQLStr + " AND LagerKz<>0"
                            FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr)
                        Else
                            FabsErrf% = ast.IndexSearch(0, pzn, FabsRecno&)
                            If (FabsErrf% = 0) Then
                                ast.GetRecord (FabsRecno& + 1)
                            End If
                        End If
                        If (FabsErrf% = 0) Then
'                            ast.GetRecord (FabsRecno& + 1)
                        Else
                            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
                            'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                            On Error Resume Next
                            TaxeRec.Close
                            Err.Clear
                            On Error GoTo DefErr
                            TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                            If (TaxeRec.EOF = False) Then
                                Call Taxe2ast(pzn$)
                                FabsErrf% = 0
                            End If
                        End If
                            
                        If (FabsErrf% = 0) Then
                            AnzParenteralSpez = 0
                            
                            Call UmspeichernSpezialitaet(TaxierTyp)
                            TaxierRec.ActMenge = TaxierRec.GStufe
                            Call ZeigeTaxierZeile(flxTaxieren.row)
                            
'                            If (ParenteralRezept >= 0) Then
'                                Call ParenteralSpezAktiv(pzn$)
'                            End If
                            If (ParenteralRezept >= 0) Then
                                If (ast.rez = "SG") Then
                                    Call CheckBtmGebuehr
                                End If
                            End If
            
                            TxtCol% = 1
                        End If
                    End If
                    Call ShowEditBox
                ElseIf (TaxierTyp = MAG_SONSTIGES) Then
                ElseIf (TaxierTyp = MAG_HILFSTAXE) Then
                    mErg$ = Matchcode(3, pzn$, txt$, txt$ <> "", True)
                    If (mErg$ <> "") Then
            '            ind% = InStr(mErg$, "@")
            '            pzn$ = Left$(mErg$, ind% - 1)
                        pzn$ = mErg$
                    
                        bHA = False
#If (WINREZX = 1) Then
#Else
                        SQLStr = "Select count(*) as iAnz FROM GRP_HA "
                        SQLStr = SQLStr & " WHERE PZN_2=" + pzn
                        On Error Resume Next
                        ABDA_Komplett_Rec.Close
                        Err.Clear
                        On Error GoTo DefErr
                        ABDA_Komplett_Rec.Open SQLStr, ABDA_Komplett_Conn
                        If (ABDA_Komplett_Rec.EOF = False) Then
                            bHA = (CheckNullInt(ABDA_Komplett_Rec!iAnz) > 0)
                        End If
                        ABDA_Komplett_Rec.Close
                        
                        If (bHA) Then
                            SollTaxierTyp = MAG_SPEZIALITAET
                            SollPzn = pzn
                            frmAuswahlHA.Show 1
                            
                            If (FormErgTxt <> "") Then
                                pzn = FormErgTxt
                                If (ArtikelDbOk) Then
                                    SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + pzn
                '                    SQLStr = SQLStr + " AND LagerKz<>0"
                                    FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr)
                                Else
                                    FabsErrf% = ast.IndexSearch(0, pzn, FabsRecno&)
                                    If (FabsErrf% = 0) Then
                                        ast.GetRecord (FabsRecno& + 1)
                                    End If
                                End If
                                If (FabsErrf% = 0) Then
        '                            ast.GetRecord (FabsRecno& + 1)
                                Else
                                    SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
                                    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
                                    On Error Resume Next
                                    TaxeRec.Close
                                    Err.Clear
                                    On Error GoTo DefErr
                                    TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
                                    If (TaxeRec.EOF = False) Then
                                        Call Taxe2ast(pzn$)
                                        FabsErrf% = 0
                                    End If
                                End If
                                    
                                If (FabsErrf% = 0) Then
                                    AnzParenteralSpez = 0
' bis 6.02
'                                    Call UmspeichernSpezialitaet(MAG_ANTEILIG)
' ab 6.03
                                    Call UmspeichernSpezialitaet(IIf(ParenteralRezept > 23, MAG_SPEZIALITAET, MAG_ANTEILIG))
                                    TaxierRec.ActMenge = TaxierRec.GStufe
                                    Call ZeigeTaxierZeile(flxTaxieren.row)
                                    
        '                            If (ParenteralRezept >= 0) Then
        '                                Call ParenteralSpezAktiv(pzn$)
        '                            End If
                                    If (ParenteralRezept >= 0) Then
                                        If (ast.rez = "SG") Then
                                            Call CheckBtmGebuehr
                                        End If
                                    End If
                    
                                    TxtCol% = 1
                                End If
                            Else
                            End If
                            Call ShowEditBox
                            Call DefErrPop: Exit Sub
                        End If
#End If
                        If (ArtikelDbOk) Then
                            If (Val(pzn$) < 0) Then
                                SQLStr$ = "SELECT * FROM Artikel WHERE Id = " + CStr(Abs(Val(pzn)))
                            Else
                                SQLStr$ = "SELECT * FROM Artikel WHERE Pzn = " + pzn
                            End If
                            FabsErrf = Hilfstaxe.OpenRecordset(HilfstaxeRec, SQLStr)
                            If (FabsErrf% = 0) Then
'                                Call hTaxe.GetRecord(FabsRecno& + 1)
                                Call UmspeichernHilfsTaxe
                                TaxierRec.ActMenge = TaxierRec.GStufe
                                Call ZeigeTaxierZeile(flxTaxieren.row)
                
                                TxtCol% = 1
                            End If
                        Else
                            If (Val(pzn$) < 0) Then
                                FabsRecno& = -Val(pzn$)
                                FabsErrf% = 0
                            Else
                                FabsErrf% = hTaxe.IndexSearch(0, pzn$, FabsRecno&)
                            End If
                            If (FabsErrf% = 0) Then
                                Call hTaxe.GetRecord(FabsRecno& + 1)
                                Call UmspeichernHilfsTaxe
                                TaxierRec.ActMenge = TaxierRec.GStufe
                                Call ZeigeTaxierZeile(flxTaxieren.row)
                
                                TxtCol% = 1
                            End If
                        End If
                    End If
                    Call ShowEditBox
                ElseIf (TaxierTyp = MAG_PREISEINGABE) Then
                    Call UmspeichernPreisEingabe
                    TaxierRec.ActPreis = CDbl(txt$) * MalFaktor#
                    Call ZeigeTaxierZeile(flxTaxieren.row)
                    
                    Call TaxSumme
                    
                    row% = flxTaxieren.row
                    If (row% = flxTaxieren.Rows - 1) Then
                        flxTaxieren.AddItem " "
                        flxTaxieren.row = flxTaxieren.Rows - 1
                        TxtCol% = 3
                        Call ShowEditBox
                    Else
                        Call txtTaxieren_KeyDown(vbKeyDown, 0)
                    End If
                End If
            End If
        End If
    End With
Else
    If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
    If (TxtCol% < 3) Then
        If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And (Chr$(KeyAscii) <> ".") And ((TxtCol% = 0) Or (InStr("ADad", Chr$(KeyAscii)) = 0)) Then
            Beep
            KeyAscii = 0
        End If
    End If
End If

Call DefErrPop
End Sub

'Sub ParenteralSpezAktiv(pzn$)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("ParenteralSpezAktiv")
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
'Dim i%, j%, l%, hTab%, row%, col%, ind%, aRow%, wi%, Anlage3Ok%
'Dim s$, h$, BetrLief$, Lief2$
'Dim M2Nr$
'Dim Anlage3DB As Database
'Dim Anlage3Rec As Recordset
'
'col = 2
'
'Anlage3Ok = (Dir("Anlage3.mdb") <> "")
'If (Anlage3Ok) Then
'    Set Anlage3DB = OpenDatabase("Anlage3.mdb", False, True)
'End If
'
'For i = 0 To UBound(ParenteralSpez)
'    ParenteralSpez(i).flag = 0
'    ParenteralEk#(i) = 0
'    ParenteralPreisProEinheit#(i) = 0
'    ParenteralAnzEinheiten#(i) = 0
'Next i
'
'i = 0
'M2Nr = ""
'SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
'If (TaxeRec.EOF = False) Then
'    M2Nr = CheckNullStr(TaxeRec!M2)
'End If
'If (M2Nr <> "") Then
'    SQLStr$ = "SELECT * FROM TAXE WHERE M2 =""" + M2Nr + """"
'    Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
'    Do
'        If (TaxeRec.EOF) Then
'            Exit Do
'        End If
'
''        FabsErrf% = ast.IndexSearch(0, Format(TaxeRec!pzn, "0000000"), FabsRecno&)
''        If (FabsErrf% = 0) Then
''            ast.GetRecord (FabsRecno& + 1)
''            Call UmspeichernSpezialitaet(MAG_SPEZIALITAET)
''            ParenteralSpez(i) = TaxierRec
''            ParenteralSpez(i).flag = 0
'''            TaxierRec.ActMenge = TaxierRec.Gstufe
''            i = i + 1
''        End If
'        Call Taxe2ast(Format(TaxeRec!pzn, "0000000"))
'        Call UmspeichernSpezialitaet(MAG_SPEZIALITAET)
'        ParenteralSpez(i) = TaxierRec
'        ParenteralSpez(i).flag = 0
'
'        SQLStr$ = "SELECT * FROM Artikel WHERE PZN = " + pzn$
'        Set Anlage3Rec = Anlage3DB.OpenRecordset(SQLStr$)
'        If Not (Anlage3Rec.EOF) Then
'            ParenteralEk#(i) = Anlage3Rec!ApoEk
'            ParenteralPreisProEinheit#(i) = Anlage3Rec!Wert_Günstigster_2Günstigster
'            ParenteralAnzEinheiten#(i) = Anlage3Rec!StoffmengeProPackung
'            ParenteralSpez(i).Kp = ParenteralEk#(i)
'        End If
'
'        i = i + 1
'
'        TaxeRec.MoveNext
'    Loop
'End If
'If (i > 0) Then
'    Load frmEdit
'    With frmEdit.flxMultiEdit
'        .Rows = 10
'        .Cols = 7
'        .ColWidth(0) = frmEdit.TextWidth(String(2, "X"))
'        .ColWidth(1) = flxTaxieren.ColWidth(0) - .ColWidth(0)
'        For j = 2 To 6
'            .ColWidth(j) = flxTaxieren.ColWidth(j - 1)
'        Next j
'
'        .ColAlignment(2) = flexAlignRightCenter
'
'        .Width = flxTaxieren.Width
'        .Height = .RowHeight(0) * .Rows + 90
'
'        .Rows = 0
'        For j = 0 To (i - 1)
'            With ParenteralSpez(j)
'                frmEdit.flxMultiEdit.AddItem Chr$(214) + vbTab + Format(.Kp, "0.00") + vbTab + .menge + vbTab + .meh + vbTab + .kurz + vbTab + .pzn
'                FabsErrf% = ast.IndexSearch(0, .pzn, FabsRecno&)
'            End With
'            If (FabsErrf% = 0) Then
'                .FillStyle = flexFillRepeat
'                .row = .Rows - 1
'                .col = 1
'                .RowSel = .row
'                .ColSel = .Cols - 1
'                .CellFontBold = True
'                .FillStyle = flexFillSingle
'            End If
'            If (ParenteralEk#(j) = 0) Then
'                .FillStyle = flexFillRepeat
'                .row = .Rows - 1
'                .col = 1
'                .RowSel = .row
'                .ColSel = .Cols - 1
'                .CellForeColor = vbRed
'                .CellFontItalic = True
'                .FillStyle = flexFillSingle
'            End If
'        Next j
'        .Height = .RowHeight(0) * .Rows + 90
'
'        .FillStyle = flexFillRepeat
'        .row = 0
'        .col = 0
'        .RowSel = .Rows - 1
'        .ColSel = .col
'        .CellFontName = "Symbol"
'        .FillStyle = flexFillSingle
'    End With
'
'    With frmEdit
'        .Left = flxTaxieren.Left '+ 45 '+ flxTaxieren.ColPos(col%) + 45
'        .Left = .Left + Me.Left + wpara.FrmBorderHeight
'        .Top = flxTaxieren.Top + flxTaxieren.RowPos(flxTaxieren.row)
'        .Top = .Top + Me.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
'        .Width = flxTaxieren.ColWidth(col%)
'        .Width = flxTaxieren.Width '- flxTaxieren.ColPos(col%) + 45
'        .Height = flxTaxieren.Height - flxTaxieren.RowHeight(0)
'    End With
'    With frmEdit.flxMultiEdit
''        .Height = frmEdit.ScaleHeight
''        frmEdit.Height = .Height
''        .Width = frmEdit.ScaleWidth
'        .Left = 0
'        .Top = 0
'        .row = 0
'        .col = 0
'        .ColSel = .Cols - 1
'
'        .Visible = True
'    End With
'
'    frmEdit.Show 1
'
'    If (EditErg%) Then
'        AnzParenteralSpez = EditAnzGefunden
'        For i = 0 To (EditAnzGefunden - 1)
'            ParenteralSpez(EditGef%(i)).flag = 1
'        Next i
'    End If
'End If
'
'If (Anlage3Ok) Then
'    Anlage3DB.Close
'End If
'
'Call DefErrPop
'End Sub
'
'Sub ParenteralSpezAuswahl(SollMenge%)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("ParenteralSpezAuswahl")
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
'Dim i%, j%, k%, l%, hTab%, row%, col%, ind%, aRow%
'Dim s$, h$, BetrLief$, Lief2$
'Dim M2Nr$
'Dim MaxAnz%(10), IstMenge%, lauf0%, lauf1%, lauf2%, lauf3%, lauf4%, KleinstLauf%(10), KleinstMenge%
'Dim IstPreis#, KleinstPreis#
'
'For i = 0 To UBound(ParenteralSpez)
'    MaxAnz(i) = 0
'    With ParenteralSpez(i)
'        If (.flag = 1) Then
'            MaxAnz(i) = SollMenge \ .Gstufe + 1
'        End If
'    End With
'Next i
'
'KleinstPreis = 99999#
'For lauf0 = 0 To MaxAnz(0)
'    For lauf1 = 0 To MaxAnz(1)
'        For lauf2 = 0 To MaxAnz(2)
'            For lauf3 = 0 To MaxAnz(3)
'                For lauf4 = 0 To MaxAnz(4)
'                    IstMenge = 0
'                    IstMenge = IstMenge + lauf0 * ParenteralSpez(0).Gstufe
'                    IstMenge = IstMenge + lauf1 * ParenteralSpez(1).Gstufe
'                    IstMenge = IstMenge + lauf2 * ParenteralSpez(2).Gstufe
'                    IstMenge = IstMenge + lauf3 * ParenteralSpez(3).Gstufe
'                    IstMenge = IstMenge + lauf4 * ParenteralSpez(4).Gstufe
''                    For i = 0 To 4
''                        IstMenge = IstMenge + Lauf(i) * ParenteralSpez(i).Gstufe
''                    Next i
'                    If (IstMenge >= SollMenge) Then
'                        IstPreis = 0
'                        IstPreis = IstPreis + lauf0 * ParenteralSpez(0).Kp
'                        IstPreis = IstPreis + lauf1 * ParenteralSpez(1).Kp
'                        IstPreis = IstPreis + lauf2 * ParenteralSpez(2).Kp
'                        IstPreis = IstPreis + lauf3 * ParenteralSpez(3).Kp
'                        IstPreis = IstPreis + lauf4 * ParenteralSpez(4).Kp
''                        For i = 0 To 4
''                            IstPreis = IstPreis + Lauf(i) * ParenteralSpez(i).Kp
''                        Next i
'                        If (IstPreis < KleinstPreis) Then
'                            KleinstMenge = IstMenge
'                            KleinstPreis = IstPreis
'                            KleinstLauf(0) = lauf0
'                            KleinstLauf(1) = lauf1
'                            KleinstLauf(2) = lauf2
'                            KleinstLauf(3) = lauf3
'                            KleinstLauf(4) = lauf4
''                            For i = 0 To 4
''                                KleinstLauf(i) = Lauf(i)
''                            Next i
'                        End If
'                    End If
'                Next lauf4
'            Next lauf3
'        Next lauf2
'    Next lauf1
'Next lauf0
'
'h = Str(KleinstPreis) + Str(KleinstMenge) + vbCrLf
'For i% = 0 To 4
'    h = h + Str(KleinstLauf(i)) + Str(ParenteralSpez(i).Gstufe) + vbCrLf
'Next i
''Call MsgBox(h)
'
'j = 0
'With flxTaxieren
'    For i% = 0 To 4
'        For k = 1 To KleinstLauf(i)
'            If (j > 0) Then
'                .AddItem " "
'                .row = .Rows - 1
'            End If
'            j = j + 1
'
'            TaxierRec = ParenteralSpez(i)
'            TaxierRec.flag = MAG_SPEZIALITAET
'            TaxierRec.ActMenge = TaxierRec.Gstufe
'            If (j = 1) Then
'                TaxierRec.ActMenge = TaxierRec.ActMenge - (KleinstMenge - SollMenge)
'            End If
'            Call ZeigeTaxierZeile(.row)
'            Call PreisSpezialitaet(ParenteralEk#(i), ParenteralPreisProEinheit#(i), ParenteralAnzEinheiten#(i))
'            Call ZeigeTaxierZeile(.row)
'
'            If (ParEnteralAufschlag(0) > 0) Then
'                .AddItem " "
'                .row = .Rows - 1
'
'                With TaxierRec
'                    .pzn = Space$(Len(.pzn))
'                    .kurz = Left$("Aufschlag " + CStr(ParEnteralAufschlag(0)) + "%" + Space$(Len(.kurz)), Len(.kurz))
'                    .menge = Space$(Len(.menge))
'                    .meh = Space$(Len(.meh))
'                    .Kp = 0
'                    .Gstufe = 0
'
'                    .ActMenge = 0#
'                    .ActPreis = xVal(flxTaxieren.TextMatrix(flxTaxieren.row - 1, 0)) * ParEnteralAufschlag(0) / 100#
'
'                    .flag = MAG_PREISEINGABE
'                End With
'                Call ZeigeTaxierZeile(.row)
'            End If
'        Next k
'    Next i
'
'    Call TaxSumme
'
'    .AddItem " "
'    .row = .Rows - 1
'    TxtCol% = 3
'    Call ShowEditBox
'End With
'
'Call DefErrPop
'End Sub

Private Sub txtTaxieren_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtTaxieren_GotFocus")
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

With txtTaxieren
    .SelStart = 0
    .SelLength = Len(.text)
End With

Call DefErrPop
End Sub

Private Sub ShowEditBox()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ShowEditBox")
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
Dim i%, F2Status%, F7Status%
Dim iTop%, iLeft%, iWidth%, iHeight%, ind%
Dim h$, txt$

With flxTaxieren
'    TxtRow% = .row
    
'    If (NoEditBox%) Then
'        Call DefErrPop: Exit Sub
'    End If

    If (.row < .TopRow) Then
        .TopRow = .row
    Else
        If (.TopRow + iArbeitAnzzeilen% - 1 <= .row) Then
            .TopRow = .row
        End If
    End If
    
    iLeft% = .Left + .ColPos(TxtCol%)
    If (para.Newline = 0) Then
        iLeft = iLeft + 30
    End If
    iWidth% = .ColWidth(TxtCol%)
    iTop% = .Top + .RowHeight(0) * (.row - .TopRow + 1)
    iHeight% = .RowHeight(.row)

    txt$ = .TextMatrix(.row, TxtCol%)
    If (TxtCol% = 0) Then
        txt$ = Format(txt$ / MalFaktor#, "0.00")
    End If
End With

With txtTaxieren
    .Left = iLeft%
    .Top = iTop%
    .Width = iWidth%
    .Height = iHeight%
    .MaxLength = 50 ' wird danach übersteuert
    .text = txt$
    .Visible = True
    If (.Visible) Then .SetFocus
End With

flxTaxieren.Enabled = False

If (TxtCol% = 3) Then
    F2Status% = False
Else
    F2Status% = True
End If
cmdF2.Enabled = F2Status%
nlcmdF2.Enabled = F2Status%

For i% = 0 To 4
    cmdAuswahl(i%).Enabled = (F2Status% = False)
    nlcmdAuswahl(i%).Enabled = (F2Status% = False)
Next i%

If (TxtCol% = 1) Then
    F7Status% = False
Else
    F7Status% = True
End If
cmdF7.Enabled = F7Status%
nlcmdF7.Enabled = F7Status%

Call DefErrPop
End Sub

'Private Sub ShowEditBox()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("ShowEditBox")
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
'Dim i%, F2Status%, F7Status%
'Dim h$
'
'With txtTaxieren
'    .Left = flxTaxieren.Left + flxTaxieren.ColPos(TxtCol%) + 30
'    .Width = flxTaxieren.ColWidth(TxtCol%)
'    .Top = flxTaxieren.Top + flxTaxieren.RowPos(flxTaxieren.row)
'    .Height = flxTaxieren.RowHeight(flxTaxieren.row)
'
'    h$ = flxTaxieren.TextMatrix(flxTaxieren.row, TxtCol%)
'    If (TxtCol% = 0) Then
'        .text = Format(h$ / MalFaktor#, "0.00")
'    Else
'        .text = h$
'    End If
'
'    .Visible = True
'    If (.Visible) Then .SetFocus
'End With
'
'flxTaxieren.Enabled = False
'
'If (TxtCol% = 3) Then
'    F2Status% = False
'Else
'    F2Status% = True
'End If
'cmdF2.Enabled = F2Status%
'
'For i% = 0 To 4
'    cmdAuswahl(i%).Enabled = (F2Status% = False)
'Next i%
'
'If (TxtCol% = 1) Then
'    F7Status% = False
'Else
'    F7Status% = True
'End If
'cmdF7.Enabled = F7Status%
'
'Call DefErrPop
'End Sub

Private Sub PruefeEingabe(InpTxt$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("PruefeEingabe")
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
Dim i%, j%, l%, AnzMulti%, row%, iFlag%
Dim lWert#
Dim txt$, txt2$, ch$, SQLStr$

TaxierTyp = 99

txt$ = InpTxt$

If (Len(txt$) = 1) Then
    If (txt$ >= "0") And (txt$ <= "4") Then
        TaxierTyp = Val(txt$)
        If (TaxierTyp = 4) Then TaxierTyp = MAG_ANTEILIG
        InpTxt$ = ""
        Call DefErrPop: Exit Sub
    End If
ElseIf (Left$(txt$, 11) = "99999999997") Then
    TaxierTyp = Val(Mid$(txt$, 12, 1))
    Call DefErrPop: Exit Sub
End If

l% = Len(txt$)
row% = flxTaxieren.row
If (Right$(txt$, 1) = "*") And (row% > 1) Then
    txt$ = Left$(txt$, l% - 1)
    AnzMulti% = Val(txt$) - 1
    If (AnzMulti% > 0) Then
        With flxTaxieren
            iFlag% = Val(.TextMatrix(row% - 1, 6))
            For i% = 1 To AnzMulti%
                .AddItem "", row%
                For j% = 0 To (.Cols - 1)
                    .TextMatrix(row%, j%) = .TextMatrix(row% - 1, j%)
                Next j%
            Next i%
            
            .FillStyle = flexFillRepeat
            .row = row%
            .col = 0
            .RowSel = .row + AnzMulti% - 1
            .ColSel = .Cols - 1
            .CellForeColor = MagDarstellung&(iFlag%, 0)
            .CellBackColor = MagDarstellung&(iFlag%, 1)
            
            .FillStyle = flexFillSingle
            
            .row = row% + AnzMulti%
        End With
    
        Call TaxSumme
        TxtCol% = 3
        Call ShowEditBox
        Call DefErrPop: Exit Sub
    End If
End If

If (l% = 9) And (Left$(txt$, 1) = "-") Then txt$ = Mid$(txt$, 2)
If ((l% >= 10) And (((Left$(txt$, 1) = "*") And (Mid$(txt$, l%, 1) = "*")) Or (InStr(txt$, "PZN") > 0))) Then
    txt2$ = ""
    For i% = 1 To l%
        ch$ = Mid$(txt$, i%, 1)
        If (InStr("0123456789", ch$) > 0) Then
            txt2$ = txt2$ + ch$
        End If
    Next i%
    txt$ = txt2$
End If

l% = Len(txt$)

For i% = 1 To l%
    ch$ = Mid$(txt$, i%, 1)
    If (InStr("0123456789.,", ch$) <= 0) Then
        TaxierTyp = MAG_HILFSTAXE
        Call DefErrPop: Exit Sub
    End If
Next i%

If (l% = 12) Then Call ActProgram.PruefZiffer(txt$)

If (InStr(txt, ",") <= 0) Then
    txt2$ = Right$(Space$(13) + txt$, 13)
    If (ArtikelDbOk) Then
        SQLStr$ = "SELECT * FROM Artikel WHERE EAN = '" + txt2 + "'"
        FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr)
        If (FabsErrf% = 0) Then
            txt$ = ass.pzn
        Else
            If (Len(txt$) > 8) Then txt$ = Mid$(txt2$, 5, 8)
        End If
        ArtikelAdoRec.Close
    Else
        FabsErrf% = ass.IndexSearch(1, txt2$, FabsRecno&)
        If (FabsErrf% = 0) Then
            ass.GetRecord (FabsRecno& + 1)
            txt$ = ass.pzn
        Else
            If (Len(txt$) > 7) Then txt$ = Mid$(txt2$, 6, 7)
        End If
    End If
    
    If (ArtikelDbOk) Then
        SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + txt
    '                    SQLStr = SQLStr + " AND LagerKz<>0"
        FabsErrf = Hilfstaxe.OpenRecordset(HilfstaxeRec, SQLStr)
    Else
        FabsErrf% = hTaxe.IndexSearch(0, txt$, FabsRecno&)
        If (FabsErrf% = 0) Then
            Call hTaxe.GetRecord(FabsRecno& + 1)
        End If
    End If
    If (FabsErrf% = 0) Then
        Call UmspeichernHilfsTaxe
        TaxierRec.ActMenge = TaxierRec.GStufe
        Call ZeigeTaxierZeile(flxTaxieren.row)
    
        TxtCol% = 1
        Call ShowEditBox
        Call DefErrPop: Exit Sub
    Else
        If (ArtikelDbOk) Then
            SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + txt
    '                    SQLStr = SQLStr + " AND LagerKz<>0"
            FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr)
        Else
            FabsErrf% = ast.IndexSearch(0, txt, FabsRecno&)
            If (FabsErrf% = 0) Then
                ast.GetRecord (FabsRecno& + 1)
            End If
        End If
        If (FabsErrf% = 0) Then
    '        ast.GetRecord (FabsRecno& + 1)
        ElseIf (InStr(txt, ",") <= 0) Then
            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + txt$
            'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
            On Error Resume Next
            TaxeRec.Close
            Err.Clear
            On Error GoTo DefErr
            TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
            If (TaxeRec.EOF = False) Then
                Call Taxe2ast(txt$)
                FabsErrf% = 0
            End If
        End If
        
        If (FabsErrf% = 0) Then
            Call UmspeichernSpezialitaet
            TaxierRec.ActMenge = TaxierRec.GStufe
            Call ZeigeTaxierZeile(flxTaxieren.row)
    
            TxtCol% = 1
            Call ShowEditBox
            
            Call DefErrPop: Exit Sub
        End If
    End If
End If

lWert# = xVal(txt$)
If (lWert# > 0) And (lWert# < 10000000) Then
    TaxierTyp = MAG_PREISEINGABE
End If
   
Call DefErrPop
End Sub

Private Function AuswahlArbEmb%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuswahlArbEmb%")
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
Dim i%, ind%, erg%
Dim spBreite&, pos&
Dim dVal#
Dim txt$
    
With frmEdit.flxEdit
    .Left = 0
    .Top = 0
    
    .Rows = 2
    .FixedRows = 1
    .FormatString = "|>Preis|>Menge|Meh|<Kurz"
    .Rows = 1
    
    .GridLines = flexGridNone
    .ForeColor = MagDarstellung&(TaxierTyp, 0)
    .BackColor = MagDarstellung&(TaxierTyp, 1)
    
    .ColWidth(0) = 0
    .ColWidth(1) = TextWidth("9999999.99")
    .ColWidth(2) = TextWidth(String(8, "9"))
    .ColWidth(3) = TextWidth(String(4, "A"))
    .ColWidth(4) = TextWidth(String(50, "A"))
    
    erg% = EinlesenPassendeArbEmb%
    
    If (erg% = False) Then
        AuswahlArbEmb% = False: Call DefErrPop: Exit Function
    End If
    
    .Height = .RowHeight(0) * 15 + 90
    
    spBreite& = 0
    For i% = 0 To (.Cols - 1)
        spBreite& = spBreite& + .ColWidth(i%)
    Next i%
    .Width = spBreite& + 90
    
    frmEdit.Height = .Height
    frmEdit.Width = .Width
    frmEdit.Left = flxTaxieren.Left
    frmEdit.Left = frmEdit.Left + Left + wpara.FrmBorderHeight
    frmEdit.Top = flxTaxieren.Top
    frmEdit.Top = frmEdit.Top + Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
    
    .row = 1
    .col = 0
    .ColSel = .Cols - 1
    .Visible = True
End With

frmEdit.Show 1

If (EditErg%) Then
    ind% = InStr(EditTxt$, vbTab)
    txt$ = Left$(EditTxt$, ind% - 1)
    
    If (Len(txt$) = 8) Then
        If (Val(txt$) = 1111112) Then
            Parenteral_AOK_LosGebiet = True
            AuswahlArbEmb% = 0
            Call DefErrPop: Exit Function
        ElseIf (Val(txt$) = 1111111) Then
            Parenteral_AOK_NordOst = True
            AuswahlArbEmb% = 0
            If (para.Newline) Then
                nlcmdTrägerLösung.Visible = True
            Else
                cmdTrägerLösung.Visible = True
            End If
            Call DefErrPop: Exit Function
        Else
            EditTxt = Mid$(EditTxt, ind + 1)
            ind% = InStr(EditTxt$, vbTab)
            dVal = xVal(Left$(EditTxt$, ind% - 1))
            
            ind = 0
            For i% = 0 To UBound(ParenteralPzn)
                If (txt$ = ParenteralPzn(i)) And (dVal = ParenteralPreis(i)) Then
                    ind = i
                    Exit For
                End If
            Next i
            With TaxierRec
                .pzn = txt$
                .menge = Space$(Len(.menge))
                .Meh = Space$(Len(.Meh))
                If (Parenteral_AOK_LosGebiet) Then
                    .kp = 0
                Else
                    .kp = ParenteralPreis(ind) * SollMenge
                End If
                .GStufe = SollMenge
                
                .kurz = ParenteralTxt(ind)
'                Call CharToOem(.kurz, .kurz)
                
                .ActMenge = 0#
                .ActPreis = 0#
                .flag = MAG_ARBEIT
            End With
        
            With chkKassenRabatt(0)
                .Value = 0
                .Visible = False
            End With
            RezepturMitFaktor% = False
            ParenteralRezept = ind
        
            If (para.Newline) Then
                nlcmdVerwurf.Visible = True
            Else
                cmdVerwurf.Visible = True
            End If
            
            With flxTaxSumme
                .ColWidth(2) = TextWidth(String(20, "X"))
                
                spBreite& = 0
                For i% = 0 To (.Cols - 1)
                    spBreite& = spBreite& + .ColWidth(i%)
                Next i%
                .Width = spBreite& + 90
            End With
            
            SubstitutionsMg = 0
            For i = 0 To 2
                SubstitutionsAbgaben(i) = 0
            Next
            SubstitutionsEinzelpreis = 0

        End If
    Else
        pos = Val(txt$)
        
        If (RezRbhLauer) Then
            If (TaxierRec.flag = MAG_GEFAESS) Then
                SQLStr = "SELECT * FROM Artikel WHERE ID=" + CStr(pos)
                RezRbhLauerRec.Open SQLStr, Hilfstaxe.ActiveConn
            Else
                SQLStr = "SELECT * FROM ArbeitsPreise WHERE ID=" + CStr(pos)
                RezRbhLauerRec.Open SQLStr, taxeAdoDB.ActiveConn
            End If
            Call UmspeichernArbEmb(TaxierRec.flag)
            RezRbhLauerRec.Close
        Else
            Seek #ARBEMB%, (128& * pos) + 1
            Get #ARBEMB%, , ArbEmbRec
            
            Call UmspeichernArbEmb(TaxierRec.flag)
        End If
    End If
        
    TaxierRec.ActPreis = TaxierRec.kp
    TaxierRec.ActMenge = TaxierRec.GStufe
        
    If (TaxierRec.flag = MAG_SONSTIGES) Then
        TaxierRec.ActMenge = 1
    ElseIf (TaxierRec.flag = MAG_GEFAESS) Then
        If (ParenteralRezept = 24) Or (ParenteralRezept = 26) Or (ParenteralRezept = 28) Or (ParenteralRezept = 30) Then
            TaxierRec.ActPreis = TaxierRec.ActPreis * 1.9
        ElseIf (ParenteralRezept = 25) Or (ParenteralRezept = 27) Or (ParenteralRezept = 29) Then
            TaxierRec.ActPreis = TaxierRec.ActPreis * 2
        Else
            TaxierRec.ActPreis = TaxierRec.ActPreis * MalFaktor#
        End If
    ElseIf (ParenteralRezept >= 0) And (ParenteralRezept < 31) Then
        NeuMalFaktor# = 1
    Else
        If (Left$(TaxierRec.kurz, 5) = "UNVER") Then
            NeuMalFaktor# = 2#
        Else
            NeuMalFaktor = 1.9
        End If
    End If
End If

AuswahlArbEmb% = EditErg%
   
Call DefErrPop
End Function

Function EinlesenPassendeArbEmb%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinlesenPassendeArbEmb%")
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
Dim i%, ind%, NeuActMenge%, ret%
Dim h$, h2$, SollTyp$, AltKurz$, AktKurz$

ret% = False

With TaxierRec
    If (RezRbhLauer) Then
        If (TaxierTyp = MAG_GEFAESS) Then
            SQLStr = "SELECT Id,PZN,Name,FloatMenge AS Menge,Einheit,Kp FROM Artikel AS A"
'            SQLStr = SQLStr + " WHERE FloatMenge=(SELECT MIN(FloatMenge) FROM Artikel WHERE (ID=A.ID) AND (FloatMenge>=" + uFormat(SollMenge, "0.00") + "))"
            SQLStr = SQLStr + " WHERE FloatMenge=(SELECT MIN(FloatMenge) FROM Artikel WHERE (NAME=A.NAME) AND (FloatMenge>=" + uFormat(SollMenge, "0.00") + "))"
            SQLStr = SQLStr + " AND (Emballage<>0)"
            SQLStr = SQLStr + " ORDER BY Name"
            RezRbhLauerRec.Open SQLStr, Hilfstaxe.ActiveConn
        Else
            SQLStr = "SELECT * FROM ArbeitsPreise AS A"
'            SQLStr = SQLStr + " WHERE Menge=(SELECT MIN(Menge) FROM Arbeitspreise WHERE (ID=A.ID) AND (Menge>=" + uFormat(SollMenge, "0.00") + "))"
            SQLStr = SQLStr + " WHERE Menge=(SELECT MIN(Menge) FROM Arbeitspreise WHERE (ARBEIT=A.ARBEIT) AND (Menge>=" + uFormat(SollMenge, "0.00") + "))"
            SQLStr = SQLStr + " ORDER BY Arbeit"
            RezRbhLauerRec.Open SQLStr, taxeAdoDB.ActiveConn
       End If
        
        Do
            If (RezRbhLauerRec.EOF) Then
                Exit Do
            End If
            
            Call UmspeichernArbEmb(TaxierTyp)
'            Do
'                If (.Gstufe >= SollMenge) Then
'                    Exit Do
'                End If
'
'                .Kp = .Kp + CheckNullDouble(RezRbhLauerRec!StaffelPreis)
'                .Gstufe = .Gstufe + CheckNullDouble(RezRbhLauerRec!Staffelmenge)
'            Loop
            If (TaxierTyp = MAG_GEFAESS) Then
                AktKurz = iTrim(CheckNullStr(RezRbhLauerRec!Name))
            Else
                AktKurz = .kurz
            End If
        
            If ((TaxierTyp = MAG_SONSTIGES) And (.GStufe = 0)) Or ((TaxierTyp <> MAG_SONSTIGES) And (.GStufe >= SollMenge#) And (AktKurz <> AltKurz$)) Then
                AltKurz$ = AktKurz
                
                h$ = Format(RezRbhLauerRec!Id, "0")
                h$ = h$ + vbTab + Format(.kp, "0.00") + vbTab + Format(.GStufe, "0.00") + vbTab
            
                h2$ = iTrim$(AktKurz)
'                Call OemToChar(h2$, h2$)
                h$ = h$ + vbTab + h2$
                
                frmEdit.flxEdit.AddItem h$
                
                ret% = True
            End If
            
            RezRbhLauerRec.MoveNext
            
            ind% = ind% + 1
        Loop
        RezRbhLauerRec.Close
    
    Else
        If (TaxierTyp = MAG_GEFAESS) Then
            ind% = AnfEmb%
            SollTyp$ = "E"
        Else
            ind% = 1
            SollTyp$ = "A"
        End If
        
        Seek #ARBEMB%, (128& * ind%) + 1
                
        'h$ = iTrim$(TmInhalt.kurz)
        'h$ = Left$(h$ + Space$(28), 28)
        
        AltKurz$ = ""
    
        Do
            Get #ARBEMB%, , ArbEmbRec
            If (EOF(ARBEMB%)) Then Exit Do
            If (ArbEmbRec.Typ <> SollTyp$) Then Exit Do
        
            Call UmspeichernArbEmb(TaxierTyp)
        
            If ((TaxierTyp = MAG_SONSTIGES) And (.GStufe = 0)) Or ((TaxierTyp <> MAG_SONSTIGES) And (.GStufe >= SollMenge#) And (.kurz <> AltKurz$)) Then
                AltKurz$ = .kurz
                
                h$ = Format(ind%, "0")
                h$ = h$ + vbTab + Format(.kp, "0.00") + vbTab + Format(.GStufe, "0.00") + vbTab
            
                h2$ = iTrim$(.kurz)
                Call OemToChar(h2$, h2$)
                h$ = h$ + vbTab + h2$
                
                frmEdit.flxEdit.AddItem h$
                
                ret% = True
            End If
            
            ind% = ind% + 1
        Loop
    End If

    If (TaxierTyp = MAG_ARBEIT) Then
        With frmEdit.flxEdit
            For i% = 0 To UBound(ParenteralPzn)
                h$ = ParenteralPzn(i)
                h$ = h$ + vbTab + Format(ParenteralPreis(i), "0.00")
                h$ = h$ + vbTab + Format(SollMenge, "0.00")
                h$ = h$ + vbTab
                h$ = h$ + vbTab + ParenteralTxt(i)
                .AddItem h$
            Next i
    
            If Not (Parenteral_AOK_LosGebiet) And Not (Parenteral_AOK_NordOst) Then
                If (Dir("MgPreis1.mdb") <> "") Then
                    h$ = "01111112"
                    h$ = h$ + vbTab '+ Format(0, "0.00")
                    h$ = h$ + vbTab '+ Format(1, "0.00")
                    h$ = h$ + vbTab
                    h$ = h$ + vbTab + "1 AOK Losgebiet"
                    .AddItem h$
                End If
                If (Dir("MgPreis2.mdb") <> "") Then
                    h$ = "01111111"
                    h$ = h$ + vbTab '+ Format(0, "0.00")
                    h$ = h$ + vbTab '+ Format(1, "0.00")
                    h$ = h$ + vbTab
                    h$ = h$ + vbTab + "2 AOK Nordost"
                    .AddItem h$
                End If
            End If
            
            .row = 1
            .col = 4
            .RowSel = .Rows - 1
            .ColSel = .col
            .Sort = 5
            .col = 0
            .ColSel = .Cols - 1
        End With
    End If
End With

EinlesenPassendeArbEmb% = ret%
        
Call DefErrPop
End Function

Sub SpeicherTaxMuster()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherTaxMuster")
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
Dim i%, j%, erg%, anz%, ok%
Dim Wohin&, TmId&
Dim TmName$, h$, sKurz$

If (ParenteralRezept >= 0) Then
    If (FormErg% > 0) Then
        SQLStr = "DELETE * FROM ParenteralTm WHERE Id=" + CStr(FormErg)
        RezSpeicherDB.Execute SQLStr
        SQLStr = "DELETE * FROM ParenteralTmZeilen WHERE TmId=" + CStr(FormErg)
        RezSpeicherDB.Execute SQLStr
    End If
    
    Set ParenteralTmRec = RezSpeicherDB.OpenRecordset("ParenteralTm", dbOpenTable)
    ParenteralTmRec.AddNew
    ParenteralTmRec!Bezeichnung = FormErgTxt$
    ParenteralTmRec.Update
    ParenteralTmRec.Bookmark = ParenteralTmRec.LastModified
    TmId = ParenteralTmRec!Id

    Set ParenteralTmRec = RezSpeicherDB.OpenRecordset("ParenteralTmZeilen", dbOpenTable)
    With flxTaxieren
        For i% = 1 To (.Rows - 1)
            If (i% > 100) Then Exit For
                
            sKurz = Trim(.TextMatrix(i%, 3))
            ok = (sKurz <> "")
            If (ok) Then
                ok = (Left(UCase(sKurz), 9) <> "AUFSCHLAG")
            End If
            
            If (ok) Then
                TaxierRec.ActPreis = iCDbl(.TextMatrix(i%, 0))
                TaxierRec.ActMenge = iCDbl(.TextMatrix(i%, 1))
                TaxierRec.Meh = .TextMatrix(i%, 2)
                TaxierRec.kurz = .TextMatrix(i%, 3)
                TaxierRec.pzn = .TextMatrix(i%, 4)
                TaxierRec.flag = Val(.TextMatrix(i%, 6))
                TaxierRec.kp = iCDbl(.TextMatrix(i%, 7))
                TaxierRec.GStufe = iCDbl(.TextMatrix(i%, 8))
                TaxierRec.Verwurf = iCDbl(flxTaxieren.TextMatrix(i%, 12))
            
                ParenteralTmRec.AddNew
                
                ParenteralTmRec!TmId = TmId
                ParenteralTmRec!AnfMagInd = i%
                
                ParenteralTmRec!ActPreis = TaxierRec.ActPreis
                ParenteralTmRec!ActMenge = TaxierRec.ActMenge
                ParenteralTmRec!einheit = TaxierRec.Meh
                ParenteralTmRec!text = TaxierRec.kurz
                ParenteralTmRec!pzn = Val(TaxierRec.pzn)
                ParenteralTmRec!flag = TaxierRec.flag
                ParenteralTmRec!kp = TaxierRec.kp
                ParenteralTmRec!GStufe = TaxierRec.GStufe
                
                ParenteralTmRec!packmittel = Val(.TextMatrix(i%, 9))
                ParenteralTmRec!AI = Val(.TextMatrix(i%, 10))
                ParenteralTmRec!WirkstoffMenge = iCDbl(.TextMatrix(i%, 11))
                
                If (Parenteral_AOK_LosGebiet) Then
                    ParenteralTmRec!menge = "1"
                ElseIf (Parenteral_AOK_NordOst) Then
                    ParenteralTmRec!menge = "2"
                End If
                
                ParenteralTmRec.Update
                j% = j% + 1
            End If
        Next i%
'        For i% = (j% + 1) To 100
'            TaxierRec.flag = 255
'            Put #MAG_SPEICHER%, , TaxierRec
'        Next i%
    End With
ElseIf (TaxmusterDBok) Then
    If (FormErg% > 0) Then
        SQLStr = "DELETE * FROM Taxmuster WHERE Id=" + CStr(FormErg)
        TaxmusterDB.Execute SQLStr
        SQLStr = "DELETE * FROM TaxmusterZeilen WHERE TaxmusterId=" + CStr(FormErg)
        TaxmusterDB.Execute SQLStr
    End If
    
    Set TaxmusterRec = TaxmusterDB.OpenRecordset("Taxmuster", dbOpenTable)
    TaxmusterRec.AddNew
    TaxmusterRec!Bezeichnung = FormErgTxt$
    TaxmusterRec!ActMenge = CInt(TeilMenge#)
'    TaxmusterRec!AnzZeilen = flxTaxieren.Rows - 1
    
    TaxmusterRec.Update
    TaxmusterRec.Bookmark = TaxmusterRec.LastModified
    TmId = TaxmusterRec!Id

    With flxTaxieren
        For i% = 1 To (.Rows - 1)
            TmInhalt.pzn = .TextMatrix(i%, 4)
            
            h$ = .TextMatrix(i%, 3)
'            Call CharToOem(h$, h$)
            TmInhalt.kurz = h$
            
            TmInhalt.ActMenge = iCDbl(.TextMatrix(i%, 1))
            TmInhalt.ActPreis = iCDbl(.TextMatrix(i%, 0))
            TmInhalt.flag = Val(.TextMatrix(i%, 6))
            TmInhalt.NextSatz = 0
            
            If (Trim(TmInhalt.kurz) <> "") Then
                j = j + 1
                SQLStr = "INSERT INTO TaxmusterZeilen (TaxmusterId,LaufNr,Pzn,Name,flag,ActMenge,ActPreis)"
                SQLStr = SQLStr + " VALUES (" + CStr(TmId) + "," + CStr(j) + "," + CStr(Val(TmInhalt.pzn)) + ",'" + TmInhalt.kurz + "'"
                SQLStr = SQLStr + "," + CStr(TmInhalt.flag) + "," + uFormat(TmInhalt.ActMenge, "0.0000") + "," + uFormat(TmInhalt.ActPreis, "0.00")
                SQLStr = SQLStr + ")"
                TaxmusterDB.Execute (SQLStr)
            End If
        Next i%
    End With
                
    SQLStr = "UPDATE Taxmuster SET AnzZeilen=" + CStr(j)
    SQLStr = SQLStr + " WHERE Id=" + CStr(TmId)
    TaxmusterDB.Execute (SQLStr)
Else
    If (FormErg% > 0) Then
        Seek TM_NAMEN%, 1& * FormErg% * Len(TmHeader) + 1
        Get #TM_NAMEN%, , TmHeader
        
        Call HoleTaxMusterZeilen(True)
    End If
    
    TmHeader.Name = FormErgTxt$
    For i% = 0 To 2
        TmHeader.Inhalt(i%).Name = Space$(Len(TmHeader.Inhalt(i%)))
    Next i%
    
    Call SpeicherTaxMusterZeilen
    
    If (FormErg% < 0) Then
        Seek #TM_NAMEN%, 1
        h$ = String(2, 0)
        Get #TM_NAMEN%, , h$
        anz% = CVI(h$)
        anz% = anz% + 1
        h$ = MKI(anz%)
        Seek #TM_NAMEN%, 1
        Put #TM_NAMEN%, , h$
        
        Wohin& = anz%   ' LOF(TM_NAMEN%) / Len(TmHeader)
    Else
        Wohin& = FormErg%
    End If
    Seek #TM_NAMEN%, (Wohin& * Len(TmHeader)) + 1
    Put #TM_NAMEN%, , TmHeader
End If
    
Call DefErrPop
End Sub

Sub SpeicherTaxMusterZeilen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherTaxMusterZeilen")
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
Dim i%, j%, erg%
Dim h$
Dim LetztTmInhalt As TaxMusterInhaltStruct

LetztTmInhalt.NextSatz = 0
j% = 0
With flxTaxieren
    For i% = 1 To (.Rows - 1)
        TmInhalt.pzn = .TextMatrix(i%, 4)
        
        h$ = .TextMatrix(i%, 3)
        Call CharToOem(h$, h$)
        TmInhalt.kurz = h$
        
        TmInhalt.ActMenge = iCDbl(.TextMatrix(i%, 1))
        TmInhalt.ActPreis = iCDbl(.TextMatrix(i%, 0))
        TmInhalt.flag = Val(.TextMatrix(i%, 6))
        TmInhalt.NextSatz = 0
        
        If (Trim(TmInhalt.kurz) <> "") Then
            erg% = SpeicherTaxMusterZeile%(TmInhalt, LetztTmInhalt)
            If (erg% = False) Then Exit For
    
            LetztTmInhalt = TmInhalt
    
            If (i% = 1) Then
                TmHeader.ErstSatz = TmInhalt.NextSatz
            End If
            
            If (i% <= 3) Then
                Call MacheTaxMusterInhalt(i% - 1)
            End If
            
            j% = j% + 1
        End If
        
    Next i%
    
    TmHeader.ActMenge = CInt(TeilMenge#)
    TmHeader.AnzZeilen = j%
End With
       
Call DefErrPop
End Sub

Sub MacheTaxMusterInhalt(ind%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MacheTaxMusterInhalt")
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
  
If (TmInhalt.flag = MAG_PREISEINGABE) Then
    TmHeader.Inhalt(ind%).Name = Format(TmInhalt.ActPreis, "    0.00")
Else
    TmHeader.Inhalt(ind%).Name = TmInhalt.kurz
End If
       
Call DefErrPop
End Sub

Private Function SpeicherTaxMusterZeile%(ActTmInhalt As TaxMusterInhaltStruct, LetztTmInhalt As TaxMusterInhaltStruct)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherTaxMusterZeile%")
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
Dim WoBinIch&

WoBinIch& = HoleTaxMusterFreiSatz&
If (WoBinIch& = False) Then
    SpeicherTaxMusterZeile% = False: Call DefErrPop: Exit Function
End If

ActTmInhalt.NextSatz = 0&
Put #TM_DATEN%, , ActTmInhalt
'if (erg != TmDatenSatzLen)
'  return(FALSE);

ActTmInhalt.NextSatz = WoBinIch&
If (LetztTmInhalt.NextSatz > 0&) Then
    Seek #TM_DATEN%, (LetztTmInhalt.NextSatz * Len(TmInhalt)) + 1
    LetztTmInhalt.NextSatz = WoBinIch&
    Put #TM_DATEN%, , LetztTmInhalt
'    if (erg != TmDatenSatzLen)
'      return(FALSE);
End If

SpeicherTaxMusterZeile% = True

Call DefErrPop
End Function

Function HoleTaxMusterFreiSatz&()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleTaxMusterFreiSatz&")
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
Dim ErsterFreier&, ret&
Dim h$
Dim iTmInhalt As TaxMusterInhaltStruct

ret& = False

Seek #TM_DATEN%, 1
h$ = String(4, 0)
Get #TM_DATEN%, , h$
ErsterFreier& = CVL(h$)

If (ErsterFreier& < 0&) Then
    HoleTaxMusterFreiSatz& = False: Call DefErrPop: Exit Function
End If

If (ErsterFreier& = 0&) Then
    ret& = LOF(TM_DATEN%) / Len(TmInhalt)
    Seek #TM_DATEN%, (ret& * Len(TmInhalt)) + 1
Else
    Seek #TM_DATEN%, (ErsterFreier& * Len(TmInhalt)) + 1
    Get #TM_DATEN%, , iTmInhalt
    
'    h$ = iTmInhalt.pzn + Left$(iTmInhalt.kurz, 8)
'    If (h$ = String(15, 0)) Then
        Seek #TM_DATEN%, 1
        h$ = MKL(iTmInhalt.NextSatz)
        Put #TM_DATEN%, , h$
        
        Seek #TM_DATEN%, (ErsterFreier& * Len(TmInhalt)) + 1
        ret& = ErsterFreier&
'    End If
End If

HoleTaxMusterFreiSatz& = ret&

Call DefErrPop
End Function
   
Private Sub HoleMagSpeicher()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleMagSpeicher")
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
Dim i%, j%
Dim lPzn&
Dim h$, PreisKz$, FaktorKz$
Dim GStufe As Double

If (MAG_SPEICHER% > 0) Then

    'MsgBox (Format(HashErstellDat, "dd.MM.yyyy"))

    With flxTaxieren
        .Visible = False
        .Rows = 1
        
        Dim CannabisFixaufschlag#
        CannabisFixaufschlag = 0
        If (ParenteralRezept > 23) Then
            Seek #MAG_SPEICHER%, ((MagSpeicherIndex% - 1) * CLng(Len(TaxierRec))) + 1
            For i% = 1 To 100
                Get #MAG_SPEICHER%, , TaxierRec
                If (EOF(MAG_SPEICHER%)) Then Exit For
                If (TaxierRec.flag = 255) Then Exit For
            
                If (InStr(UCase(TaxierRec.kurz), "FIXZUSCHLAG") > 0) Then
                    CannabisFixaufschlag = TaxierRec.ActPreis
                    Exit For
                End If
            Next
        End If
        
        Seek #MAG_SPEICHER%, ((MagSpeicherIndex% - 1) * CLng(Len(TaxierRec))) + 1
        If (ParenteralRezept > 15) And (ParenteralRezept < 24) Then
            Dim sBTM$, sPzns$, sSonderPzn$
            sBTM = ""
            sPzns = ""
            'sSonderPzn = PznString(Val(IIf((ParenteralRezept = 22) Or (ParenteralRezept = 23), ParenteralPzn(20), ParenteralPzn(ParenteralRezept)))) + "55" + Format(SubstitutionsMg * 10, "00000") + "14" + Format(SubstitutionsEinzelpreis * 100, "000000000")
            sSonderPzn = PznString(Val(IIf((ParenteralRezept = 22) Or (ParenteralRezept = 23), ParenteralPzn(20), ParenteralPzn(ParenteralRezept)))) + "55" + HashFaktor(SubstitutionsMg * 10) + "14" + HashPreis(SubstitutionsEinzelpreis * 100)
            For i% = 1 To 100
                Get #MAG_SPEICHER%, , TaxierRec
                If (EOF(MAG_SPEICHER%)) Then Exit For
                If (TaxierRec.flag = 255) Then Exit For
            
                .AddItem " "
                Call ZeigeTaxierZeile(.Rows - 1)
                
                PreisKz = "14"
                FaktorKz = "11"
                
                With TaxierRec
                    lPzn = Val(.pzn)
                    
                    If (.GStufe <> 0) Then
                        GStufe = .GStufe
                    Else
                        GStufe = 1
                    End If
                    If (.flag = MAG_HILFSTAXE) Or (.flag = MAG_SPEZIALITAET) Then
                        h$ = .pzn + FaktorKz + HashFaktor(.ActMenge / GStufe * 1000) + PreisKz + HashPreis(.ActPreis * 100)
                        sPzns = sPzns + h$
'                    ElseIf (.pzn = "02567001") Then
'                        sBTM = .pzn + FaktorKz + Format(1000 * .ActMenge, "00000") + "81" + Format(.ActPreis * 100, "000000000")
                    ElseIf (lPzn = 2567001) Or (lPzn = 9999637) Or (lPzn = 6461110) Or (lPzn = 2567018) Then
                        If (lPzn = 2567001) Then  'BTM
                            PreisKz = "81"
                        ElseIf (lPzn = 9999637) Then  'Beschaffungskosten
                            PreisKz = "82"
                        ElseIf (lPzn = 6461110) Then  'Botendienst
                            PreisKz = "83"
                        ElseIf (lPzn = 2567018) Then  'Noctu
                            PreisKz = "80"
                        End If
                        h = .pzn + FaktorKz + HashFaktor(1000 * .ActMenge) + PreisKz + HashPreis(.ActPreis * 100)
                        sBTM = sBTM + h$
                    ElseIf (.flag = MAG_GEFAESS) Then
                        h$ = PznString(Val(.pzn)) + FaktorKz + HashFaktor(1000) + PreisKz + HashPreis(.ActPreis * 100)
                        sPzns = sPzns + h$
                    End If
                End With
            Next i%
            If (pCharge <> "") Then
                pCharge = Left(pCharge, Len(pCharge) - 4)
                For i = 0 To 2
                    If (SubstitutionsAbgaben(i) <= 0) Then
                        Exit For
                    End If
                    For j = 1 To SubstitutionsAbgaben(i)
                        ParenteralPara = ParenteralPara + pCharge + Format(i + 1, "00") + Format(j, "00") + sPzns
                    Next j
                    ParenteralPara = ParenteralPara + sSonderPzn
                Next i
                If (sBTM <> "") Then
                    ParenteralPara = ParenteralPara + sBTM
                End If
            End If
        Else
            For i% = 1 To 100
                Get #MAG_SPEICHER%, , TaxierRec
                If (EOF(MAG_SPEICHER%)) Then Exit For
                If (TaxierRec.flag = 255) Then Exit For
            
                .AddItem " "
                Call ZeigeTaxierZeile(.Rows - 1)
                
                PreisKz = "14"
                If (Parenteral_AOK_LosGebiet) Then
                    PreisKz = "15"
                End If
                
                FaktorKz = "11"
                If (TaxierRec.Verwurf) Then
                    FaktorKz = "99"
                End If
                
                With TaxierRec
                    lPzn = Val(.pzn)
                    
    '                If (.Gstufe <> 0) Then
    '                    If (.flag = MAG_HILFSTAXE) Then
    '    '                    h$ = Mid(.pzn, 2) + "11" + Format(.ActMenge / .Gstufe * 1000, "00000") + PreisKz + Format(.ActPreis * 100, "000000000")
    '                        h$ = .pzn + FaktorKz + Format(.ActMenge / .Gstufe * 1000, "00000") + PreisKz + Format(.ActPreis * 100, "000000000")
    '                        ParenteralPara = ParenteralPara + h$
    '                    ElseIf (.flag = MAG_SPEZIALITAET) Then
    '    '                    h$ = Mid(.pzn, 2) + "11" + Format(.ActMenge / .Gstufe * 1000, "00000") + PreisKz + Format(.Kp * 100, "000000000")
    ''                        h$ = .pzn + FaktorKz + Format(.ActMenge / .Gstufe * 1000, "00000") + PreisKz + Format(.kp * 100, "000000000")
    '                        h$ = .pzn + FaktorKz + Format(.ActMenge / .Gstufe * 1000, "00000") + PreisKz + Format(.ActPreis * 100, "000000000")
    '                        ParenteralPara = ParenteralPara + h$
    '                    ElseIf (.flag = MAG_ARBEIT) And ((.pzn = "09999092") Or (.pzn = "02567478") Or (.pzn = "02567461") Or (.pzn = "09999146") Or (.pzn = "09999169")) Then
    '    '                    h$ = Mid(.pzn, 2) + "11" + Format(.ActMenge / .Gstufe * 1000, "00000") + PreisKz + Format(.Kp * 100, "000000000")
    '                        h$ = "06460518" + FaktorKz + Format(1000, "00000") + "74" + Format(.kp * 100, "000000000")
    '                        ParenteralPara = ParenteralPara + h$
    '                    ElseIf (.flag = MAG_GEFAESS) And (ParenteralRezept > 15) Then
    '    '                    h$ = Mid(.pzn, 2) + "11" + Format(.ActMenge / .Gstufe * 1000, "00000") + PreisKz + Format(.Kp * 100, "000000000")
    '                        h$ = .pzn + FaktorKz + Format(1000, "00000") + PreisKz + Format(.ActPreis * 100, "000000000")
    '                        ParenteralPara = ParenteralPara + h$
    '                    End If
    '                Else
    '                    If (.flag = MAG_HILFSTAXE) Then
    '                        h$ = .pzn + FaktorKz + Format(.ActMenge * 1000, "00000") + PreisKz + Format(.ActPreis * 100, "000000000")
    '                        ParenteralPara = ParenteralPara + h$
    '                    ElseIf (.flag = MAG_SPEZIALITAET) Then
    ''                        h$ = .pzn + FaktorKz + Format(.ActMenge * 1000, "00000") + PreisKz + Format(.kp * 100, "000000000")
    '                        h$ = .pzn + FaktorKz + Format(.ActMenge * 1000, "00000") + PreisKz + Format(.ActPreis * 100, "000000000")
    '                        ParenteralPara = ParenteralPara + h$
    '                    ElseIf (.flag = MAG_ARBEIT) And ((.pzn = "09999092") Or (.pzn = "02567478") Or (.pzn = "02567461") Or (.pzn = "09999146") Or (.pzn = "09999169")) Then
    '    '                    h$ = Mid(.pzn, 2) + "11" + Format(.ActMenge / .Gstufe * 1000, "00000") + PreisKz + Format(.Kp * 100, "000000000")
    '                        h$ = "06460518" + FaktorKz + Format(1000, "00000") + "74" + Format(.kp * 100, "000000000")
    '                        ParenteralPara = ParenteralPara + h$
    '                    ElseIf (.flag = MAG_GEFAESS) And (ParenteralRezept > 15) Then
    '    '                    h$ = Mid(.pzn, 2) + "11" + Format(.ActMenge / .Gstufe * 1000, "00000") + PreisKz + Format(.Kp * 100, "000000000")
    '                        h$ = .pzn + FaktorKz + Format(1000, "00000") + PreisKz + Format(.ActPreis * 100, "000000000")
    '                        ParenteralPara = ParenteralPara + h$
    '                    End If
    '                End If
                    
                    Dim ActPreis As Double
                    If (.GStufe <> 0) Then
                        GStufe = .GStufe
                    Else
                        GStufe = 1
                    End If
'                    If (.flag = MAG_HILFSTAXE) Or (.flag = MAG_SPEZIALITAET) Or ((ParenteralRezept < 0) And (.flag = MAG_ANTEILIG)) Then
                    If (.flag = MAG_HILFSTAXE) Or (.flag = MAG_SPEZIALITAET) Or (.flag = MAG_ANTEILIG) Then
                        ActPreis = .ActPreis
'                        If (ParenteralRezept > 15) And (ParenteralRezept < 21) Then
                        If (ParenteralRezept > 23) Then
                            ActPreis = ActPreis + CannabisFixaufschlag
                            CannabisFixaufschlag = 0
                        End If
                        h$ = .pzn + FaktorKz + HashFaktor(.ActMenge / GStufe * 1000) + PreisKz + HashPreis(ActPreis * 100)
                        ParenteralPara = ParenteralPara + h$
                    ElseIf (.flag = MAG_ARBEIT) And ((.pzn = "09999092") Or (.pzn = "02567478") Or (.pzn = "02567461") Or (.pzn = "09999146") Or (.pzn = "09999169")) Then
                        h$ = "06460518" + FaktorKz + HashFaktor(1000) + "74" + HashPreis(.kp * 100)
                        ParenteralPara = ParenteralPara + h$
'                    ElseIf (.pzn = "02567001") Then     'BTM-Gebühr
'                            h$ = .pzn + FaktorKz + Format(1000 * .ActMenge, "00000") + "81" + Format(.ActPreis * 100, "000000000")
'                            ParenteralPara = ParenteralPara + h$
'                    ElseIf (.pzn = "09999637") Then     'Beschaffungskosten
'                            h$ = .pzn + FaktorKz + Format(1000 * .ActMenge, "00000") + "82" + Format(.ActPreis * 100, "000000000")
'                    ElseIf (.pzn = "06461110") Then     'Botendienst
'                            h$ = .pzn + FaktorKz + Format(1000 * .ActMenge, "00000") + "83" + Format(.ActPreis * 100, "000000000")
'                    ElseIf (.pzn = "02567018") Then     'Noctu-Gebühr
'                            h$ = .pzn + FaktorKz + Format(1000 * .ActMenge, "00000") + "80" + Format(.ActPreis * 100, "000000000")
                    ElseIf (lPzn = 2567001) Or (lPzn = 9999637) Or (lPzn = 6461110) Or (lPzn = 2567018) Then
                        If (lPzn = 2567001) Then  'BTM
                            PreisKz = "81"
                        ElseIf (lPzn = 9999637) Then  'Beschaffungskosten
                            PreisKz = "82"
                        ElseIf (lPzn = 6461110) Then  'Botendienst
                            PreisKz = "83"
                        ElseIf (lPzn = 2567018) Then  'Noctu
                            PreisKz = "80"
                        End If
                        h = .pzn + FaktorKz + HashFaktor(1000 * .ActMenge) + PreisKz + HashPreis(.ActPreis * 100)
                        ParenteralPara = ParenteralPara + h$
                    
                    ElseIf (lPzn = 9999117) Or (lPzn = 9999206) Then 'Import RX, Import Substanz RX
                        h = .pzn + FaktorKz + HashFaktor(1) + PreisKz + HashPreis(.ActPreis * 100)
                        ParenteralPara = ParenteralPara + h$
                    
                    ElseIf (lPzn = 6461334) Then 'Artikel ohne PZN
                        h = .pzn + FaktorKz + HashFaktor(1) + "11" + HashPreis(.ActPreis * 100)
                        ParenteralPara = ParenteralPara + h$
                    
                    ElseIf (ParenteralRezept > 15) Or (ParenteralRezept < 0) Then
                        If (InStr(UCase(.kurz), "FIX-AUFSCHLAG") > 0) Then
                            h$ = "06460518" + FaktorKz + HashFaktor(1000) + IIf(PreisKz_62_70, "70", "74") + HashPreis(.ActPreis * 100)
                            ParenteralPara = ParenteralPara + h$
                        ElseIf (.flag = MAG_ARBEIT) And (.ActPreis > 0) Then
                            h$ = "06460518" + FaktorKz + HashFaktor(1000) + IIf(PreisKz_62_70, "62", "74") + HashPreis(.ActPreis * 100)
                            ParenteralPara = ParenteralPara + h$
                        ElseIf (.flag = MAG_GEFAESS) Then
                            If (Val(.pzn) = 0) Then
                                h$ = "06461328" + FaktorKz + HashFaktor(1) + PreisKz + HashPreis(.ActPreis * 100)
                            Else
                                h$ = PznString(Val(.pzn)) + FaktorKz + HashFaktor(1000) + PreisKz + HashPreis(.ActPreis * 100)
                            End If
                            ParenteralPara = ParenteralPara + h$
                        End If
                    End If
                End With
            Next i%
        End If
        
        .Visible = True
    End With
    
    Call TaxSumme
End If
    
Call DefErrPop
End Sub

'Private Sub HoleAnfMag()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("HoleAnfMag")
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
'Dim dPreis#
'Dim SQLStr$
'
'With flxTaxieren
'    .Visible = False
'    .Rows = 1
'
'    SQLStr$ = "SELECT * FROM ANFMAG WHERE ABHOLNR = " + Str$(AnfMagIndex& / 100)
'    SQLStr$ = SQLStr$ + " and ABHOLIND = " + Str$(AnfMagIndex& Mod 100)
'    SQLStr$ = SQLStr$ + " ORDER BY LAUFNR"
'    Set AnfMagRec = AbholerDB.OpenRecordset(SQLStr$)
'
'    Do
'        If (AnfMagRec.EOF) Then Exit Do
'
'        With TaxierRec
'            .pzn = sCheckNull$(AnfMagRec!pzn)
'            .kurz = sCheckNull$(AnfMagRec!kurz)
'            .menge = sCheckNull$(AnfMagRec!menge)
'            .meh = sCheckNull$(AnfMagRec!meh)
'            .flag = AnfMagRec!flag
'            .Kp = AnfMagRec!Kp
'            .Gstufe = AnfMagRec!Gstufe
'            .ActMenge = AnfMagRec!ActMenge
'            .ActPreis = AnfMagRec!ActPreis
'        End With
'
'        .AddItem " "
'        Call ZeigeTaxierZeile(.Rows - 1)
'
'        AnfMagRec.MoveNext
'    Loop
'
'    .Visible = True
'End With
'
'Call TaxSumme
'
'Call DefErrPop
'End Sub
'
'Private Sub SpeicherAnfMag()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("SpeicherAnfMag")
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
'Dim i%, j%
'Dim iKurz$, SQLStr$
'
'With flxTaxSumme
'    AnfMagPreis# = xVal(.TextMatrix(.Rows - 1, 0))
'End With
'
'SQLStr$ = "DELETE * FROM ANFMAG WHERE ABHOLNR = " + Str$(AnfMagIndex& / 100)
'SQLStr$ = SQLStr$ + " and ABHOLIND = " + Str$(AnfMagIndex& Mod 100)
'AbholerDB.Execute (SQLStr$)
'
'
'With flxTaxieren
'    j% = 1
'    For i% = 1 To (.Rows - 1)
'        If (i% > 100) Then Exit For
'
'        iKurz$ = Trim(.TextMatrix(i%, 3))
'        If (iKurz$ <> "") Then
'            AnfMagRec.AddNew
'
'            AnfMagRec!AbholNr = AnfMagIndex& / 100
'            AnfMagRec!AbholInd = AnfMagIndex& Mod 100
'            AnfMagRec!LaufNr = j%
'            j% = j% + 1
'
'            AnfMagRec!ActPreis = iCDbl(.TextMatrix(i%, 0))
'            AnfMagRec!ActMenge = iCDbl(.TextMatrix(i%, 1))
'            AnfMagRec!meh = .TextMatrix(i%, 2)
'            AnfMagRec!kurz = iKurz$
'            AnfMagRec!pzn = .TextMatrix(i%, 4)
'            AnfMagRec!flag = Val(.TextMatrix(i%, 6))
'            AnfMagRec!Kp = iCDbl(.TextMatrix(i%, 7))
'            AnfMagRec!Gstufe = iCDbl(.TextMatrix(i%, 8))
'
'            AnfMagRec.Update
'        End If
'    Next i%
'End With
'
'Call DefErrPop
'End Sub

Sub SpeicherMagSpeicher()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherMagSpeicher")
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
Dim i%, j%

If (MAG_SPEICHER% > 0) Then
    
    With flxTaxieren
        Seek #MAG_SPEICHER%, ((MagSpeicherIndex% - 1) * CLng(Len(TaxierRec))) + 1
        For i% = 1 To (.Rows - 1)
            If (i% > 100) Then Exit For
                
            TaxierRec.ActPreis = iCDbl(.TextMatrix(i%, 0))
            TaxierRec.ActMenge = iCDbl(.TextMatrix(i%, 1))
            TaxierRec.Meh = .TextMatrix(i%, 2)
            TaxierRec.kurz = .TextMatrix(i%, 3)
            TaxierRec.pzn = .TextMatrix(i%, 4)
            TaxierRec.flag = Val(.TextMatrix(i%, 6))
            TaxierRec.kp = iCDbl(.TextMatrix(i%, 7))
            TaxierRec.GStufe = iCDbl(.TextMatrix(i%, 8))
            TaxierRec.Verwurf = Abs(.TextMatrix(i%, 5) = "V")
            
            If (Trim(TaxierRec.kurz) <> "") Then
                Put #MAG_SPEICHER%, , TaxierRec
                j% = j% + 1
            End If
        Next i%
        For i% = (j% + 1) To 100
            TaxierRec.flag = 255
            Put #MAG_SPEICHER%, , TaxierRec
        Next i%
    End With
End If
    
Call DefErrPop
End Sub

Sub HoleTaxMusterZeilen(Optional LoeschFlag% = False)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleTaxMusterZeilen")
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
Dim dPreis#
Dim SatzPtr&, ErsterFreier&
Dim h$, h2$
Dim iTmInhalt As TaxMusterInhaltStruct

SatzPtr& = TmHeader.ErstSatz
For i% = 1 To TmHeader.AnzZeilen
    If (SatzPtr& <= 0&) Then Exit For
    
    Seek #TM_DATEN%, (SatzPtr& * Len(TmInhalt)) + 1
    Get #TM_DATEN%, , TmInhalt
    If (EOF(TM_DATEN%)) Then Exit For
    
    If (LoeschFlag%) Then
        Seek #TM_DATEN%, 1
        h$ = String(4, 0)
        Get #TM_DATEN%, , h$
        ErsterFreier& = CVL(h$)
        
        With iTmInhalt
            .pzn = String(Len(.pzn), 0)
            .dummy = String(Len(.dummy), 0)
            .kurz = String(Len(.kurz), 0)
            .dummy2 = String(Len(.dummy2), 0)
            .ActMenge = 0
            .ActPreis = 0
            .flag = 55  '?
            .NextSatz = ErsterFreier&
        End With

        Seek #TM_DATEN%, (SatzPtr& * Len(TmInhalt)) + 1
        Put #TM_DATEN%, , iTmInhalt
        
        Seek #TM_DATEN%, 1
        h$ = MKL(SatzPtr&)
        Put #TM_DATEN%, , h$
    Else
        Call HoleTaxMusterZeile
        flxTaxieren.AddItem " "
        Call ZeigeTaxierZeile(flxTaxieren.Rows - 1)
    End If
            
    SatzPtr& = TmInhalt.NextSatz
Next i%

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

Private Sub nlcmdf2_Click()
Call cmdF2_Click
End Sub

Private Sub nlcmdf5_Click()
Call cmdF5_Click
End Sub

Private Sub nlcmdF7_Click()
Call cmdF7_Click
End Sub

Private Sub nlcmdAuswahl_Click(index As Integer)
Call cmdAuswahl_Click(index)
End Sub

Private Sub nlcmdDarstellung_click()
Call cmdDarstellung_Click
End Sub

Private Sub nlcmdTaxmuster_click()
Call cmdTaxmuster_Click
End Sub

Private Sub nlcmdTrägerLösung_click()
Call cmdTrägerLösung_Click
End Sub

Private Sub nlcmdVerwurf_click()
Call cmdVerwurf_Click
End Sub

Private Sub nlcmdSonderfälle_Click()
Call cmdSonderfälle_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If (para.Newline) Then
    If (KeyAscii = 13) Then
'        Call nlcmdOk_Click
        Exit Sub
    ElseIf (KeyAscii = 27) And (nlcmdEsc.Visible) Then
        Call nlcmdEsc_Click
        Exit Sub
    End If
End If
    
' passiert in textbox_keypress
'If (TypeOf ActiveControl Is TextBox) Then
'    If (iEditModus% <> 1) Then
'        If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
'        If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And (((iEditModus% <> 2) And (iEditModus% <> 4)) Or (Chr$(KeyAscii) <> ".")) Then
'            Beep
'            KeyAscii = 0
'        End If
'    End If
'End If

End Sub

Private Sub picControlBox_Click(index As Integer)

If (index = 0) Then
    Me.WindowState = vbMinimized
ElseIf (index = 1) Then
    Me.WindowState = vbNormal
Else
    SumPreis = 0
    
    If (TaxmusterDBok) Then
        TaxmusterDB.Close
    Else
        If (TM_NAMEN% > 0) Then Close (TM_NAMEN%)
        If (TM_DATEN% > 0) Then Close (TM_DATEN%)
    End If
    
    If (MAG_SPEICHER% > 0) Then Close (MAG_SPEICHER%)
    
    Unload Me
End If

End Sub

Private Sub lblchkKassenRabatt_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("lblchkKassenRabatt_Click")
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

With chkKassenRabatt(index)
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

Private Sub chkKassenRabatt_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("chkKassenRabatt_GotFocus")
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

Call nlCheckBox(chkKassenRabatt(index).Name, index)

Call DefErrPop
End Sub

Private Sub chkKassenRabatt_LostFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("chkKassenRabatt_LostFocus")
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

Call nlCheckBox(chkKassenRabatt(index).Name, index, 0)

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



