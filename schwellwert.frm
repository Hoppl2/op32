VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSchwellwerte 
   Caption         =   "Rabatt-Tabelle"
   ClientHeight    =   7095
   ClientLeft      =   1410
   ClientTop       =   1020
   ClientWidth     =   10185
   Icon            =   "schwellwert.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   10185
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   6960
      TabIndex        =   32
      Top             =   6480
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   5880
      TabIndex        =   31
      Top             =   6480
      Width           =   855
   End
   Begin TabDlg.SSTab tabSchwellwerte 
      Height          =   5325
      Left            =   120
      TabIndex        =   33
      Top             =   480
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   9393
      _Version        =   327681
      Style           =   1
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
      TabCaption(0)   =   "&1 - Rabatt-Tabellen"
      TabPicture(0)   =   "schwellwert.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "flxRabattTabelle(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "flxRabattTabelle(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "optRabattTabelle(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "optRabattTabelle(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdF5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdF2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&2 - Ausnahmen"
      TabPicture(1)   =   "schwellwert.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtAusnahmenHerst(5)"
      Tab(1).Control(1)=   "txtAusnahmenHerst(4)"
      Tab(1).Control(2)=   "txtAusnahmenHerst(3)"
      Tab(1).Control(3)=   "txtAusnahmenHerst(2)"
      Tab(1).Control(4)=   "txtAusnahmenHerst(1)"
      Tab(1).Control(5)=   "txtAusnahmenWg(9)"
      Tab(1).Control(6)=   "txtAusnahmenWg(8)"
      Tab(1).Control(7)=   "txtAusnahmenWg(7)"
      Tab(1).Control(8)=   "txtAusnahmenWg(6)"
      Tab(1).Control(9)=   "txtAusnahmenWg(5)"
      Tab(1).Control(10)=   "txtAusnahmenWg(4)"
      Tab(1).Control(11)=   "txtAusnahmenWg(3)"
      Tab(1).Control(12)=   "txtAusnahmenWg(2)"
      Tab(1).Control(13)=   "txtAusnahmenWg(1)"
      Tab(1).Control(14)=   "chkAusnahmen(0)"
      Tab(1).Control(15)=   "chkAusnahmen(1)"
      Tab(1).Control(16)=   "chkAusnahmen(2)"
      Tab(1).Control(17)=   "chkAusnahmen(3)"
      Tab(1).Control(18)=   "chkAusnahmen(4)"
      Tab(1).Control(19)=   "chkAusnahmen(5)"
      Tab(1).Control(20)=   "chkAusnahmen(6)"
      Tab(1).Control(21)=   "txtAusnahmenBM"
      Tab(1).Control(22)=   "txtAusnahmenZW"
      Tab(1).Control(23)=   "txtAusnahmenWg(0)"
      Tab(1).Control(24)=   "txtAusnahmenHerst(0)"
      Tab(1).ControlCount=   25
      TabCaption(2)   =   "&3 - Sonstiges"
      TabPicture(2)   =   "schwellwert.frx":0182
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cboSonstiges"
      Tab(2).Control(1)=   "lblSonstiges"
      Tab(2).ControlCount=   2
      Begin VB.TextBox txtAusnahmenHerst 
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   -67800
         MaxLength       =   5
         TabIndex        =   30
         Text            =   "Wwwww"
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox txtAusnahmenHerst 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -68880
         MaxLength       =   5
         TabIndex        =   29
         Text            =   "Wwwww"
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox txtAusnahmenHerst 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -69840
         MaxLength       =   5
         TabIndex        =   28
         Text            =   "Wwwww"
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox txtAusnahmenHerst 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -67920
         MaxLength       =   5
         TabIndex        =   27
         Text            =   "Wwwww"
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox txtAusnahmenHerst 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -68880
         MaxLength       =   5
         TabIndex        =   26
         Text            =   "Wwwww"
         Top             =   4320
         Width           =   735
      End
      Begin VB.TextBox txtAusnahmenWg 
         Enabled         =   0   'False
         Height          =   375
         Index           =   9
         Left            =   -65760
         MaxLength       =   2
         TabIndex        =   19
         Text            =   "99"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtAusnahmenWg 
         Enabled         =   0   'False
         Height          =   375
         Index           =   8
         Left            =   -66720
         MaxLength       =   2
         TabIndex        =   18
         Text            =   "99"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtAusnahmenWg 
         Enabled         =   0   'False
         Height          =   375
         Index           =   7
         Left            =   -67560
         MaxLength       =   2
         TabIndex        =   17
         Text            =   "99"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtAusnahmenWg 
         Enabled         =   0   'False
         Height          =   375
         Index           =   6
         Left            =   -68400
         MaxLength       =   2
         TabIndex        =   16
         Text            =   "99"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtAusnahmenWg 
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   -69240
         MaxLength       =   2
         TabIndex        =   15
         Text            =   "99"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtAusnahmenWg 
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   -65760
         MaxLength       =   2
         TabIndex        =   14
         Text            =   "99"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtAusnahmenWg 
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   -66720
         MaxLength       =   2
         TabIndex        =   13
         Text            =   "99"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtAusnahmenWg 
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   -67560
         MaxLength       =   2
         TabIndex        =   12
         Text            =   "99"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtAusnahmenWg 
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   -68400
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "99"
         Top             =   1800
         Width           =   735
      End
      Begin VB.ComboBox cboSonstiges 
         Height          =   315
         Left            =   -72240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown-Liste
         TabIndex        =   35
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox chkAusnahmen 
         Caption         =   "&Kühlartikel"
         Height          =   375
         Index           =   0
         Left            =   -72480
         TabIndex        =   6
         Top             =   840
         Width           =   2535
      End
      Begin VB.CheckBox chkAusnahmen 
         Caption         =   "&Sonderangebote"
         Height          =   375
         Index           =   1
         Left            =   -72600
         TabIndex        =   7
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CheckBox chkAusnahmen 
         Caption         =   "&Betäubungsmittel"
         Height          =   375
         Index           =   2
         Left            =   -72600
         TabIndex        =   8
         Top             =   1560
         Width           =   2535
      End
      Begin VB.CheckBox chkAusnahmen 
         Caption         =   " &Warengruppen"
         Height          =   375
         Index           =   3
         Left            =   -72600
         TabIndex        =   9
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CheckBox chkAusnahmen 
         Caption         =   "bis Bestell&menge"
         Height          =   375
         Index           =   4
         Left            =   -72600
         TabIndex        =   20
         Top             =   3240
         Width           =   2535
      End
      Begin VB.CheckBox chkAusnahmen 
         Caption         =   "bis &Zeilenwert"
         Height          =   375
         Index           =   5
         Left            =   -72600
         TabIndex        =   22
         Top             =   3720
         Width           =   2535
      End
      Begin VB.CheckBox chkAusnahmen 
         Caption         =   "&Hersteller"
         Height          =   375
         Index           =   6
         Left            =   -72480
         TabIndex        =   24
         Top             =   4320
         Width           =   2535
      End
      Begin VB.TextBox txtAusnahmenBM 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -69840
         MaxLength       =   3
         TabIndex        =   21
         Text            =   "9999"
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox txtAusnahmenZW 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -69960
         MaxLength       =   7
         TabIndex        =   23
         Text            =   "9999.99"
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox txtAusnahmenWg 
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   -69240
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "99"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtAusnahmenHerst 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -69960
         MaxLength       =   5
         TabIndex        =   25
         Text            =   "Wwwww"
         Top             =   4320
         Width           =   735
      End
      Begin VB.CommandButton cmdF2 
         Caption         =   "Einfügen (F2)"
         Height          =   450
         Left            =   4200
         TabIndex        =   4
         Top             =   2280
         Width           =   1200
      End
      Begin VB.CommandButton cmdF5 
         Caption         =   "Entfernen (F5)"
         Height          =   450
         Left            =   4200
         TabIndex        =   5
         Top             =   2880
         Width           =   1200
      End
      Begin VB.OptionButton optRabattTabelle 
         Caption         =   "Einkaufswert-Tabelle"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   0
         Top             =   960
         Width           =   2895
      End
      Begin VB.OptionButton optRabattTabelle 
         Caption         =   "BM-/Zeilenrabatt - Tabelle"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   1
         Top             =   1320
         Width           =   2895
      End
      Begin MSFlexGridLib.MSFlexGrid flxRabattTabelle 
         Height          =   2040
         Index           =   0
         Left            =   2280
         TabIndex        =   2
         Top             =   2040
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   3598
         _Version        =   65541
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         BackColorBkg    =   -2147483633
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxRabattTabelle 
         Height          =   2040
         Index           =   1
         Left            =   3120
         TabIndex        =   3
         Top             =   2040
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   3598
         _Version        =   65541
         Rows            =   0
         Cols            =   0
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         BackColorBkg    =   -2147483633
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   2
         GridLines       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin VB.Label lblSonstiges 
         Caption         =   "Umsatz speichern bei Lieferant"
         Height          =   615
         Left            =   -74760
         TabIndex        =   34
         Top             =   720
         Width           =   4455
      End
   End
End
Attribute VB_Name = "frmSchwellwerte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim VerglStr$(4)

Private Const DefErrModul = "SCHWELLWERT.FRM"

Private Sub chkAusnahmen_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("chkAusnahmen_Click")
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
Dim i%

If (index = 3) Then
    For i% = 0 To 9
        txtAusnahmenWg(i%).Enabled = chkAusnahmen(index).Value
    Next i%
ElseIf (index = 4) Then
    txtAusnahmenBM.Enabled = chkAusnahmen(index).Value
ElseIf (index = 5) Then
    txtAusnahmenZW.Enabled = chkAusnahmen(index).Value
ElseIf (index = 6) Then
    For i% = 0 To 5
        txtAusnahmenHerst(i%).Enabled = chkAusnahmen(index).Value
    Next i%
End If

Call clsError.DefErrPop
End Sub

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
Dim i%, j%, ind%

ind% = 0
If (flxRabattTabelle(1).Visible) Then ind% = 1

With flxRabattTabelle(ind%)
    For j% = (.Rows - 2) To .row Step -1
        For i% = 0 To .Cols - 1
            .TextMatrix(j% + 1, i%) = .TextMatrix(j%, i%)
        Next i%
    Next j%
    For i% = 0 To .Cols - 1
        .TextMatrix(.row, i%) = ""
    Next i%
End With

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
Dim i%, ind%

ind% = 0
If (flxRabattTabelle(1).Visible) Then ind% = 1

With flxRabattTabelle(ind%)
    For i% = 0 To .Cols - 1
        .TextMatrix(.row, i%) = ""
    Next i%
End With

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
Dim ind%

If (ActiveControl.Name = cmdOk.Name) Then
    Call SpeicherRabattTabelle
    Unload Me
ElseIf (ActiveControl.Name = flxRabattTabelle(0).Name) Then
    ind% = ActiveControl.index
    If (ind% = 1) And (flxRabattTabelle(1).col = 2) Then
        Call EditSchwellwerteLst
    Else
        Call EditSchwellwerteTxt
    End If
End If

Call clsError.DefErrPop
End Sub

Private Sub flxRabattTabelle_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("flxRabattTabelle_KeyDown")
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

If (KeyCode = vbKeyF2) Then
    cmdF2.Value = True
ElseIf (KeyCode = vbKeyF5) Then
    cmdF5.Value = True
End If

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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%, erg%, xpos%, ydiff%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%, maxSp%, val2%, iLief%
Dim MenuHeight&, ScrollHeight&, val1&
Dim val#
Dim h$, h2$, FormStr$
Dim c As Control

LifZus1.GetRecord (AktWumsatzLief% + 1)

VerglStr$(0) = "<"
VerglStr$(1) = "<="
VerglStr$(2) = "="
VerglStr$(3) = ">="
VerglStr$(4) = ">"

tabSchwellwerte.Tab = 1
For i% = 1 To 9
'    Load txtAusnahmenWg(i%)
    txtAusnahmenWg(i%).text = txtAusnahmenWg(0).text
    txtAusnahmenWg(i%).Visible = True
    txtAusnahmenWg(i%).TabIndex = txtAusnahmenWg(i% - 1).TabIndex + 1
Next i%
For i% = 1 To 5
'    Load txtAusnahmenHerst(i%)
    txtAusnahmenHerst(i%).text = txtAusnahmenHerst(0).text
    txtAusnahmenHerst(i%).Visible = True
    txtAusnahmenHerst(i%).TabIndex = txtAusnahmenHerst(i% - 1).TabIndex + 1
Next i%

Call wPara1.InitFont(Me)

On Error Resume Next
For Each c In Controls
    If (TypeOf c Is TextBox) Then
        c.Width = TextWidth(c.text) + 90
        c.text = ""
    End If
Next
On Error GoTo DefErr


tabSchwellwerte.Left = wPara1.LinksX
tabSchwellwerte.Top = wPara1.TitelY


Font.Name = wPara1.FontName(1)
Font.Size = wPara1.FontSize(1)


cmdOk.Width = wPara1.ButtonX
cmdOk.Height = wPara1.ButtonY
cmdEsc.Width = cmdOk.Width
cmdEsc.Height = cmdOk.Height
cmdF5.Width = TextWidth(cmdF5.Caption) + 150
cmdF5.Height = cmdOk.Height
cmdF2.Width = cmdF5.Width
cmdF2.Height = cmdOk.Height



'Font.Bold = False   ' True

tabSchwellwerte.Tab = 0
For i% = 0 To 1
    optRabattTabelle(i%).Left = wPara1.LinksX
Next i%
optRabattTabelle(0).Top = 900   '2 * wPara1.TitelY
For i% = 1 To 1
    optRabattTabelle(i%).Top = optRabattTabelle(i% - 1).Top + optRabattTabelle(i% - 1).Height + 60
Next i%

With flxRabattTabelle(1)
    .Cols = 5
    .Rows = 6
    .FixedRows = 1
    .FixedCols = 0
    .row = 1
    
    .Top = optRabattTabelle(1).Top + optRabattTabelle(1).Height + 210
    .Left = wPara1.LinksX
    .Height = .RowHeight(0) * .Rows + 90
        
    .FormatString = ">AbBM|>BisBM|>Vgl|>zWert|>Rabatt (%)"
    .SelectionMode = flexSelectionFree
    
    .ColWidth(0) = TextWidth("AbBM  ")
    .ColWidth(1) = TextWidth("BisBM  ")
    .ColWidth(2) = TextWidth("Vgl  ")
    .ColWidth(3) = TextWidth("999999999  ")
    .ColWidth(4) = TextWidth("Rabatt (%)    ")

    spBreite% = 0
    For i% = 0 To (.Cols - 1)
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
    
    For i% = 0 To 4
        val2% = LifZus1.AbBM(i%)
        If (val2% > 0) Then .TextMatrix(i% + 1, 0) = Format(val2%, "0")
        val2% = LifZus1.BisBM(i%)
        If (val2% > 0) Then .TextMatrix(i% + 1, 1) = Format(val2%, "0")
        val2% = LifZus1.Vergleich(i%)
'        If (val2% > 0) Then .TextMatrix(i% + 1, 2) = Format(val2%, "0")
        If (val2% < 5) Then .TextMatrix(i% + 1, 2) = VerglStr$(val2%)
        val1& = LifZus1.zWert(i%)
        If (val1& > 0) Then .TextMatrix(i% + 1, 3) = Format(val1&, "0")
        val# = LifZus1.BmEkRabatt(i%)
        If (val# > 0) Then .TextMatrix(i% + 1, 4) = Format(val#, "0.00")
    Next i%
End With

With flxRabattTabelle(0)
    .Cols = 2
    .Rows = 6
    .FixedRows = 1
    .FixedCols = 0
    .row = 1
    
    .Top = flxRabattTabelle(1).Top
    .Left = flxRabattTabelle(1).Left
    .Height = flxRabattTabelle(1).Height
    .Width = flxRabattTabelle(1).Width
        
    .FormatString = ">Schwellwert (DM)|>Rabatt (%)"
    .SelectionMode = flexSelectionFree
    
    For i% = 0 To 1
        .ColWidth(i%) = (.Width - 90) / 2
    Next i%
    
'    .ColWidth(0) = TextWidth("Schwellwert (DM)     ")
'    .ColWidth(1) = TextWidth("Rabatt (%)    ")
'
'    spBreite% = 0
'    For i% = 0 To (.Cols - 1)
'        spBreite% = spBreite% + .ColWidth(i%)
'    Next i%
'    .Width = spBreite% + 90
    
    For i% = 0 To 4
        val1& = LifZus1.Schwellwert(i%)
        If (val1& > 0) Then .TextMatrix(i% + 1, 0) = Format(val1&, "0")
        val# = LifZus1.Rabatt(i%)
        If (val# > 0) Then .TextMatrix(i% + 1, 1) = Format(val#, "0.00")
    Next i%
End With

If (LifZus1.TabTyp) Then
    optRabattTabelle(1).Value = True
'    flxRabattTabelle(1).Visible = True
Else
    optRabattTabelle(0).Value = True
'    flxRabattTabelle(0).Visible = True
End If

cmdF2.Top = flxRabattTabelle(0).Top
cmdF2.Left = flxRabattTabelle(0).Left + flxRabattTabelle(0).Width + 150
cmdF5.Top = cmdF2.Top + cmdF2.Height + 90
cmdF5.Left = cmdF2.Left




tabSchwellwerte.Tab = 1
For i% = 0 To 6
    chkAusnahmen(i%).Value = LifZus1.AusnahmenKz(i%)
Next i%
For i% = 0 To 9
    txtAusnahmenWg(i%).text = Format(LifZus1.AusnahmenWg(i%), "0")
Next i%
txtAusnahmenBM.text = Format(LifZus1.AusnahmenBM, "0")
txtAusnahmenZW.text = Format(LifZus1.AusnahmenZW, "0.00")
For i% = 0 To 5
    txtAusnahmenHerst(i%).text = LifZus1.AusnahmenHerst(i%)
Next i%

For i% = 0 To 6
    chkAusnahmen(i%).Left = wPara1.LinksX
Next i%
chkAusnahmen(0).Top = optRabattTabelle(0).Top
For i% = 1 To 6
    chkAusnahmen(i%).Top = chkAusnahmen(i% - 1).Top + chkAusnahmen(i% - 1).Height + 150
    If (i% = 4) Then chkAusnahmen(i%).Top = chkAusnahmen(i%).Top + chkAusnahmen(i% - 1).Height + 150
Next i%

xpos% = chkAusnahmen(2).Left + chkAusnahmen(2).Width
ydiff% = (txtAusnahmenBM.Height - chkAusnahmen(0).Height) / 2

txtAusnahmenWg(0).Left = xpos%
txtAusnahmenWg(0).Top = chkAusnahmen(3).Top - ydiff%
For i% = 1 To 4
    txtAusnahmenWg(i%).Left = txtAusnahmenWg(i% - 1).Left + txtAusnahmenWg(i% - 1).Width + 45
    txtAusnahmenWg(i%).Top = txtAusnahmenWg(0).Top
Next i%
txtAusnahmenWg(5).Left = xpos%
txtAusnahmenWg(5).Top = txtAusnahmenWg(0).Top + txtAusnahmenWg(0).Height + 60
For i% = 6 To 9
    txtAusnahmenWg(i%).Left = txtAusnahmenWg(i% - 1).Left + txtAusnahmenWg(i% - 1).Width + 45
    txtAusnahmenWg(i%).Top = txtAusnahmenWg(5).Top
Next i%

txtAusnahmenBM.Left = xpos%
txtAusnahmenZW.Left = xpos%
txtAusnahmenBM.Top = chkAusnahmen(4).Top - ydiff%
txtAusnahmenZW.Top = chkAusnahmen(5).Top - ydiff%

txtAusnahmenHerst(0).Left = xpos%
txtAusnahmenHerst(0).Top = chkAusnahmen(6).Top - ydiff%
For i% = 1 To 2
    txtAusnahmenHerst(i%).Left = txtAusnahmenHerst(i% - 1).Left + txtAusnahmenHerst(i% - 1).Width + 45
    txtAusnahmenHerst(i%).Top = txtAusnahmenHerst(0).Top
Next i%
txtAusnahmenHerst(3).Left = xpos%
txtAusnahmenHerst(3).Top = txtAusnahmenHerst(0).Top + txtAusnahmenHerst(0).Height + 60
For i% = 4 To 5
    txtAusnahmenHerst(i%).Left = txtAusnahmenHerst(i% - 1).Left + txtAusnahmenHerst(i% - 1).Width + 45
    txtAusnahmenHerst(i%).Top = txtAusnahmenHerst(3).Top
Next i%



tabSchwellwerte.Tab = 2
lblSonstiges.Left = wPara1.LinksX
lblSonstiges.Top = 900   '2 * wPara1.TitelY

cboSonstiges.Left = lblSonstiges.Left + lblSonstiges.Width + 150
ydiff% = (cboSonstiges.Height - lblSonstiges.Height) / 2
cboSonstiges.Top = lblSonstiges.Top - ydiff%

If (LifZus1.WumsatzLief = 0) Then
    iLief% = AktWumsatzLief%
Else
    iLief% = LifZus1.WumsatzLief
End If

With cboSonstiges
    .Clear
    For i% = 1 To 200
        Call Lif1.GetRecord(i% + 1)
        h$ = Lif1.kurz
        h$ = UCase(Trim$(h$))
        If (h$ <> "") Then
            If (Asc(Left$(h$, 1)) >= 32) Then
                h$ = h$ + " (" + Mid$(Str$(i%), 2) + ")"
                .AddItem h$
                If (i% = iLief%) Then h2$ = h$
            End If
        End If
    Next i%
    For i% = 0 To (.ListCount - 1)
        .ListIndex = i%
        If (.text = h2$) Then Exit For
    Next i%
End With
    



tabSchwellwerte.Height = txtAusnahmenHerst(3).Top + txtAusnahmenHerst(3).Height + 2 * wPara1.TitelY
tabSchwellwerte.Width = cmdF2.Left + cmdF2.Width + 2 * wPara1.LinksX


cmdOk.Top = tabSchwellwerte.Top + tabSchwellwerte.Height + 150
cmdEsc.Top = cmdOk.Top


Me.Width = tabSchwellwerte.Left + tabSchwellwerte.Width + 2 * wPara1.LinksX
Me.Height = cmdOk.Top + cmdOk.Height + wPara1.TitelY + 90 + wPara1.FrmCaptionHeight
Caption = AktWumsatzTyp$

cmdOk.Left = (Me.ScaleWidth - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

tabSchwellwerte.Tab = 0
Call TabDisable
Call TabEnable(0)

Call clsError.DefErrPop
End Sub

Sub EditSchwellwerteTxt()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("EditSchwellwerteTxt")
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
Dim EditRow%, EditCol%, ind%, arow%
Dim dVal#
Dim h2$

ind% = ActiveControl.index

EditRow% = flxRabattTabelle(ind%).row
EditCol% = flxRabattTabelle(ind%).col

EditModus% = 0
If ((ind% = 0) And (EditCol% = 1)) Or ((ind% = 1) And (EditCol% = 4)) Then
    EditModus% = 4
End If

With flxRabattTabelle(ind%)
    arow% = .row
    .row = 0
    .CellFontBold = True
    .row = arow%
End With
            
Load frmEdit2

With frmEdit2
    .Left = tabSchwellwerte.Left + flxRabattTabelle(ind%).Left + flxRabattTabelle(ind%).ColPos(EditCol%) + 45
    .Left = .Left + Me.Left + wPara1.FrmBorderHeight
    .Top = tabSchwellwerte.Top + flxRabattTabelle(ind%).Top + EditRow% * flxRabattTabelle(ind%).RowHeight(0)
    .Top = .Top + Me.Top + wPara1.FrmBorderHeight + wPara1.FrmCaptionHeight
    .Width = flxRabattTabelle(ind%).ColWidth(EditCol%)
    .Height = frmEdit2.txtEdit.Height 'flxarbeit(0).RowHeight(1)
End With
With frmEdit2.txtEdit
    .Width = frmEdit2.ScaleWidth
    .Left = 0
    .Top = 0
    h2$ = flxRabattTabelle(ind%).TextMatrix(EditRow%, EditCol%)
    .text = h2$
    .BackColor = vbWhite
    .Visible = True
End With

frmEdit2.Show 1
           
With flxRabattTabelle(ind%)
    arow% = .row
    .row = 0
    .CellFontBold = False
    .row = arow%
            
    If (EditErg%) Then
        dVal# = val(EditTxt$)
        If (EditModus% = 0) Then
            h2$ = Format(dVal#, "0")
        Else
            h2$ = Format(dVal#, "0.00")
        End If
        .TextMatrix(EditRow%, EditCol%) = h2$
        If (.col < .Cols - 2) Then .col = .col + 1
    End If
End With

Call clsError.DefErrPop
End Sub

Sub EditSchwellwerteLst()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("EditSchwellwerteLst")
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
Dim i%, EditRow%, EditCol%, ind%, arow%
Dim dVal#
Dim h2$, s$

ind% = ActiveControl.index

EditRow% = flxRabattTabelle(ind%).row
EditCol% = flxRabattTabelle(ind%).col

With flxRabattTabelle(ind%)
    arow% = .row
    .row = 0
    .CellFontBold = True
    .row = arow%
End With
            
Load frmEdit2

With frmEdit2
    .Left = tabSchwellwerte.Left + flxRabattTabelle(ind%).Left + flxRabattTabelle(ind%).ColPos(EditCol%) + 45
    .Left = .Left + Me.Left + wPara1.FrmBorderHeight
    .Top = tabSchwellwerte.Top + flxRabattTabelle(ind%).Top + flxRabattTabelle(ind%).RowPos(flxRabattTabelle(ind%).TopRow)
    .Top = .Top + Me.Top + wPara1.FrmBorderHeight + wPara1.FrmCaptionHeight
    .Width = flxRabattTabelle(ind%).ColWidth(EditCol%)
    .Height = flxRabattTabelle(ind%).Height - flxRabattTabelle(ind%).RowPos(flxRabattTabelle(ind%).TopRow)
End With

With frmEdit2.lstEdit
    .Height = frmEdit2.ScaleHeight
    frmEdit2.Height = .Height
    .Width = frmEdit2.ScaleWidth
    .Left = 0
    .Top = 0
    
    .Clear
    For i% = 0 To 4
        .AddItem VerglStr$(i%)
    Next i%
    
    .ListIndex = 0
    s$ = RTrim$(flxRabattTabelle(ind%).TextMatrix(EditRow%, EditCol%))
    If (s$ <> "") Then
        For i% = 0 To (.ListCount - 1)
            .ListIndex = i%
            h2$ = .text
            If (s$ = h2$) Then
                Exit For
            End If
        Next i%
    End If
    
    .Visible = True
End With
   
frmEdit2.Show 1
           
With flxRabattTabelle(ind%)
    arow% = .row
    .row = 0
    .CellFontBold = False
    .row = arow%
End With
            

If (EditErg%) Then
    h2$ = EditTxt$
    With flxRabattTabelle(ind%)
        .TextMatrix(EditRow%, EditCol%) = h2$
        If (.col < .Cols - 2) Then .col = .col + 1
    End With
End If

Call clsError.DefErrPop
End Sub

Sub SpeicherRabattTabelle()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("SpeicherRabattTabelle")
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
Dim i%, j%, pos%, MitAusnahmen%, vergl%, ind%
Dim Schwell$, Rabatt$, h$

LifZus1.GetRecord (AktWumsatzLief% + 1)

If (optRabattTabelle(0).Value) Then
    LifZus1.TabTyp = 0
Else
    LifZus1.TabTyp = 1
End If

With flxRabattTabelle(0)
    pos% = 0
    For i% = 1 To (.Rows - 1)
        Schwell$ = Trim(.TextMatrix(i%, 0))
        Rabatt$ = Trim(.TextMatrix(i%, 1))
        If (Rabatt$ <> "") Then
            LifZus1.Schwellwert(pos%) = val(Schwell$)
            LifZus1.Rabatt(pos%) = CDbl(Rabatt$)
            pos% = pos% + 1
        End If
    Next i%
    Do
        If (pos% >= 5) Then Exit Do
        LifZus1.Schwellwert(pos%) = 0#
        LifZus1.Rabatt(pos%) = 0!
        pos% = pos% + 1
    Loop
End With

With flxRabattTabelle(1)
    pos% = 0
    For i% = 1 To (.Rows - 1)
        Rabatt$ = Trim(.TextMatrix(i%, 4))
        If (Rabatt$ <> "") Then
            LifZus1.AbBM(pos%) = val(.TextMatrix(i%, 0))
            LifZus1.BisBM(pos%) = val(.TextMatrix(i%, 1))
            
            vergl% = 0
            h$ = .TextMatrix(i%, 2)
            For j% = 0 To 4
                If (h$ = VerglStr$(j%)) Then
                    vergl% = j%
                    Exit For
                End If
            Next j%
            LifZus1.Vergleich(pos%) = vergl%
            
            LifZus1.zWert(pos%) = val(.TextMatrix(i%, 3))
            LifZus1.BmEkRabatt(pos%) = val(.TextMatrix(i%, 4))
            pos% = pos% + 1
        End If
    Next i%
    Do
        If (pos% >= 5) Then Exit Do
        LifZus1.AbBM(pos%) = 0
        LifZus1.BisBM(pos%) = 0
        LifZus1.Vergleich(pos%) = 0
        LifZus1.zWert(pos%) = 0#
        LifZus1.BmEkRabatt(pos%) = 0#
        pos% = pos% + 1
    Loop
End With

MitAusnahmen% = False
For i% = 0 To 6
    LifZus1.AusnahmenKz(i%) = Abs(chkAusnahmen(i%).Value)
    If (chkAusnahmen(i%).Value) Then MitAusnahmen% = True
Next i%
LifZus1.HatAusnahmen = Abs(MitAusnahmen%)

For i% = 0 To 9
    LifZus1.AusnahmenWg(i%) = val(txtAusnahmenWg(i%).text)
Next i%
LifZus1.AusnahmenBM = val(txtAusnahmenBM.text)
LifZus1.AusnahmenZW = val(txtAusnahmenZW.text)
For i% = 0 To 5
    LifZus1.AusnahmenHerst(i%) = txtAusnahmenHerst(i%).text
Next i%

h$ = LTrim$(RTrim$(cboSonstiges.text))
If (h$ <> "") Then
    ind% = InStr(h$, "(")
    If (ind% > 0) Then
        h$ = Mid$(h$, ind% + 1)
        ind% = InStr(h$, ")")
        h$ = Left$(h$, ind% - 1)
LifZus1.WumsatzLief = val(h$)
    End If
End If

LifZus1.PutRecord (AktWumsatzLief% + 1)

Call clsError.DefErrPop
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_KeyPress")
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

If (TypeOf ActiveControl Is TextBox) Then
    If (ActiveControl.Name <> txtAusnahmenHerst(0).Name) Then
        If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
'        If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And ((EditModus% <> 4) Or (Chr$(KeyAscii) <> ".")) Then
        If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And ((ActiveControl.Name <> txtAusnahmenZW.Name) Or (Chr$(KeyAscii) <> ".")) Then
            Beep
            KeyAscii = 0
        End If
    End If
End If

Call clsError.DefErrPop
End Sub

Private Sub optRabattTabelle_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtAusnahmenWg_GotFocus")
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

flxRabattTabelle(index).Visible = True
flxRabattTabelle((index + 1) Mod 2).Visible = False

Call clsError.DefErrPop
End Sub

Private Sub txtAusnahmenWg_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtAusnahmenWg_GotFocus")
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

With txtAusnahmenWg(index)
    .SelStart = 0
    .SelLength = Len(.text)
End With

Call clsError.DefErrPop
End Sub

Private Sub txtAusnahmenBM_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtAusnahmenBM_GotFocus")
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

With txtAusnahmenBM
    .SelStart = 0
    .SelLength = Len(.text)
End With

Call clsError.DefErrPop
End Sub

Private Sub txtAusnahmenZW_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtAusnahmenZW_GotFocus")
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
Dim i%
Dim h$

With txtAusnahmenZW
    h$ = .text
    For i% = 1 To Len(h$)
        If (Mid$(h$, i%, 1) = ",") Then Mid$(h$, i%, 1) = "."
    Next i%
    .text = h$
    .SelStart = 0
    .SelLength = Len(.text)
End With

Call clsError.DefErrPop
End Sub

Private Sub txtAusnahmenHerst_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtAusnahmenHerst_GotFocus")
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

With txtAusnahmenHerst(index)
    .SelStart = 0
    .SelLength = Len(.text)
End With

Call clsError.DefErrPop
End Sub

Private Sub tabSchwellwerte_Click(PreviousTab As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("tabSchwellwerte_Click")
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
If (tabSchwellwerte.Visible = False) Then Call clsError.DefErrPop: Exit Sub

Call TabDisable
Call TabEnable(tabSchwellwerte.Tab)

Call clsError.DefErrPop
End Sub

Sub TabDisable()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("TabDisable")
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
Dim i%

For i% = 0 To 1
    flxRabattTabelle(i%).Visible = False
    optRabattTabelle(i%).Visible = False
Next i%
cmdF2.Visible = False
cmdF5.Visible = False


For i% = 0 To 6
    chkAusnahmen(i%).Visible = False
Next i%
For i% = 0 To 9
    txtAusnahmenWg(i%).Visible = False
Next i%
txtAusnahmenBM.Visible = False
txtAusnahmenZW.Visible = False
For i% = 0 To 5
    txtAusnahmenHerst(i%).Visible = False
Next i%

lblSonstiges.Visible = False
cboSonstiges.Visible = False

Call clsError.DefErrPop
End Sub

Sub TabEnable(hTab%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("TabEnable")
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
Dim i%

If (hTab% = 0) Then
    For i% = 0 To 1
        flxRabattTabelle(i%).Visible = True
        optRabattTabelle(i%).Visible = True
    Next i%
    cmdF2.Visible = True
    cmdF5.Visible = True
ElseIf (hTab% = 1) Then
    For i% = 0 To 6
        chkAusnahmen(i%).Visible = True
    Next i%
    For i% = 0 To 9
        txtAusnahmenWg(i%).Visible = True
    Next i%
    txtAusnahmenBM.Visible = True
    txtAusnahmenZW.Visible = True
    For i% = 0 To 5
        txtAusnahmenHerst(i%).Visible = True
    Next i%
    chkAusnahmen(0).SetFocus
Else
    lblSonstiges.Visible = True
    cboSonstiges.Visible = True
    cboSonstiges.SetFocus
End If

Call clsError.DefErrPop
End Sub



