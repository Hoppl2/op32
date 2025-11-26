VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlbutton.ocx"
Begin VB.Form frmImportOrIdent 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   3600
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   5295
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5295
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   840
      Picture         =   "ImportOrIdent.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   600
      Picture         =   "ImportOrIdent.frx":00B9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   360
      Picture         =   "ImportOrIdent.frx":016D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdF2 
      Caption         =   "F2"
      Height          =   390
      Left            =   3600
      TabIndex        =   3
      Top             =   2040
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxEdit 
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1931
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   390
      Left            =   3240
      TabIndex        =   2
      Top             =   840
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   1200
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
End
Attribute VB_Name = "frmImportOrIdent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "IMPORTORIDENT.FRM"

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

EditErg% = False
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
Dim i%, ind%, uhr%, st%, min%, falsch%
Dim h2$

With flxEdit
    EditTxt$ = ""
    For i% = 0 To 2
        If (i% >= .Cols) Then Exit For
        If (EditTxt$ <> "") Then EditTxt$ = EditTxt$ + vbTab
        EditTxt$ = EditTxt$ + .TextMatrix(.row, i%)
    Next i%
    EditTxt = .TextMatrix(.row, 0) + vbTab
End With
EditErg% = True
Unload Me

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

Call ZeigeIdente

Call DefErrPop
End Sub

Private Sub flxEdit_DblClick()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxEdit_DblClick")
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

cmdOk.Value = True

Call DefErrPop
End Sub

Private Sub flxEdit_KeyDown(KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxEdit_KeyDown")
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

If (KeyCode = vbKeyF2) And (cmdF2.Enabled) Then
    cmdF2.Value = True
End If

Call DefErrPop
End Sub

Sub flxEdit_RowColChange()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxEdit_RowColChange")
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
Dim i%, OrgCol%
Dim lBack&
Static InRowColChange%

If (InRowColChange) Then
    Call DefErrPop: Exit Sub
End If

InRowColChange = True
With flxEdit
    OrgCol = .col
    
    If (para.Newline) Then
        .SelectionMode = flexSelectionFree
        If (.Visible) Then
            .col = 1
            If (.CellBackColor = vbGreen) Or (.CellBackColor = wpara.nlGreen) Then
                .BackColorSel = .CellBackColor
            Else
                .BackColorSel = wpara.nlFlexBackColorSel  ' RGB(232, 217, 172)
            End If
            .col = 7
            .ColSel = 7
            If (.CellBackColor <> 0) Then
                .col = 0
                .ColSel = 5
            Else
                .col = 0
                .ColSel = .Cols - 1
            End If
            .col = 4
            .ColSel = 7
        End If
    Else
        .col = 1
        If (.CellBackColor = vbGreen) Or (.CellBackColor = wpara.nlGreen) Then
            .BackColorSel = .CellBackColor
        ElseIf (para.Newline) Then
            .BackColorSel = wpara.nlFlexBackColorSel  ' RGB(232, 217, 172)
        Else
            .BackColorSel = vbHighlight
        End If
    
        .col = OrgCol
    End If
End With
InRowColChange = False
    
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
Dim i%, MaxWi%, wi%, wi1%, wi2%

Call wpara.InitFont(Me)

cmdOk.Left = 10000
cmdEsc.Left = 10000
cmdF2.Left = -10000

'If (para.Newline) Then
'    flxEdit.SelectionMode = flexSelectionFree
'End If

EditErg% = False

Call DefErrPop
End Sub

Sub ZeigeIdente()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeIdente")
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
Dim i%, ind%, row%, AnzGilt%, AutIdemNeu%
Dim spBreite&
Dim OrgAvp#, OrgKp#, AVP#, unten#, oben#, grenze#
Dim s$, SQLStr$, kz$, OrgPZN$, pzn$, Titel$

SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + IdentPzn$
'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
On Error Resume Next
TaxeRec.Close
Err.Clear
On Error GoTo DefErr
TaxeRec.open SQLStr, taxeAdoDB.ActiveConn
If (TaxeRec.EOF) Then
    Call DefErrPop: Exit Sub
End If
If (TaxeRec!Identgruppe = 0) Then
    Call DefErrPop: Exit Sub
End If

Titel$ = "Auswahl IDENTE"

With flxEdit
    .Rows = .FixedRows
    
    SQLStr$ = "SELECT * FROM TAXE WHERE IDENTGRUPPE = " + Str$(TaxeRec!Identgruppe)
'    SQLStr = SQLStr + " and Einheit = """ + TaxeRec!einheit + """"
    SQLStr = SQLStr + " and Einheit = '" + TaxeRec!einheit + "'"
    SQLStr = SQLStr + " ORDER BY MENGE,VK"
    'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
    On Error Resume Next
    TaxeRec.Close
    Err.Clear
    On Error GoTo DefErr
    TaxeRec.open SQLStr, taxeAdoDB.ActiveConn

    If (TaxeRec.EOF) Then
        Call DefErrPop: Exit Sub
    End If
    
    While Not (TaxeRec.EOF)

        AVP# = TaxeRec!vk / 100#
    
        s$ = PznString(TaxeRec!pzn) + vbTab + Trim(TaxeRec!Name) + vbTab
        s$ = s$ + TaxeRec!menge + vbTab + TaxeRec!einheit + vbTab
        s$ = s$ + kz$ + vbTab + Format(AVP#, "0.00") + vbTab
        s$ = s$ + TaxeRec!HerstellerKB + vbTab + Trim(TaxeRec!Name) + vbTab + TaxeRec!ArtStatus
        .AddItem s$
        
        If (ArtikelDbOk) Then
            SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + PznString(TaxeRec!pzn)
        '                    SQLStr = SQLStr + " AND LagerKz<>0"
            FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr, 0)
        Else
            FabsErrf% = ass.IndexSearch(0, Format(TaxeRec!pzn, "0000000"), FabsRecno&)
        End If
        If (FabsErrf% = 0) Then
            .FillStyle = flexFillRepeat
            .row = .Rows - 1
            .col = 0
            .RowSel = .row
            .ColSel = .Cols - 1
            .CellFontBold = True
            .FillStyle = flexFillSingle
        End If
    
        TaxeRec.MoveNext
    Wend
    
    .Height = .RowHeight(0) * .Rows + 90
    If (.Height > wpara.WorkAreaHeight) Then
'        .ColWidth(9) = wpara.FrmScrollHeight
        .Height = wpara.WorkAreaHeight - 150
        .Height = ((.Height - 90) \ .RowHeight(0)) * .RowHeight(0) + 90
    End If
    
    With Me
        .Caption = Titel$
        .Height = flxEdit.Height + wpara.FrmCaptionHeight + 60
        .Top = frmAction.Top + (frmAction.Height - .Height) / 2
    End With
    
    .row = 1
    .col = 0
    .ColSel = .Cols - 1
    .Visible = True
End With

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
If (para.Newline) And (Me.Visible) Then
    CurrentX = wpara.NlFlexBackY
    CurrentY = (wpara.NlCaptionY - TextHeight(Caption)) / 2
    ForeColor = vbBlack
    Me.Print Caption

    Beep
    Call flxEdit_RowColChange
End If
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
    ElseIf (KeyAscii = 27) Then
        Call nlcmdEsc_Click
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


