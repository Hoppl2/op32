VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlbutton.ocx"
Begin VB.Form frmMarktPzn 
   AutoRedraw      =   -1  'True
   Caption         =   "Anzeige der Markt-PZN"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6600
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   4080
      Picture         =   "MarktPzn.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   3840
      Picture         =   "MarktPzn.frx":00B9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   3600
      Picture         =   "MarktPzn.frx":016D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2640
      TabIndex        =   1
      Top             =   3600
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxTaxMuster 
      Height          =   2700
      Index           =   0
      Left            =   360
      TabIndex        =   0
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
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
End
Attribute VB_Name = "frmMarktPzn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iEditModus%

Private Const DefErrModul = "MARKTPZN.FRM"

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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%, lief%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%, FeldInd%
Dim iAdd%, iAdd2%
Dim h$, h2$, FormStr$
Dim c As Control

iEditModus = 1

FormErg% = False

LennartzPzn = 0

Call wpara.InitFont(Me)

With flxTaxMuster(0)
    .Rows = 2
    .FixedRows = 1
    .FormatString = "<|<Pzn|<Bezeichnung|>Menge|<Meh|>EK|>Menge2^Menge|<Meh2^Meh|>Dichte|>Menge3^Menge|<Meh3^Meh|^Umrechn.|>CalcEK|^Berechnung|^HerKB"
    .Rows = 1
    
    .ColWidth(0) = TextWidth(String(4, "A"))
    .ColWidth(1) = TextWidth(String(8, "9"))
    .ColWidth(2) = TextWidth(String(25, "A"))
    .ColWidth(3) = TextWidth(String(5, "A"))
    .ColWidth(4) = TextWidth(String(3, "A"))
    .ColWidth(5) = TextWidth(String(8, "9"))
    .ColWidth(6) = TextWidth(String(5, "A"))
    .ColWidth(7) = TextWidth(String(3, "A"))
    .ColWidth(8) = TextWidth(String(6, "9"))
    .ColWidth(9) = TextWidth(String(5, "A"))
    .ColWidth(10) = TextWidth(String(3, "A"))
    .ColWidth(11) = TextWidth(String(10, "9"))
    .ColWidth(12) = TextWidth(String(8, "9"))
    .ColWidth(13) = TextWidth(String(15, "9"))
    .ColWidth(14) = TextWidth(String(6, "A"))
    
    For i = 0 To 14
        .ColWidth(i) = .ColWidth(i) + TextWidth(String(1, "A"))
    Next i
   
    Breite1% = 0
    For i% = 0 To (.Cols - 1)
        Breite1% = Breite1% + .ColWidth(i%)
    Next i%
    .Width = Breite1% + 90
    .Height = .RowHeight(0) * 10 + 90
    
    .Top = wpara.TitelY
    .Left = wpara.LinksX
   
'    Call flxTaxMusterBefuellen
End With

Font.Bold = False   ' True

cmdEsc.Top = flxTaxMuster(0).Top + flxTaxMuster(0).Height + 300

Me.Width = flxTaxMuster(0).Left + flxTaxMuster(0).Width + 2 * wpara.LinksX

cmdEsc.Width = wpara.ButtonX
cmdEsc.Height = wpara.ButtonY

cmdEsc.Left = (Me.Width - cmdEsc.Width) / 2

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

If (para.Newline) Then
    iAdd = wpara.NlFlexBackY
    iAdd2 = wpara.NlCaptionY
    
    With flxTaxMuster(0)
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
        
'        iArbeitAnzzeilen = 20
'        .Height = .RowHeight(0) * iArbeitAnzzeilen%
    End With
    
    Width = Width + 2 * iAdd
    Height = Height + 5 * iAdd

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
        .Top = flxTaxMuster(0).Top + flxTaxMuster(0).Height + iAdd + 600
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .Default = cmdEsc.Default
        .Cancel = cmdEsc.Cancel
        .Visible = True
    End With
    cmdEsc.Visible = False

    Me.Height = nlcmdEsc.Top + nlcmdEsc.Height + wpara.FrmCaptionHeight + 450

    nlcmdEsc.Left = (Me.ScaleWidth - nlcmdEsc.Width) / 2

    Call wpara.NewLineWindow(Me, nlcmdEsc.Top)
    With flxTaxMuster(0)
        RoundRect hdc, (.Left - iAdd) / Screen.TwipsPerPixelX, (.Top - iAdd) / Screen.TwipsPerPixelY, (.Left + .Width + iAdd) / Screen.TwipsPerPixelX, (.Top + .Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
    End With

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

                With c.Container
                    .ForeColor = RGB(180, 180, 180) ' vbWhite
                    .FillStyle = vbSolid
                    .FillColor = c.BackColor

                    RoundRect .hdc, (c.Left - 60) / Screen.TwipsPerPixelX, (c.Top - 30) / Screen.TwipsPerPixelY, (c.Left + c.Width + 60) / Screen.TwipsPerPixelX, (c.Top + c.Height + 15) / Screen.TwipsPerPixelY, 10, 10
                End With
            End If
        End If
    Next
    On Error GoTo DefErr
Else
    nlcmdEsc.Visible = False
End If

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

With flxTaxMuster(0)
    .col = 2
    .col = 1
    .HighLight = flexHighlightNever
End With

Call ShowMarktPzn

Call DefErrPop
End Sub

Private Sub ShowMarktPzn()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ShowMarktPzn")
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
Dim sSQL$, pzn&
Dim dDichte#

dDichte = 0
pzn = 0

sSQL = "Select GRP_HA.*, HA.Dichte FROM GRP_HA LEFT JOIN HA ON GRP_HA.PZN_2=HA.PZN WHERE GRP_HA.PZN_1=" + SollPzn
On Error Resume Next
ABDA_Komplett_Rec.Close
Err.Clear
On Error GoTo DefErr
ABDA_Komplett_Rec.Open sSQL, ABDA_Komplett_Conn
If (ABDA_Komplett_Rec.EOF = False) Then
    pzn = CheckNullLong(ABDA_Komplett_Rec!PZN_2)
    dDichte = xVal(CheckNullStr(ABDA_Komplett_Rec!Dichte))
    If (dDichte = 0) Then
        dDichte = 1
    Else
        dDichte = dDichte / 10000
    End If
End If
ABDA_Komplett_Rec.Close

If (pzn > 0) Then
    Dim i%
    Dim GStufe#(1), Meh$(1), h$

'            .Dichte = dDichte
    SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + SollPzn
    FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr)
    If (FabsErrf% = 0) Then
'                            ast.GetRecord (FabsRecno& + 1)
    Else
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + SollPzn
        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        On Error Resume Next
        TaxeRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
        If (TaxeRec.EOF = False) Then
            Call Taxe2ast(SollPzn)
            FabsErrf% = 0
        End If
    End If
        
    If (FabsErrf% = 0) Then
        With ast
            h = "FAM" + vbTab + .pzn + vbTab + .kurz + vbTab + .meng + vbTab + .Meh + vbTab
            h = h + Format(.aep, "0.00") + vbTab
            flxTaxMuster(0).AddItem (h)
    
            Meh(0) = .Meh
            GStufe(0) = GPMenge(Meh(0), .MengNeu)
        End With
    End If


    SQLStr$ = "SELECT * FROM ARTIKEL WHERE PZN = " + PznString(pzn)
    FabsErrf = Artikel.OpenRecordset(ArtikelAdoRec, SQLStr)
    If (FabsErrf% = 0) Then
    '    ast.GetRecord (FabsRecno& + 1)
    Else
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + PznString(pzn)
        'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
        On Error Resume Next
        TaxeRec.Close
        Err.Clear
        On Error GoTo DefErr
        TaxeRec.Open SQLStr, taxeAdoDB.ActiveConn
        If (TaxeRec.EOF = False) Then
            Call Taxe2ast(PznString(pzn))
            FabsErrf% = 0
        End If
    End If
        
    If (FabsErrf% = 0) Then
'                AuswahlHaPzn = AdvReader2(AdvReader2.GetOrdinal("Pzn"))

        With ast
            h = "HA" + vbTab + .pzn + vbTab + .kurz + vbTab + .meng + vbTab + .Meh + vbTab
            h = h + Format(.aep, "0.00") + vbTab
            flxTaxMuster(0).AddItem (h)
    
            Meh(1) = .Meh
            GStufe(1) = GPMenge(Meh(1), .MengNeu)
        End With
    End If

    For i = 0 To 1
        If (Meh(i) = "G") And (Meh((i + 1) Mod 2) = "MG") Then
            GStufe(i) = GStufe(i) * 1000#
            Meh(i) = "MG"
            Exit For
        End If
    Next

    If (SollTaxierTyp = MAG_GEFAESS) Then
'        .Meh = "ST"
'        sMenge = "1"
        GStufe(0) = 1
        Meh(0) = "ST"
        GStufe(1) = 1
        Meh(1) = "ST"
    End If

    h = h & vbCrLf + "Mengen:" + vbCrLf
    For i = 0 To 1
        h = h & IIf(i = 0, "FAM", "HA") + ":  " + CStr(GStufe(i)) + " " + Meh(i) + vbCrLf
        flxTaxMuster(0).TextMatrix(i + 1, 6) = CStr(GStufe(i))
        flxTaxMuster(0).TextMatrix(i + 1, 7) = Meh(i)
    Next

    If (Meh(0) <> Meh(1)) Then
        '                MsgBox("<" + Meh(0) + ">  <" + Meh(1) + ">")
        For i = 0 To 1
            If (Meh(i) = "ML") And (Meh((i + 1) Mod 2) = "G") Then
                GStufe(i) = GStufe(i) * dDichte
                Meh(i) = "G"
                '                                        h &= vbCrLf + vbCrLf + "Dichte: " + dDichte.ToString + "  GStufe: " + GStufe(i).ToString + "   Meh: " + Meh(i)
                Exit For
            End If
        Next
        
        h = h & vbCrLf + "Dichte: " + CStr(dDichte) + vbCrLf
        flxTaxMuster(0).TextMatrix(2, 8) = CStr(dDichte)
        For i = 0 To 1
            h = h & IIf(i = 0, "FAM", "HA") + ":  " + CStr(GStufe(i)) + " " + Meh(i) + vbCrLf
            flxTaxMuster(0).TextMatrix(i + 1, 9) = CStr(GStufe(i))
            flxTaxMuster(0).TextMatrix(i + 1, 10) = Meh(i)
        Next
    End If

    Dim Faktor, kp As Double
    Faktor = GStufe(0) / GStufe(1)
    kp = ast.aep * Faktor
    'MsgBox(AdvReader2(AdvReader2.GetOrdinal("ApothekenEk1")).ToString + " " + Faktor.ToString + " " + (AdvReader2(AdvReader2.GetOrdinal("ApothekenEk1")) * Faktor).ToString + " " + .kp.ToString)

    h = h & vbCrLf + "Mengen-Faktor:  " + Format(Faktor, "0.00")
    h = h & vbCrLf + "FAM-EK neu:  " + Format(kp, "0.00")
    flxTaxMuster(0).TextMatrix(1, 11) = Format(Faktor, "0.00")
    flxTaxMuster(0).TextMatrix(1, 12) = Format(kp, "0.00")
    flxTaxMuster(0).TextMatrix(1, 13) = "(" + Format(ast.aep, "0.00") + " * " + Format(Faktor, "0.00") + ")"
    
    TaxeRec.Close

'    MsgBox (h)

'            .AI = 1
End If

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

Private Sub nlcmdEsc_Click()
Call cmdEsc_Click
End Sub

'Private Sub nlcmdf5_Click()
'Call cmdF5_Click
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If (para.Newline) Then
    If (KeyAscii = 13) Then
'        Call nlcmdOk_Click
'        Exit Sub
    ElseIf (KeyAscii = 27) And (nlcmdEsc.Visible) Then
        Call nlcmdEsc_Click
        Exit Sub
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

