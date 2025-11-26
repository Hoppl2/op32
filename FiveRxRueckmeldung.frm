VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlbutton.ocx"
Begin VB.Form frmFiveRxRueckmeldung 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4275
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4275
   Begin nlCommandButton.nlCommand nlcmdF6 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   3840
      Picture         =   "FiveRxRueckmeldung.frx":0000
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
      Index           =   1
      Left            =   3600
      Picture         =   "FiveRxRueckmeldung.frx":00B9
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
      Index           =   0
      Left            =   3360
      Picture         =   "FiveRxRueckmeldung.frx":016D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdF6 
      Caption         =   "&Drucken (F6)"
      Height          =   450
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   1680
      TabIndex        =   2
      Top             =   1920
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxFiveRxRueckmeldungen 
      Height          =   1320
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2328
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      ScrollBars      =   2
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
End
Attribute VB_Name = "frmFiveRxRueckmeldung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type PAPER_STRUCT
    SizeHorz As Double
    SizeVert As Double
    PrintLeft As Double
    PrintTop As Double
    PrintRight As Double
    PrintBottom As Double
End Type
Dim PaperInMM As PAPER_STRUCT

Dim iEditModus%

Private Const DefErrModul = "FIVERXRUECKMELDUNG.FRM"

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

'Private Sub cmdF6_Click()
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("cmdF6_Click")
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
'
'Call IkAuswertungAusdruck
'
'Call DefErrPop
'End Sub

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
Dim i%, wi%, spBreite%
Dim iAdd%, iAdd2%
Dim c As Control

iEditModus = 1

Me.Caption = "FiveRx-Rückmeldungen"

Call wpara.InitFont(Me)

With flxFiveRxRueckmeldungen
    .Rows = 2
    .FixedRows = 1
    .FormatString = "^HF|^Typ|<Fehler|<Kurz / Lang"
    .Rows = 1
    
    .ColWidth(0) = TextWidth(String(3, "9"))
    .ColWidth(1) = TextWidth(String(10, "A"))
    .ColWidth(2) = TextWidth(String(10, "A"))
    .ColWidth(3) = TextWidth(String(100, "A"))

    .Top = wpara.TitelY
    .Left = wpara.LinksX
    .Height = .RowHeight(0) * 21 + 90

    wi% = 0
    For i% = 0 To (.Cols - 1)
        wi% = wi% + .ColWidth(i%)
    Next i%
    .Width = wi% + 90
    
    .ColWidth(3) = TextWidth(String(100, "A"))  ' .ColWidth(3) * 3
    .ScrollBars = flexScrollBarBoth

    Call FiveRxRueckmeldungenBefuellen
    
    .row = 1
    .HighLight = flexHighlightNever
    .BackColor = Me.BackColor
End With


Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

Me.Width = flxFiveRxRueckmeldungen.Left + flxFiveRxRueckmeldungen.Width + 2 * wpara.LinksX

With cmdF6
    .Left = flxFiveRxRueckmeldungen.Left
    .Top = flxFiveRxRueckmeldungen.Top + flxFiveRxRueckmeldungen.Height + 150 * wpara.BildFaktor
    .Width = TextWidth(.Caption) + 150
    .Height = wpara.ButtonY
End With

With cmdEsc
    .Top = cmdF6.Top
    .Width = wpara.ButtonX%
    .Height = wpara.ButtonY%
    .Left = flxFiveRxRueckmeldungen.Left + flxFiveRxRueckmeldungen.Width - .Width
End With

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY% + 90 + wpara.FrmCaptionHeight

If (para.Newline) Then
    iAdd = wpara.NlFlexBackY
    iAdd2 = wpara.NlCaptionY
    
    With flxFiveRxRueckmeldungen
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
    End With
    
    cmdEsc.Top = cmdEsc.Top + 2 * iAdd
    
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
    
    With nlcmdEsc
        .Init
'        .Left = Me.ScaleWidth - .Width - 150
        .Top = flxFiveRxRueckmeldungen.Top + flxFiveRxRueckmeldungen.Height + iAdd + 600
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .Default = cmdEsc.Default
        .Cancel = cmdEsc.Cancel
        .Visible = True
    End With
    cmdEsc.Visible = False

    With nlcmdF6
        .Init
        .Left = cmdF6.Left
        .Top = nlcmdEsc.Top
        .Caption = cmdF6.Caption
        .TabIndex = cmdF6.TabIndex
        .Enabled = cmdF6.Enabled
        .Default = cmdF6.Default
        .Cancel = cmdF6.Cancel
        .Visible = True
    End With
    cmdF6.Visible = False

    nlcmdEsc.Left = (Me.Width - nlcmdEsc.Width) / 2

    Me.Height = nlcmdEsc.Top + nlcmdEsc.Height + wpara.FrmCaptionHeight + 450

    Call wpara.NewLineWindow(Me, nlcmdEsc.Top)
    With flxFiveRxRueckmeldungen
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

'                If (c.Visible) Then
                    With c.Container
                        .ForeColor = RGB(180, 180, 180) ' vbWhite
                        .FillStyle = vbSolid
                        .FillColor = c.BackColor

                        RoundRect .hdc, (c.Left - 60) / Screen.TwipsPerPixelX, (c.Top - 30) / Screen.TwipsPerPixelY, (c.Left + c.Width + 60) / Screen.TwipsPerPixelX, (c.Top + c.Height + 15) / Screen.TwipsPerPixelY, 10, 10
                    End With
'                End If
'            ElseIf (TypeOf c Is CheckBox) Then
'                c.Height = 0
'                c.Width = c.Height
'                If (c.Name = "chkHistorie") Then
'                    If (c.Index > 0) Then
'                        Load lblchkHistorie(c.Index)
'                    End If
'                    With lblchkHistorie(c.Index)
'                        .BackStyle = 0 'duchsichtig
'                        .Caption = c.Caption
'                        .Left = c.Left + 300
'                        .Top = c.Top
'                        .Width = TextWidth(.Caption) + 90
'                        .TabIndex = c.TabIndex
'                        .Visible = True
'                    End With
'                End If
            End If
        End If
    Next
    On Error GoTo DefErr
    
Else
    nlcmdEsc.Visible = False
    nlcmdF6.Visible = False
End If
'''''''''

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

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

If (para.Newline) Then
    If (KeyCode = vbKeyF6) Then
        nlcmdF6.Value = True
    End If
Else
    If (KeyCode = vbKeyF6) Then
        cmdF6.Value = True
    End If
End If

Call DefErrPop
End Sub

Private Sub FiveRxRueckmeldungenBefuellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FiveRxRueckmeldungenBefuellen")
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
Dim h$, h2$, SQLStr$


Dim fWert%, fCode%, fTCode%, PosNr%, fHauptFehler%, fKurzText$, fLangText$
Dim nrow&, prtw&
Dim eRezeptId$
Dim tagTransaktionsId$, tagStatus$, tagStatus2$, tagId$, StatusInfo$, fStatus$, fKommentar$
Dim valTransaktionsId$, valStatus$, valStatus2$, valId$, ValVerbesserung$, sAnzeige$

'MsgBox (XmlResponse)

tagStatus = "EREZEPTSTATUS"
tagStatus2 = "STATUS"
tagTransaktionsId = "EREZEPTID"

Do
    valStatus = XmlAbschnitt(XmlResponse, tagStatus)
    If (valStatus = "") Then
        tagStatus = "EREZEPTVORPRUEFUNGSTATUS"
        tagStatus2 = "VSTATUS"
        tagTransaktionsId = "EREZEPTID"
        valStatus = XmlAbschnitt(XmlResponse, tagStatus)
    End If
    If (valStatus = "") Then
        Exit Do
    End If

    h2$ = XmlAbschnitt(valStatus, tagTransaktionsId)
    eRezeptId = h2
    If (eRezeptId <> "") Then
'        AbrechnungsStatus = 1
'        AbrechnungsStatusStr$ = ""
'        AbrechnungsStatusLangStr$ = ""

        valStatus2 = XmlAbschnitt(valStatus, tagStatus2)
        Me.Caption = "FiveRx-Rückmeldungen: " + valStatus2

'        AbrechnungsStatusStr$ = valStatus2
'        For i = 0 To UBound(FiveRxRezeptStatus)
'            If (AbrechnungsStatusStr = FiveRxRezeptStatus(i)) Then
'                AbrechnungsStatus = i + 1
'                Exit For
'            End If
'        Next i

        For i = 0 To 10
            StatusInfo = XmlAbschnitt(valStatus, "STATUSINFO")
            If (StatusInfo <> "") Then
                fCode = Val(XmlAbschnitt(StatusInfo, "FCODE"))
                fStatus = XmlAbschnitt(StatusInfo, "FSTATUS")
                fKommentar = XmlAbschnitt(StatusInfo, "FKOMMENTAR")
                fKommentar = Left(fKommentar, 100)

                fWert = xVal(XmlAbschnitt(StatusInfo, "FWERT"))
                fTCode = Val(XmlAbschnitt(StatusInfo, "FTCODE"))
                PosNr = Val(XmlAbschnitt(StatusInfo, "POSNR"))
                fKurzText = XmlAbschnitt(StatusInfo, "FKURZTEXT")
                fLangText = XmlAbschnitt(StatusInfo, "FLANGTEXT")
                fHauptFehler = Abs(UCase(XmlAbschnitt(StatusInfo, "FHAUPTFEHLER")) = "TRUE")
                
'                If (AbrechnungsStatusLangStr$ = "") Then
'                    AbrechnungsStatusLangStr$ = fKommentar
'                End If
                
                h = IIf(fHauptFehler, Chr(214), "") + vbTab
                h = h + fStatus + vbTab
                h = h + Str(fCode) + " / " + Str(fTCode) + vbTab
                h = h + fKurzText + vbCrLf + fLangText + vbTab
                With flxFiveRxRueckmeldungen
                    .AddItem h
'                    prtw = Me.TextWidth(.TextMatrix(.Rows - 1, 3))
'                    nrow = Int(prtw / .ColWidth(3)) + 3
'                    .RowHeight(.Rows - 1) = nrow * .RowHeight(0)
                    .RowHeight(.Rows - 1) = Me.TextHeight(fKurzText + vbCrLf + fLangText + vbCrLf) + 150
                    If (UCase(fStatus) = "FEHLER") Then
                        .FillStyle = flexFillRepeat
                        .row = .Rows - 1
                        .col = 0
                        .RowSel = .row
                        .ColSel = .Cols - 1
                        .CellFontBold = True
                        .FillStyle = flexFillSingle
                    End If
                End With
            Else
                Exit For
            End If
        Next i
'        If (AbrechnungsStatusLangStr$ = "") Then
'            AbrechnungsStatusLangStr$ = FiveRxRezeptStatus(AbrechnungsStatus - 1)
'        End If
    End If
Loop
With flxFiveRxRueckmeldungen
    If (.Rows = 1) Then
        .AddItem vbTab + "(leer)"
    End If
    
    .FillStyle = flexFillRepeat
    .row = 1
    .col = 0
    .RowSel = .Rows - 1
    .ColSel = .col
    .CellFontName = "Symbol"
    .FillStyle = flexFillSingle
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

'Private Sub nlcmdF6_Click()
'Call cmdF6_Click
'End Sub

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







