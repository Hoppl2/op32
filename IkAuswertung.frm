VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlButton.ocx"
Begin VB.Form frmIkAuswertung 
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
      Picture         =   "IkAuswertung.frx":0000
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
      Picture         =   "IkAuswertung.frx":00B9
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
      Picture         =   "IkAuswertung.frx":016D
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
   Begin MSFlexGridLib.MSFlexGrid flxIkAuswertung 
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
Attribute VB_Name = "frmIkAuswertung"
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

Private Const DefErrModul = "IKAUSWERTUNG.FRM"

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

Private Sub cmdF6_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdF6_Click")
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

Call IkAuswertungAusdruck

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
Dim i%, wi%, spBreite%
Dim iAdd%, iAdd2%
Dim c As Control

iEditModus = 1

Me.Caption = "Auswertung der gedruckten Rezepte nach Ik-Nummern"

Call wpara.InitFont(Me)

With flxIkAuswertung
    .Rows = 2
    .FixedRows = 1
    .FormatString = ">IK|<Name|>AnzRezepte|"
    .Rows = 1
    
    .ColWidth(0) = TextWidth(String(9, "9"))
    .ColWidth(1) = TextWidth(String(40, "A"))
    .ColWidth(2) = TextWidth(String(12, "9"))
    .ColWidth(3) = wpara.FrmScrollHeight

    .Top = wpara.TitelY
    .Left = wpara.LinksX
    .Height = .RowHeight(0) * 21 + 90

    wi% = 0
    For i% = 0 To (.Cols - 1)
        wi% = wi% + .ColWidth(i%)
    Next i%
    .Width = wi% + 90

    Call IkAuswertungBefuellen
    
    .row = 1
    .HighLight = flexHighlightNever
    .BackColor = Me.BackColor
End With


Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

Me.Width = flxIkAuswertung.Left + flxIkAuswertung.Width + 2 * wpara.LinksX

With cmdF6
    .Left = flxIkAuswertung.Left
    .Top = flxIkAuswertung.Top + flxIkAuswertung.Height + 150 * wpara.BildFaktor
    .Width = TextWidth(.Caption) + 150
    .Height = wpara.ButtonY
End With

With cmdEsc
    .Top = cmdF6.Top
    .Width = wpara.ButtonX%
    .Height = wpara.ButtonY%
    .Left = flxIkAuswertung.Left + flxIkAuswertung.Width - .Width
End With

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY% + 90 + wpara.FrmCaptionHeight

If (para.Newline) Then
    iAdd = wpara.NlFlexBackY
    iAdd2 = wpara.NlCaptionY
    
    With flxIkAuswertung
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
        .Top = flxIkAuswertung.Top + flxIkAuswertung.Height + iAdd + 600
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
    With flxIkAuswertung
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

Private Sub IkAuswertungBefuellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("IkAuswertungBefuellen")
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
Dim h$, SQLStr$

SQLStr$ = "SELECT * FROM AutIdemKrankenkassen ORDER BY AnzRezepte DESC"
Set AutIdemKkRec = RezSpeicherDB.OpenRecordset(SQLStr$)
Do
    If (AutIdemKkRec.EOF) Then
        Exit Do
    End If

    h$ = Format(AutIdemKkRec!IK, "0") + vbTab
    h$ = h$ + Trim(AutIdemKkRec!Name) + vbTab
    h$ = h$ + Format(AutIdemKkRec!AnzRezepte, "0") + vbTab
'    For i% = 0 To 8
        flxIkAuswertung.AddItem h$
'    Next i%

    AutIdemKkRec.MoveNext
Loop

With flxIkAuswertung
    If (.Rows = 1) Then
        .AddItem vbTab + "(leer)"
    End If
End With
    
Call DefErrPop
End Sub

Sub IkAuswertungAusdruck()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("IkAuswertungAusdruck")
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
Dim i%, j%, k%, ret%, rInd%, y%, sp%(5), anz%, ind%, AktLief%, Erst%, RowHe%
Dim ZentrierX%, DruckerWechsel%, MitArbInd%, gesBreite&, x%
Dim RetourWert#
Dim tx$, h$, AktDruckerName$, Titel$, h2$, AuftragsNr$, SQLStr$, Profil$

Call StartAnimation(frmAction, "Ausdruck wird erstellt ...")
    
With frmAction
    .picEan.Left = .picAnimationBack.Left + .picAnimationBack.Width
    .picEan.Top = .picAnimationBack.Top
    .picEan.Height = .picAnimationBack.Height
'    .picEan.Visible = True
End With

AnzDruckSpalten% = 2
ReDim DruckSpalte(AnzDruckSpalten% - 1)

With DruckSpalte(0)
    .Titel = "IK-Ean"
    .TypStr = String$(11, "X")
    .Ausrichtung = "L"
End With
With DruckSpalte(1)
    .Titel = "Krankenkasse"
    .TypStr = String$(24, "X")
    .Ausrichtung = "L"
End With
    

Printer.ScaleMode = vbTwips
Printer.Font.Name = "Arial"
Printer.Font.Size = 11  'nötig wegen Canon-BJ; sonst ab 2.Ausdruck falsch
Printer.Font.Size = 12
    
For i% = 0 To (AnzDruckSpalten% - 1)
    If (i% = 0) Then
        DruckSpalte(0).StartX = 60
    Else
        DruckSpalte(i%).StartX = DruckSpalte(i% - 1).StartX + DruckSpalte(i% - 1).BreiteX + Printer.TextWidth("  ")
    End If
    DruckSpalte(i%).BreiteX = Printer.TextWidth(RTrim(DruckSpalte(i%).TypStr))
Next i%
sp%(0) = DruckSpalte(1).StartX + DruckSpalte(1).BreiteX + 210
    
    
Printer.CurrentY = 150
With flxIkAuswertung
    For i% = 1 To (.Rows - 1)
        If (i% > 10) Then
            Exit For
        End If
        
        y% = Printer.CurrentY
        tx$ = "98600" + Format(Val(.TextMatrix(i%, 0)), "0000000") + "9"
        Call EanPruef(tx$)
        Call EanDruck(frmAction.picEan, 13, tx$)
        Printer.PaintPicture frmAction.picEan.Image, DruckSpalte(0).StartX, y%, DruckSpalte(0).BreiteX, 450, 0, 0
        
        Printer.CurrentX = DruckSpalte(1).StartX
        Printer.CurrentY = y% + 150
        
        tx$ = .TextMatrix(i%, 1)
        For j% = 12 To 8 Step -1
            Printer.Font.Size = j%
            If (Printer.TextWidth(tx$) < DruckSpalte(1).BreiteX) Then
                Exit For
            End If
        Next j%
        
        Do
            If (Printer.TextWidth(tx$) < DruckSpalte(1).BreiteX) Then
                Exit Do
            End If
            If (tx$ = "") Then
                Exit Do
            End If
            tx$ = Left$(tx$, Len(tx$) - 1)
        Loop
    
        Printer.Print tx$;
        
        Printer.CurrentY = y% + 690
    Next i%

    y% = Printer.CurrentY + 450
    tx$ = "9860000000001"
    Call EanDruck(frmAction.picEan, 13, tx$)
    Printer.PaintPicture frmAction.picEan.Image, DruckSpalte(0).StartX, y%, DruckSpalte(0).BreiteX, 450, 0, 0
        
    Printer.CurrentX = DruckSpalte(1).StartX
    Printer.CurrentY = y% + 150
    
    tx$ = "IK-Eingabe"
    Printer.Font.Size = 11
    Printer.Print tx$;
    
    .Redraw = False
    .row = .FixedRows
    .col = 1
    .RowSel = .Rows - 1
    .ColSel = .col
    .Sort = 5
    .Redraw = True
    
    AnzDruckSpalten% = 2
    ReDim DruckSpalte(AnzDruckSpalten% - 1)
    
    With DruckSpalte(0)
        .Titel = "IK"
        .TypStr = String$(8, "9")
        .Ausrichtung = "L"
    End With
    With DruckSpalte(1)
        .Titel = "Krankenkasse"
        .TypStr = String$(28, "X")
        .Ausrichtung = "L"
    End With
        
    
    Printer.ScaleMode = vbTwips
    Printer.Font.Name = "Arial"
    Printer.Font.Size = 12  'nötig wegen Canon-BJ; sonst ab 2.Ausdruck falsch
    Printer.Font.Size = 11
        
    For i% = 0 To (AnzDruckSpalten% - 1)
        If (i% = 0) Then
            DruckSpalte(0).StartX = sp%(0)
        Else
            DruckSpalte(i%).StartX = DruckSpalte(i% - 1).StartX + DruckSpalte(i% - 1).BreiteX + Printer.TextWidth("  ")
        End If
        DruckSpalte(i%).BreiteX = Printer.TextWidth(RTrim(DruckSpalte(i%).TypStr))
    Next i%
    
    RowHe% = Printer.TextHeight("Äg")
            
    Printer.CurrentY = 150
    For i% = 1 To (.Rows - 1)
        tx$ = Format(Val(.TextMatrix(i%, 0)), "0000000")
        Printer.CurrentX = DruckSpalte(0).StartX
        Printer.Print tx$;
        
        tx$ = .TextMatrix(i%, 1)
        For j% = 11 To 8 Step -1
            Printer.Font.Size = j%
            If (Printer.TextWidth(tx$) < DruckSpalte(1).BreiteX) Then
                Exit For
            End If
        Next j%
        
        Do
            If (Printer.TextWidth(tx$) < DruckSpalte(1).BreiteX) Then
                Exit Do
            End If
            If (tx$ = "") Then
                Exit Do
            End If
            tx$ = Left$(tx$, Len(tx$) - 1)
        Loop
    
        Printer.CurrentX = DruckSpalte(1).StartX
        Printer.CurrentY = Printer.CurrentY + (RowHe% - Printer.TextHeight(tx$)) / 2
        Printer.Print tx$;
        
        Printer.Font.Size = 11
        Printer.Print
    
        If (Printer.CurrentY > Printer.ScaleHeight - 1000) Then
            Exit For
        End If
    Next i%
End With

Call HoleDruckRänder

Printer.DrawStyle = vbDot
x% = twpX(PaperInMM.SizeHorz - PaperInMM.PrintRight)
y% = twpY(PaperInMM.SizeVert / 2 - PaperInMM.PrintTop)
Printer.Line (0, y%)-(x%, y%), vbButtonFace, B
x% = twpX(PaperInMM.SizeHorz / 2 - PaperInMM.PrintLeft)
y% = twpY(PaperInMM.SizeVert - PaperInMM.PrintBottom)
Printer.Line (x%, 0)-(x%, y%), vbButtonFace, B
Printer.DrawStyle = vbSolid

Printer.EndDoc

frmAction.picEan.Visible = False

Call StopAnimation(frmAction)

Call DefErrPop
End Sub

'Sub InitDruckZeile(Optional ZentrierX% = False)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("InitDruckZeile")
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
'Dim i%, j%, zentr%
'Dim gesBreite&
'Dim h$
'
'DruckSpalte(0).StartX = 0
'DruckSpalte(0).StartX = MachLinkenRand(20)
'
'For j% = 12 To 5 Step -1
'    Printer.ScaleMode = vbTwips
'    Printer.Font.Name = "Arial"
'    Printer.Font.Size = 18  'nötig wegen Canon-BJ; sonst ab 2.Ausdruck falsch
'    Printer.Font.Size = j%
'
'    For i% = 0 To (AnzDruckSpalten% - 1)
'        If (i% > 0) Then
'            DruckSpalte(i%).StartX = DruckSpalte(i% - 1).StartX + DruckSpalte(i% - 1).BreiteX + Printer.TextWidth("  ")
'        End If
'        DruckSpalte(i%).BreiteX = Printer.TextWidth(RTrim$(DruckSpalte(i%).TypStr))
'    Next i%
'
'    gesBreite& = DruckSpalte(AnzDruckSpalten% - 1).StartX + DruckSpalte(AnzDruckSpalten% - 1).BreiteX
'    If (gesBreite& < Printer.ScaleWidth) Then Exit For
'Next j%
'
'DruckFontSize% = j%
'
'If (ZentrierX%) Then
'    zentr% = (Printer.ScaleWidth - gesBreite&) / 2
'    For i% = 0 To (AnzDruckSpalten% - 1)
'        DruckSpalte(i%).StartX = DruckSpalte(i%).StartX + zentr%
'    Next i%
'End If
'
'Call DefErrPop
'End Sub

'Function MachLinkenRand%(SollRandInMM%)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("MachLinkenRand%")
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
'Dim ret%
'Dim LRandInPixel&, ORandInPixel&
'Dim PrTwipsPerPixelX!, PrTwipsPerpixelY!, LRandInMM!, ORandInMM!
'
''Randbreiten in Pixel
'LRandInPixel& = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX)
'ORandInPixel& = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY)
'
''Auflösung Drucker
'PrTwipsPerPixelX = Printer.TwipsPerPixelX
'PrTwipsPerpixelY = Printer.TwipsPerPixelY
'
''Randbreiten in Millimeter
'LRandInMM! = Printer.ScaleX(LRandInPixel& * PrTwipsPerPixelX!, vbTwips, vbMillimeters)
'ORandInMM! = Printer.ScaleY(ORandInPixel& * PrTwipsPerpixelY!, vbTwips, vbMillimeters)
'
'ret% = 0
'If (SollRandInMM% > LRandInMM!) Then
'    ret% = Printer.ScaleX(SollRandInMM% - LRandInMM!, vbMillimeters, vbTwips)
'End If
'
'MachLinkenRand% = ret%
'
'Call DefErrPop
'End Function

Sub HoleDruckRänder()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HoleDruckRänder")
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
Dim ret%
Dim x&, y&
Dim x1#, y1#
Dim LRandInPixel&, ORandInPixel&
Dim PrPapierX&, PrPapierY&
Dim PrBereichX&, PrBereichY&
Dim PrTwipsPerPixelX!, PrTwipsPerpixelY!, LRandInTwips!, ORandInTwips!, LRandInMM!, ORandInMM!

'Papiergrösse
'in Pixel
x1# = GetDeviceCaps(Printer.hdc, PHYSICALWIDTH)
y1# = GetDeviceCaps(Printer.hdc, PHYSICALHEIGHT)
'in MM
PaperInMM.SizeHorz = CInt(PrinterScaleX(x1#, vbPixels, vbMillimeters))
PaperInMM.SizeVert = CInt(PrinterScaleY(y1#, vbPixels, vbMillimeters))

'Randbreiten links/oben
'in Pixel
x1# = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX)
y1# = GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY)
'in MM
PaperInMM.PrintLeft = CInt(PrinterScaleX(x1#, vbPixels, vbMillimeters))
PaperInMM.PrintTop = CInt(PrinterScaleY(y1#, vbPixels, vbMillimeters))

'Druckbereich
'in Pixel
x& = GetDeviceCaps(Printer.hdc, HORZRES)
y& = GetDeviceCaps(Printer.hdc, VERTRES)
'in MM
x1# = PrinterScaleX(x1#, vbPixels, vbMillimeters)
y1# = PrinterScaleY(y1#, vbPixels, vbMillimeters)

'Randbreiten rechts/unten
'in MM
PaperInMM.PrintRight = CInt(PaperInMM.PrintLeft + x1#)
PaperInMM.PrintBottom = CInt(PaperInMM.PrintTop + y1#)

'Auflösung Drucker
PrTwipsPerPixelX = Printer.TwipsPerPixelX
PrTwipsPerpixelY = Printer.TwipsPerPixelY

''Randbreiten in Twips
'LRandInTwips! = LRandInPixel& * PrTwipsPerPixelX!
'ORandInTwips! = ORandInPixel& * PrTwipsPerpixelY!
'
''Randbreiten in Millimeter
'LRandInMM! = Printer.ScaleX(LRandInPixel& * PrTwipsPerPixelX!, vbTwips, vbMillimeters)
'ORandInMM! = Printer.ScaleY(ORandInPixel& * PrTwipsPerpixelY!, vbTwips, vbMillimeters)

'LRandPhysisch! = LRandInMM! / Ratio!
'ORandPhysisch! = ORandInMM! / Ratio!

Call DefErrPop
End Sub


Function twpX&(mm&, Optional iModus% = 0)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("twpX&")
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
Dim dMm#

dMm# = mm&
    
twpX& = PrinterScaleX(dMm#, vbMillimeters, vbTwips)

Call DefErrPop
End Function

Function twpY&(mm&)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("twpY&")
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
Dim dMm#

dMm# = mm&
    
twpY& = PrinterScaleY(dMm#, vbMillimeters, vbTwips)

Call DefErrPop
End Function


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

Private Sub nlcmdF6_Click()
Call cmdF6_Click
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







