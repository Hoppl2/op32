VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRueckmeldungen 
   Caption         =   "Rueckmeldungen"
   ClientHeight    =   4245
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   4305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   4305
   Begin VB.CommandButton cmdF6 
      Caption         =   "Drucken (F6)"
      Height          =   450
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   1200
   End
   Begin VB.Frame fmeDefekte 
      Caption         =   "Defekte"
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4095
      Begin MSFlexGridLib.MSFlexGrid flxRueck 
         Height          =   1080
         Left            =   840
         TabIndex        =   0
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1905
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
         ScrollBars      =   2
         SelectionMode   =   1
      End
   End
   Begin VB.Frame fmeBestaetigung 
      Caption         =   "Annahmebestätigung"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   3855
      Begin VB.Label lblBestaetigung 
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   3480
      Width           =   855
   End
End
Attribute VB_Name = "frmRueckmeldungen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "RUECKMELDUNGEN.FRM"

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
Call DruckeRueckmeldungen
Unload Me
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

If (KeyCode = vbKeyF6) Then
    cmdF6.Value = True
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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%
Dim h$, h2$, FormStr$

Call wpara.InitFont(Me)


Font.Bold = False   ' True

With flxRueck
    .Cols = 3
    .Rows = 2
    .FixedRows = 1
    .FixedCols = 0

    .Top = 2 * wpara.TitelY
    .Left = wpara.LinksX
    .Height = .RowHeight(0) * 11 + 90
    
    .FormatString = "<Artikel|>Eh|^Pzn|>BM|>FehlM|<Defektgrund"
    .Rows = 1
    .SelectionMode = flexSelectionByRow
    
    Call RueckmeldungenBefuellen
    If (.Rows = 1) Then
        .AddItem "keine Defekte empfangen !"
    End If
    
    MaxWi% = TextWidth(String(20, "W"))
    For i% = 1 To .Rows - 1
        h$ = .TextMatrix(i%, 0)
        If (TextWidth(h$) > MaxWi%) Then
            MaxWi% = TextWidth(h$)
        End If
    Next i%
    
    .ColWidth(0) = MaxWi%
    .ColWidth(1) = TextWidth("1000 ST")
    .ColWidth(2) = TextWidth("99999999")
    .ColWidth(3) = TextWidth("99999")
    .ColWidth(4) = TextWidth("99999")
    .ColWidth(5) = TextWidth(String(15, "W"))

    spBreite% = 0
    For i% = 0 To .Cols - 1
        .ColWidth(i%) = .ColWidth(i%) + TextWidth("X")
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
    
    If (.Width > (frmAction.Width - 600)) Then
        .Width = frmAction.Width - 600
    End If
End With


With fmeDefekte
    .Width = flxRueck.Width + 2 * wpara.LinksX
    .Height = flxRueck.Top + flxRueck.Height + wpara.TitelY%
    .Left = wpara.LinksX
    .Top = wpara.TitelY
End With


Me.Width = fmeDefekte.Width + 2 * wpara.LinksX


With lblBestaetigung
    .Width = flxRueck.Width
    .Height = TextHeight(Bestaetigung$) + 90
    .Left = wpara.LinksX%
    .Top = 2 * wpara.TitelY%
    .Caption = Bestaetigung$
End With


With fmeBestaetigung
    .Width = fmeDefekte.Width
    .Height = lblBestaetigung.Top + lblBestaetigung.Height + wpara.TitelY%
    .Left = fmeDefekte.Left
    .Top = fmeDefekte.Top + fmeDefekte.Height + 150
End With

With cmdOk
    .Width = wpara.ButtonX
    .Height = wpara.ButtonY
    .Left = (Me.ScaleWidth - .Width) / 2
    .Top = fmeBestaetigung.Top + fmeBestaetigung.Height + 150
End With

With cmdF6
    .Width = TextWidth(.Caption) + 150
    .Height = cmdOk.Height
    .Left = fmeDefekte.Left
    .Top = cmdOk.Top
End With

Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

Call DefErrPop
End Sub

Sub RueckmeldungenBefuellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("RueckmeldungenBefuellen")
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
Dim i%, ind%, SatzTyp%
Dim h$, h2$
        
Bestaetigung$ = ""

With frmRueckmeldungen.flxRueck
    If (MaxSendSatz% > 0) Then
        For i% = 1 To MaxSendSatz%
            h$ = Trim$(SendSatz$(i%))
            
            SatzTyp% = Val(Left$(h$, 2))
            
            If (para.Land = "A") Then
                If (SatzTyp% = 3) Then
                
                    h2$ = Mid$(h$, 15, 35) + vbTab + vbTab + Mid$(h$, 50, 7) + vbTab
                    h2$ = h2$ + Str$(Val(Mid$(h$, 5, 4))) + vbTab + Str$(Val(Mid$(h$, 9, 4))) + vbTab
                    h2$ = h2$ + Mid$(h$, 13, 2)
                    .AddItem h2$
                    .row = .Rows - 1
                    .FillStyle = flexFillRepeat
                    .col = 0
                    .ColSel = .Cols - 1
                    .CellFontBold = True
                    .FillStyle = flexFillSingle
                End If
            Else
                If (SatzTyp% = 4) Then
                    h2$ = Mid$(h$, 44, 26) + vbTab + Mid$(h$, 35, 9) + vbTab + Mid$(h$, 22, 7) + vbTab
                    h2$ = h2$ + Str$(Val(Mid$(h$, 18, 4))) + vbTab + Str$(Val(Mid$(h$, 31, 4))) + vbTab
                    h2$ = h2$ + Mid$(h$, 70, 15)
                    .AddItem h2$
                    .row = .Rows - 1
                    .FillStyle = flexFillRepeat
                    .col = 0
                    .ColSel = .Cols - 1
                    .CellFontBold = True
                    .FillStyle = flexFillSingle
                ElseIf (SatzTyp% = 6) Or (SatzTyp% = 8) Then
                    If (SatzTyp% = 6) Then
                        h$ = Mid$(h$, 18)
                    Else
                        h$ = Mid$(h$, 13)
                    End If
                    
                    Do
                        ind% = InStr(h$, "<<")
                        If (ind% > 0) Then
                            h$ = Left$(h$, ind% - 1) + Mid$(h$, ind% + 1)
                        Else
                            Exit Do
                        End If
                    Loop
                    
                    Do
                        ind% = InStr(h$, ">")
                        If (ind% > 0) Then
                            h$ = Left$(h$, ind% - 1) + Mid$(h$, ind% + 1)
                        Else
                            Exit Do
                        End If
                    Loop
                    
                    Do
                        If (Left$(h$, 1) = "<") Then
                            h$ = Mid$(h$, 2)
                        Else
                            h$ = Trim(h$)
                            Exit Do
                        End If
                    Loop
                    
                    Do
                        If (h$ = "") Then Exit Do
                        ind% = InStr(h$, "<")
                        If (ind% > 0) Then
                            h2$ = Left$(h$, ind% - 1)
                            h$ = Mid$(h$, ind% + 1)
                        Else
                            h2$ = h$
                            h$ = ""
                        End If
                        If (SatzTyp% = 6) Then
                            .AddItem "   " + h2$
                        Else
                            Bestaetigung$ = Bestaetigung$ + h2$ + vbCrLf
                        End If
                    Loop
                End If
                
    '            .AddItem Left$(h$, 2) + vbTab + Mid$(h$, 3, 3) + vbTab + Mid$(h$, 6)
            End If
    
        Next i%
    Else
        Bestaetigung$ = "keine Rückmeldungen empfangen !"
    End If
End With

'With frmRueckmeldungen.flxRueck
'    If (MaxSendSatz% > 1) Then
'        For i% = 1 To MaxSendSatz%
'            h$ = Trim$(SendSatz$(i%))
'            .AddItem Left$(h$, 2) + vbTab + Mid$(h$, 3, 3) + vbTab + Mid$(h$, 6)
'        Next i%
'    Else
'        .AddItem "" + vbTab + "" + vbTab + "keine Rückmeldungen empfangen !"
'    End If
'End With

Call DefErrPop
End Sub


Sub DruckeRueckmeldungen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DruckeRueckmeldungen")
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
Dim i%, j%, pos%, sp%(9), SN%, Y%, Max%, ind%, iLief%, iRufzeit%
Dim header$, tx$, h$, KopfZeile$

Call StartAnimation(frmAction, "Ausdruck wird erstellt ...")

frmAction.lstSortierung.Clear

KopfZeile$ = "Rückmeldungen"
header$ = "(unbekannt)"
iLief% = Val(Left$(GesendetDatei$, 3))
If (iLief% > 0) And (iLief% <= lif.AnzRec) Then
    lif.GetRecord (iLief% + 1)
    h$ = RTrim$(lif.Name(0))
    
    iRufzeit% = Val(Mid$(GesendetDatei$, 4, 4))
    h$ = h$ + "  (" + Format(iRufzeit% \ 100, "00") + ":" + Format(iRufzeit% Mod 100, "00") + ")"
    
    If (InStr(GesendetDatei$, "m.") > 0) Then h$ = h$ + "  manuell"
    header$ = h$
End If
        

With frmRueckmeldungen.flxRueck
    For i% = 1 To (.Rows - 1)
        h$ = Format(i%, "000")
        For j% = 0 To (.Cols - 1)
            h$ = h$ + .TextMatrix(i%, j%) + vbTab
        Next j%
        frmAction.lstSortierung.AddItem h$
    Next i%
End With


Printer.ScaleMode = vbTwips
Printer.Font.Name = "Arial"

DruckSeite% = 0
    
Call DruckKopf(header$, "", KopfZeile$)
        
sp%(0) = Printer.TextWidth(String$(28, "X"))
sp%(1) = sp%(0) + Printer.TextWidth("1000 ST")
sp%(2) = sp%(1) + Printer.TextWidth(String$(8, "9"))
sp%(3) = sp%(2) + Printer.TextWidth(String$(5, "9"))
sp%(4) = sp%(3) + Printer.TextWidth(String$(6, "9"))
sp%(5) = sp%(4) + Printer.TextWidth(String$(15, "W"))

Printer.CurrentX = 0
Printer.Print "A R T I K E L";
Printer.CurrentX = sp%(0)
Printer.Print "Einheit";
Printer.CurrentX = sp%(1)
Printer.Print "P Z N";
tx$ = "B M"
Printer.CurrentX = sp%(3) - Printer.TextWidth(tx$)
Printer.Print tx$;
tx$ = "FehlM"
Printer.CurrentX = sp%(4) - Printer.TextWidth(tx$)
Printer.Print tx$;
Printer.CurrentX = sp%(4) + Printer.TextWidth("x")
Printer.Print "DefGrund";

Printer.Print " "
Y% = Printer.CurrentY
Printer.Line (0, Y%)-(sp%(5), Y%)
                
    
With frmAction.lstSortierung
    For i% = 1 To .ListCount
        .ListIndex = i% - 1
        h$ = Mid$(.text, 4)
        
        If (Left$(h$, 3) = "   ") Then
            Printer.FontSize = 10
        Else
            Printer.FontSize = 12
        End If
                
        For j% = 0 To 5
            ind% = InStr(h$, vbTab)
            tx$ = Left$(h$, ind% - 1)
            h$ = Mid$(h$, ind% + 1)
            
            If (j% = 0) Then
                Printer.CurrentX = 0
            ElseIf (j% = 2) Then
                Printer.CurrentX = sp%(j% - 1)
            ElseIf (j% = 5) Then
                Printer.CurrentX = sp%(j% - 1) + Printer.TextWidth("x")
            Else
                tx$ = Trim(tx$)
                Printer.CurrentX = sp%(j%) - Printer.TextWidth(tx$ + "x")
            End If
            
            Printer.Print tx$;
        Next j%
        
        Printer.Print " "
        
        If (Printer.CurrentY > Printer.ScaleHeight - 1000) Then
            Call DruckFuss
            Call DruckKopf(header$, "", KopfZeile$)
        End If
    Next i%
End With
    
Printer.Print " "
Printer.FontSize = 10
Printer.Print frmRueckmeldungen.lblBestaetigung.Caption

Call DruckFuss(False)

Printer.EndDoc

Call StopAnimation(frmAction)

Call DefErrPop
End Sub



