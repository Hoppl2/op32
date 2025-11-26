VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGesendet 
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4305
   Begin VB.CommandButton cmdF6 
      Caption         =   "Drucken (F6)"
      Height          =   450
      Left            =   2880
      TabIndex        =   2
      Top             =   3360
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Default         =   -1  'True
      Height          =   450
      Left            =   360
      TabIndex        =   3
      Top             =   3000
      Width           =   1200
   End
   Begin VB.CommandButton cmdF2 
      Caption         =   "Rückmeldungen (F2)"
      Height          =   450
      Left            =   3000
      TabIndex        =   1
      Top             =   2760
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxGesendet 
      Height          =   2280
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
      ScrollBars      =   2
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmGesendet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "GESENDET.FRM"

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
frmRueckmeldungen.Show 1
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
Call DruckeBestellung(GesendetDatei$)
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

If (KeyCode = vbKeyF2) Then
    cmdF2.Value = True
ElseIf (KeyCode = vbKeyF6) Then
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
Dim i%, spBreite%, ind%, iLief%, iRufzeit%
Dim DRUCKHANDLE%
Dim h$, h2$

Call wpara.InitFont(Me)

With flxGesendet
    .Cols = 8
    .Rows = 2
    .FixedRows = 1
    .FormatString = "<PZN|<Name|>Menge|^Meh|>BM|>NR|>Zeilenwert|"
        
    Font.Bold = True
    .ColWidth(0) = TextWidth("9999999")
    .ColWidth(1) = TextWidth("Xxxxxx Xxxxxx Xxxxxx Xxxxxx")
    .ColWidth(2) = TextWidth("XXXXXX")
    .ColWidth(3) = TextWidth("XXX")
    .ColWidth(4) = TextWidth("9999")
    .ColWidth(5) = TextWidth("9999")
    .ColWidth(6) = TextWidth("999999.99")
    .ColWidth(7) = wpara.FrmScrollHeight
    Font.Bold = False
    
    spBreite% = 0
    For i% = 0 To .Cols - 1
        If (.ColWidth(i%) > 0) Then
            .ColWidth(i%) = .ColWidth(i%) + TextWidth("X")
        End If
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
    
    .Rows = 1
    MaxSendSatz% = 0
    
    DRUCKHANDLE% = FileOpen("winw\" + GesendetDatei$, "I")
    Do While Not EOF(DRUCKHANDLE%)
        Line Input #DRUCKHANDLE%, h$
        If (Left$(h$, 4) = "RM: ") Then
            MaxSendSatz% = MaxSendSatz% + 1
            ReDim Preserve SendSatz$(MaxSendSatz%)
            SendSatz$(MaxSendSatz%) = Mid$(h$, 5)
        Else
            .Rows = .Rows + 1
            .row = .Rows - 1
            For i% = 0 To 5
                ind% = InStr(h$, vbTab)
                h2$ = Left$(h$, ind% - 1)
                h$ = Mid$(h$, ind% + 1)
                .TextMatrix(.row, i% + 1) = h2$
            Next i%
            .TextMatrix(.row, 0) = Left$(h$, Len(h$) - 1)
        End If
    Loop
    Close #DRUCKHANDLE%
    
    If (.Rows = 1) Then .Rows = 2

    .row = 1
    .col = 1
    .RowSel = .Rows - 1 ' AnzBestellArtikel%
    .ColSel = 3
    .Sort = 5

'    Call ActProgram.PreisKalkBefuellen

    .Top = wpara.TitelY
    .Left = wpara.LinksX
    .Height = 11 * .RowHeight(0) + 90

    .SelectionMode = flexSelectionByRow
    
    .row = 1
    .col = 0
    .ColSel = .Cols - 1
End With
    

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

With cmdF2
    .Width = TextWidth(.Caption) + 150
    .Height = wpara.ButtonY
    .Left = flxGesendet.Left + flxGesendet.Width + 150
    .Top = flxGesendet.Top
End With

With cmdF6
    .Width = cmdF2.Width
    .Height = cmdF2.Height
    .Left = cmdF2.Left
    .Top = cmdF2.Top + cmdF2.Height + 150
End With

Me.Width = cmdF2.Left + cmdF2.Width + 2 * wpara.LinksX

With cmdEsc
    .Top = flxGesendet.Top + flxGesendet.Height + 150 * wpara.BildFaktor
    .Width = wpara.ButtonX%
    .Height = wpara.ButtonY%
    .Left = (ScaleWidth - .Width) / 2
End With

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY% + 90 + wpara.FrmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2


h$ = "gesendete Bestellung: "
iLief% = Val(Left$(GesendetDatei$, 3))
If (iLief% > 0) And (iLief% <= lif.AnzRec) Then
    lif.GetRecord (iLief% + 1)
    h2$ = RTrim$(lif.Name(0))
    
    iRufzeit% = Val(Mid$(GesendetDatei$, 4, 4))
    h2$ = h2$ + "  (" + Format(iRufzeit% \ 100, "00") + ":" + Format(iRufzeit% Mod 100, "00") + ")"
    
    If (InStr(GesendetDatei$, "m.") > 0) Then h2$ = h2$ + "  manuell"
    h$ = h$ + h2$
End If
Caption = h$

If (MaxSendSatz% = 0) Then cmdF2.Enabled = False

Call DefErrPop
End Sub

