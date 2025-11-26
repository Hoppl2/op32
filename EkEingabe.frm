VERSION 5.00
Begin VB.Form frmEkEingabe 
   ClientHeight    =   2910
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   3930
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   3930
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2160
      TabIndex        =   2
      Top             =   1920
      Width           =   1200
   End
   Begin VB.TextBox txtEkEingabe 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Text            =   "9999"
      Top             =   270
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   600
      TabIndex        =   1
      Top             =   1920
      Width           =   1200
   End
   Begin VB.Label lblEkEingabeZusatz 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label lblEkEingabe 
      Caption         =   "&Bestellmenge (1 - 9999)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "frmEkEingabe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const DefErrModul = "EKEINGABE.FRM"

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

EkEingabeErg% = False
Unload Me

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

EkEingabePreis# = Val(txtEkEingabe.text)    'cdbl

If (EkEingabeModus%) And (EkEingabePreis# = 0#) Then
    Beep
Else
    EkEingabeErg% = True
    Unload Me
End If

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
If (ActiveControl.Name = txtEkEingabe.Name) Then
    If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
    If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And (Chr$(KeyAscii) <> ".") Then
        Beep
        KeyAscii = 0
    End If
End If
Call clsError.DefErrPop
End Sub

'Static Sub NNEkEingabe(text$, aep#, NNAEP#, NNAEP%, Art%)
'If NNAEP% Then
'  t$ = fnzent$("E i n g a b e   d e s   N e t t o - A E P")
'  LOCATE 8, 3: Call farbe("Ni"): Print t$ + Space$(75 - Len(t$));: Call farbe
'  LOCATE 10, 2: Print "PZN     Artikel"; Space$(40); "Netto-AEP pro Stück";
'  LOCATE 12, 2: Print USING; "& #####.## "; text$; aep#;
'  If Art% = 2 Then
'    LOCATE 14, 2: Print "Großhandelsangebot mit Teillieferung!";
'  ElseIf Art% = -1 Then
'    LOCATE 14, 2: Print "Lieferantenkonditionen fehlen oder AEP gleich Null";
'  ElseIf Art% = -2 Then
'    LOCATE 14, 2: Print "Teillieferung bei Bestellung mit Naturalrabatt!";
'  End If
'NNEin:
'  xp$ = "6112083": X$ = Str$(NNAEP#): Call ein500: NNAEP# = Val(X$)
'  If Abs(NNAEP#) < 0.01 Then Beep: GoTo NNEin     '3.11
'  LOCATE 14, 2: Print Space$(78);
'Else
'  t$ = fnzent$("E i n g a b e   d e s   A E P")
'  LOCATE 8, 3: Call farbe("Ni"): Print t$ + Space$(75 - Len(t$));: Call farbe
'  LOCATE 10, 2: Print "PZN     Artikel"; Space$(41); "AEP pro Stück";
'  LOCATE 12, 2: Print USING; "& #####.## "; text$; aep#;
'  xp$ = "6112083": X$ = Str$(aep#): Call ein500: aep# = Val(X$)
'End If
'LOCATE 8, 1: Print Space$(79): LOCATE 10, 1: Print Space$(79): LOCATE 12, 1: P
'End Sub

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
Dim i%, MaxWi%, MaxHe%, wi%, wi1%, wi2%
Dim h$

If (EkEingabeModus%) Then
    h$ = "Eingabe des NN-Aep"
Else
    h$ = "Eingabe des Aep"
End If
h$ = Left$(h$ + Space$(20), 30)
lblEkEingabe.Caption = h$
txtEkEingabe.text = Format(EkEingabePreis#, "0.00")

h$ = ""
If (EkEingabeModus%) Then
    If (EkEingabeArt% = 2) Then
        h$ = "Großhandelsangebot mit Teillieferung!"
    ElseIf (EkEingabeArt% = -1) Then
        h$ = "Lieferantenkonditionen fehlen oder AEP gleich Null"
    ElseIf (EkEingabeArt% = -2) Then
        h$ = "Teillieferung bei Bestellung mit Naturalrabatt!"
    End If
End If
lblEkEingabeZusatz.Caption = h$

Call wPara1.InitFont(Me)


lblEkEingabe.Top = wPara1.TitelY%
lblEkEingabe.Left = wPara1.LinksX
txtEkEingabe.Top = lblEkEingabe.Top
txtEkEingabe.Left = lblEkEingabe.Left + lblEkEingabe.Width + 150
lblEkEingabeZusatz.Top = lblEkEingabe.Top + lblEkEingabe.Height + 150
lblEkEingabeZusatz.Left = wPara1.LinksX


cmdOk.Width = wPara1.ButtonX
cmdOk.Height = wPara1.ButtonY
cmdOk.Top = lblEkEingabeZusatz.Top + lblEkEingabeZusatz.Height + 1500

cmdEsc.Width = cmdOk.Width
cmdEsc.Height = cmdOk.Height
cmdEsc.Top = cmdOk.Top


If (EkEingabeModus%) Then
    MaxHe% = lblEkEingabeZusatz.Top + lblEkEingabeZusatz.Height
Else
    MaxHe% = lblEkEingabe.Top + lblEkEingabe.Height
End If

Me.Height = MaxHe% + wPara1.TitelY + 90 + wPara1.FrmCaptionHeight

wi1% = txtEkEingabe.Left + txtEkEingabe.Width
wi2% = lblEkEingabeZusatz.Width
If (wi1% > wi2%) Then
    MaxWi% = wi1%
Else
    MaxWi% = wi2%
End If
Me.Width = MaxWi% + 3 * wPara1.LinksX

Me.Left = ProjektForm.Left + (ProjektForm.Width - Me.Width) / 2
Me.Top = ProjektForm.Top + (ProjektForm.Height - Me.Height) / 2

cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

Me.Caption = EkEingabeText$

Call clsError.DefErrPop
End Sub

Private Sub txtekeingabe_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtEkEingabe_GotFocus")
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
Dim i%, h$

With txtEkEingabe
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
