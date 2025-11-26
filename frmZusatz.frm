VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmZusatz 
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   4305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4305
   Begin VB.CommandButton cmdEsc 
      Caption         =   "ESC"
      Height          =   450
      Left            =   1800
      TabIndex        =   4
      Top             =   3600
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdF2 
      Caption         =   "Edit (F2)"
      Height          =   450
      Left            =   3000
      TabIndex        =   1
      Top             =   2760
      Width           =   1200
   End
   Begin VB.TextBox txtZusatz 
      BorderStyle     =   0  'Kein
      Height          =   255
      Index           =   0
      Left            =   0
      MaxLength       =   37
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid flxZusatz 
      Height          =   2280
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4022
      _Version        =   65541
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
Attribute VB_Name = "frmZusatz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEsc_Click()
Unload Me
End Sub

Private Sub cmdF2_Click()
Dim i%

flxZusatz.Visible = False
'For i% = 0 To 4
'    txtZusatz(i%).Visible = True
'Next i%
Call ZeigeTextBoxen
cmdF2.Enabled = False

cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

cmdOk.Cancel = False

cmdEsc.Cancel = True
cmdEsc.Visible = True

End Sub

Private Sub cmdOk_Click()
Dim ind%

If (txtZusatz(0).Visible) Then
    If (ActiveControl.Name = txtZusatz(0).Name) Then
        ind% = ActiveControl.index
        If (Trim(ActiveControl.text) = "") Or (ind% = 4) Then
            cmdOk.SetFocus
        Else
            txtZusatz(ind% + 1).SetFocus
        End If
    Else
        Call clsDialog.ZusatzSpeichern
        Unload Me
    End If
Else
    Unload Me
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ind%

If (KeyCode = vbKeyF2) Then
    cmdF2.Value = True
ElseIf (KeyCode = vbKeyDown) Then
    If (ActiveControl.Name = txtZusatz(0).Name) Then
        ind% = ActiveControl.index
        If (ind% < 4) Then
            txtZusatz(ind% + 1).SetFocus
        End If
        KeyCode = 0
    End If
ElseIf (KeyCode = vbKeyUp) Then
    If (ActiveControl.Name = txtZusatz(0).Name) Then
        ind% = ActiveControl.index
        If (ind% > 0) Then
            txtZusatz(ind% - 1).SetFocus
        End If
        KeyCode = 0
    End If
End If
End Sub

Private Sub Form_Load()

Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%
Dim h$, h2$, FormStr$

Call wPara1.InitFont(Me)


Font.Bold = False   ' True

With flxZusatz
    .Cols = 1
    .Rows = 2
    .FixedRows = 1
    .FixedCols = 0
    
    .Top = wPara1.TitelY
    .Left = wPara1.LinksX
    
    .FormatString = "Zusatztext"
    .Rows = 1
    .SelectionMode = flexSelectionByRow
    
    If (ZusatzFensterTyp$ = ZUSATZ_ARTIKEL) Then
        .ColWidth(0) = TextWidth(String(26, "W"))
        .Rows = 6
        .Height = .RowHeight(0) * 6 + 90
    ElseIf (ZusatzFensterTyp$ = ZUSATZ_LIEFERANTEN) Then
        .ColWidth(0) = TextWidth(String(40, "W"))
        .Rows = 101
        .Height = .RowHeight(0) * 11 + 90
    End If
    
    .Width = .ColWidth(0) + 90
End With

For i% = 1 To (flxZusatz.Rows - 2)
    Load txtZusatz(i%)
    txtZusatz(i%).TabIndex = i% + 2
Next i%

For i% = 0 To UBound(txtZusatz)
    With txtZusatz(i%)
'        .Top = flxZusatz.Top + i * flxZusatz.RowHeight(1)
        .Left = flxZusatz.Left
        .Height = flxZusatz.RowHeight(1)
        .Width = flxZusatz.Width
    End With
Next i%

Font.Name = wPara1.FontName(1)
Font.Size = wPara1.FontSize(1)

cmdF2.Width = TextWidth(cmdF2.Caption) + 150
cmdF2.Height = wPara1.ButtonY
cmdF2.Left = flxZusatz.Left + flxZusatz.Width + 150
cmdF2.Top = flxZusatz.Top

Me.Width = cmdF2.Left + cmdF2.Width + 2 * wPara1.LinksX

cmdOk.Width = wPara1.ButtonX
cmdOk.Height = wPara1.ButtonY
cmdOk.Left = (Me.ScaleWidth - cmdOk.Width) / 2
cmdOk.Top = flxZusatz.Top + flxZusatz.Height + 150

cmdEsc.Width = wPara1.ButtonX
cmdEsc.Height = wPara1.ButtonY
cmdEsc.Top = cmdOk.Top

Me.Height = cmdOk.Top + cmdOk.Height + wPara1.TitelY + 90 + wPara1.FrmCaptionHeight

flxZusatz.Visible = True
For i% = 0 To 4
    txtZusatz(i%).Visible = False
Next i%

cmdOk.Default = True
cmdOk.Cancel = True
cmdOk.Visible = True

End Sub

Sub ZeigeTextBoxen()
Dim i%

For i% = 0 To UBound(txtZusatz)
    With txtZusatz(i%)
        .Top = flxZusatz.Top + i * flxZusatz.RowHeight(1)
        .Visible = True
    End With
Next i%

End Sub
