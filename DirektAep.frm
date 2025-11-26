VERSION 5.00
Begin VB.Form frmDirektAepKalk 
   Caption         =   "Kontrolle Rechnungs-AEP"
   ClientHeight    =   5880
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   5805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   5805
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Nächster"
      Height          =   450
      Index           =   1
      Left            =   1800
      TabIndex        =   13
      Top             =   4560
      Width           =   1200
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Voriger"
      Height          =   450
      Index           =   0
      Left            =   360
      TabIndex        =   12
      Top             =   4560
      Width           =   1200
   End
   Begin VB.TextBox txtDirektAepKalk 
      Alignment       =   1  'Rechts
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
      Index           =   5
      Left            =   2880
      MaxLength       =   9
      TabIndex        =   3
      Text            =   "9999999.9"
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtDirektAepKalk 
      Alignment       =   1  'Rechts
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
      Index           =   4
      Left            =   2880
      MaxLength       =   9
      TabIndex        =   1
      Text            =   "9999999.9"
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox txtDirektAepKalk 
      Alignment       =   1  'Rechts
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
      Index           =   3
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   11
      Text            =   "999.9"
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txtDirektAepKalk 
      Alignment       =   1  'Rechts
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
      Index           =   0
      Left            =   3000
      MaxLength       =   9
      TabIndex        =   5
      Text            =   "999999.99"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtDirektAepKalk 
      Alignment       =   1  'Rechts
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
      Index           =   2
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   9
      Text            =   "999.9"
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txtDirektAepKalk 
      Alignment       =   1  'Rechts
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
      Index           =   1
      Left            =   2880
      MaxLength       =   9
      TabIndex        =   7
      Text            =   "999999.99"
      Top             =   1110
      Width           =   615
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   3120
      TabIndex        =   14
      Top             =   5160
      Width           =   1200
   End
   Begin VB.Label lblDirektAepKalkHeader 
      Caption         =   "AEP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   15
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label lblDirektAepKalk 
      Caption         =   "(AEP*Stück) - Rabatte"
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
      Index           =   5
      Left            =   240
      TabIndex        =   2
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label lblDirektAepKalk 
      Caption         =   "AEP - Rabatte"
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
      Index           =   4
      Left            =   240
      TabIndex        =   0
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label lblDirektAepKalk 
      Caption         =   "Zeilen-Rabatt in %"
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
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label lblDirektAepKalk 
      Caption         =   "AEP"
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
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblDirektAepKalk 
      Caption         =   "Fakturen-Rabatt in %"
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
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label lblDirektAepKalk 
      Caption         =   "AEP * Stück"
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
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "frmDirektAepKalk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DirektAep#
Dim DirektLm%
Dim DirektFr!, DirektZr!

Dim rInd%

Dim AktTextBox%


Private Const DefErrModul = "DIREKTAEP.FRM"

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
Call SpeicherAepKalk
Unload Me
Call DefErrPop
End Sub

Private Sub cmdChange_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdChange_Click")
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

Call SpeicherAepKalk
With frmAction.flxarbeit(0)
    If (Index = 0) Then
        .row = .row - 1
    Else
        .row = .row + 1
    End If
    rInd% = SucheFlexZeile(False)
    If (rInd% > 0) Then
        Call InitControls
    End If
End With

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

If (KeyCode = vbKeyPageUp) Then
    If (cmdChange(0).Enabled) Then
        cmdChange(0).Value = True
    End If
ElseIf (KeyCode = vbKeyPageDown) Then
    If (cmdChange(1).Enabled) Then
        cmdChange(1).Value = True
    End If
End If

Call DefErrPop
End Sub

Private Sub SpeicherAepKalk()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SpeicherAepKalk")
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
Dim aep#
Dim h$

rInd% = SucheFlexZeile(False)
If (rInd% > 0) Then
    h$ = Trim(txtDirektAepKalk(4).text)
    ind% = InStr(h$, ",")
    If (ind% > 0) Then Mid$(h$, ind%, 1) = "."
    
    aep# = Val(h$)
    If (ww.WuAEP <> aep#) Then
        ww.WuAEP = aep#
        ww.PutRecord (rInd% + 1)
        Call ActProgram.AuslesenDateiSatz(rInd%, True)
    End If
End If

Call DefErrPop
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_KeyPress")
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

If (KeyAscii = vbKeyReturn) Then
    ind% = ActiveControl.Index + 1
    If (ind% > 5) Then ind% = 0
    txtDirektAepKalk(ind%).SetFocus
End If

If (ActiveControl.Name = txtDirektAepKalk(0).Name) Then
    If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
    If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And (Chr$(KeyAscii) <> ".") Then
        Beep
        KeyAscii = 0
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
Dim i%, MaxWi%, wi%, wi1%, wi2%
Dim fr!, zr!
Dim AepAng#
Dim h$, h2$
Dim c As Control

EditTxt$ = ""

rInd% = SucheFlexZeile(False)
    
lblDirektAepKalkHeader.Caption = "PANTOZOL 20 MG TAB MAGENSAFT  100 ST"

Call wpara.InitFont(Me)
   
On Error Resume Next
For Each c In Controls
    If (TypeOf c Is TextBox) Then
        c.Width = TextWidth(c.text) + 90
        c.text = ""
    End If
Next
On Error GoTo DefErr
  
Call InitControls

lblDirektAepKalkHeader.Top = wpara.TitelY%
txtDirektAepKalk(0).Top = lblDirektAepKalkHeader.Top + lblDirektAepKalkHeader.Height + 300
For i% = 1 To 5
    txtDirektAepKalk(i%).Top = txtDirektAepKalk(i% - 1).Top + txtDirektAepKalk(i% - 1).Height + 90
Next i%

lblDirektAepKalkHeader.Left = wpara.LinksX
lblDirektAepKalk(0).Left = wpara.LinksX
lblDirektAepKalk(0).Top = txtDirektAepKalk(0).Top
For i% = 1 To 5
    lblDirektAepKalk(i%).Left = lblDirektAepKalk(i% - 1).Left
    lblDirektAepKalk(i%).Top = txtDirektAepKalk(i%).Top
Next i%

MaxWi% = 0
For i% = 0 To 5
    wi% = lblDirektAepKalk(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

txtDirektAepKalk(0).Left = lblDirektAepKalk(0).Left + MaxWi% + 300
For i% = 1 To 5
    txtDirektAepKalk(i%).Left = txtDirektAepKalk(i% - 1).Left
Next i%

wi% = lblDirektAepKalkHeader.Left + lblDirektAepKalkHeader.Width
MaxWi% = txtDirektAepKalk(0).Left + txtDirektAepKalk(0).Width
If (wi% > MaxWi%) Then MaxWi% = wi%


''''''''''
With cmdChange(1)
    .Width = TextWidth(.Caption) + 150
    .Height = wpara.ButtonY
End With
With cmdChange(0)
    .Width = cmdChange(1).Width
    .Height = cmdChange(1).Height
End With

cmdChange(0).Top = lblDirektAepKalk(5).Top + lblDirektAepKalk(5).Height + 300
cmdChange(1).Top = cmdChange(0).Top
'''''''''


Me.Width = MaxWi% + 2 * wpara.LinksX

cmdChange(0).Left = (Me.Width - (cmdChange(0).Width * 2 + 300)) / 2
cmdChange(1).Left = cmdChange(0).Left + cmdChange(0).Width + 300


cmdEsc.Width = wpara.ButtonX
cmdEsc.Height = wpara.ButtonY

cmdEsc.Top = cmdChange(0).Top
cmdEsc.Left = (Me.Width - cmdEsc.Width) / 2

Me.Height = cmdChange(0).Top + cmdChange(0).Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight   '+ cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight
cmdEsc.Top = cmdEsc.Top + 1000

If (AepKalkX% = 0) And (AepKalkY% = 0) Then
    Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
    Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2
Else
    Me.Left = AepKalkX%
    If (Me.Left + Me.Width > Screen.Width) Then
        Me.Left = Screen.Width - Me.Width - 30
    End If
    Me.Top = AepKalkY%
    If (Me.Top + Me.Height > wpara.WorkAreaHeight) Then
        Me.Top = wpara.WorkAreaHeight - Me.Height - 30
    End If
End If

Call DefErrPop
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Form_QueryUnload")
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
Dim l&

If (Me.Left <> AepKalkX%) Then
    AepKalkX% = Me.Left
    If (AepKalkX% < 0) Then AepKalkX% = 0
    l& = WritePrivateProfileString(UserSection$, "AepKalkX", Str$(AepKalkX%), INI_DATEI)
End If
If (Me.Top <> AepKalkY%) Then
    AepKalkY% = Me.Top
    If (AepKalkY% < 0) Then AepKalkY% = 0
    l& = WritePrivateProfileString(UserSection$, "AepKalkY", Str$(AepKalkY%), INI_DATEI)
End If

Call DefErrPop
End Sub

Private Sub txtDirektAepKalk_Change(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtDirektAepKalk_Change")
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
Dim Wert#, AepAng#
Static BereitsAktiv%

If (BereitsAktiv%) Or (txtDirektAepKalk(Index).Visible = False) Then Call DefErrPop: Exit Sub
If (txtDirektAepKalk(Index).Enabled = False) Then Call DefErrPop: Exit Sub

BereitsAktiv% = True

Wert# = xVal(txtDirektAepKalk(Index).text)

If (Index = 0) Then
    DirektAep# = Wert#
'    txtDirektAepKalk(0).text = Format(DirektAep#, "0.00")
    txtDirektAepKalk(1).text = Format(DirektAep# * DirektLm%, "0.00")
    
    AepAng# = DirektAep# * (100# - (DirektZr! + DirektFr!)) / 100#
    txtDirektAepKalk(4).text = Format(AepAng#, "0.00")
    txtDirektAepKalk(5).text = Format(AepAng# * DirektLm%, "0.00")
ElseIf (Index = 1) Then
    If (DirektLm% > 0) Then
        DirektAep# = Wert# / DirektLm%
    Else
        DirektAep# = 0#
    End If
    txtDirektAepKalk(0).text = Format(DirektAep#, "0.00")
'    txtDirektAepKalk(1).text = Format(DirektAep# * DirektLm%, "0.00")
    
    AepAng# = DirektAep# * (100# - (DirektZr! + DirektFr!)) / 100#
    txtDirektAepKalk(4).text = Format(AepAng#, "0.00")
    txtDirektAepKalk(5).text = Format(AepAng# * DirektLm%, "0.00")
ElseIf (Index = 2) Then
    DirektFr! = Wert#
    If (DirektFr! > 0!) Then DirektBezugFaktRabatt# = DirektFr!
    AepAng# = DirektAep# * (100# - (DirektZr! + DirektFr!)) / 100#
    txtDirektAepKalk(4).text = Format(AepAng#, "0.00")
    txtDirektAepKalk(5).text = Format(AepAng# * DirektLm%, "0.00")
ElseIf (Index = 3) Then
    DirektZr! = Wert#
    AepAng# = DirektAep# * (100# - (DirektZr! + DirektFr!)) / 100#
    txtDirektAepKalk(4).text = Format(AepAng#, "0.00")
    txtDirektAepKalk(5).text = Format(AepAng# * DirektLm%, "0.00")
ElseIf (Index = 4) Then
    AepAng# = Wert#
'    DirektAep# = AepAng# / (100# - (DirektZr! + DirektFr!)) * 100#
'    txtDirektAepKalk(0).text = Format(DirektAep#, "0.00")
'    txtDirektAepKalk(1).text = Format(DirektAep# * DirektLm%, "0.00")
    If (DirektAep# > 0) Then
        DirektZr! = 100 - ((AepAng# / (100# - DirektFr!) * 100#) / DirektAep#) * 100#
    Else
        DirektZr! = 0
    End If
    txtDirektAepKalk(3).text = Format(DirektZr!, "0.0")
    txtDirektAepKalk(5).text = Format(AepAng# * DirektLm%, "0.00")
ElseIf (Index = 5) Then
    AepAng# = Wert# / DirektLm%
'    DirektAep# = AepAng# / (100# - (DirektZr! + DirektFr!)) * 100#
'    txtDirektAepKalk(0).text = Format(DirektAep#, "0.00")
'    txtDirektAepKalk(1).text = Format(DirektAep# * DirektLm%, "0.00")
    txtDirektAepKalk(4).text = Format(AepAng# * DirektLm%, "0.00")
    If (DirektAep# > 0) Then
        DirektZr! = 100 - ((AepAng# / (100# - DirektFr!) * 100#) / DirektAep#) * 100#
    Else
        DirektZr! = 0
    End If
    txtDirektAepKalk(3).text = Format(DirektZr!, "0.0")
End If

BereitsAktiv% = False

Call DefErrPop
End Sub

Private Sub txtDirektAepKalk_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtDirektAepKalk_GotFocus")
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

With txtDirektAepKalk(Index)
    .SelStart = 0
    .SelLength = Len(.text)
End With
AktTextBox% = Index

Call DefErrPop
End Sub

Private Sub InitControls()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InitControls")
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
Dim AepAng#
Dim SQLStr$


For i% = 0 To 5
    txtDirektAepKalk(i%).Enabled = False
Next i%

lblDirektAepKalkHeader.Caption = ww.txt

DirektZr! = ww.zr
If (DirektZr! <= -200) Then
    DirektZr! = DirektZr! + 200
ElseIf (DirektZr! >= 200) Then
    DirektZr! = DirektZr! - 200
End If
                
DirektFr! = DirektBezugFaktRabatt#
If (ww.zr < 0) Then
    If (DirektZr! = -123.45) Then DirektZr! = 0#
    DirektZr! = DirektZr! * (-1)
Else
    DirektFr! = 0!
End If

DirektAep# = 0  ' ww.WuAEP
SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + ww.pzn
Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
If (TaxeRec.EOF = False) Then
    DirektAep# = TaxeRec!EK / 100
End If
If (DirektAep# = 0) Then
    DirektAep# = ww.WuAEP / (100# - DirektFr!) / 100#
End If

DirektLm% = ww.WuRm

txtDirektAepKalk(0).text = Format(DirektAep#, "0.00")
txtDirektAepKalk(1).text = Format(DirektAep# * DirektLm%, "0.00")

txtDirektAepKalk(2).text = Format(DirektFr!, "0.0")
txtDirektAepKalk(3).text = Format(DirektZr!, "0.0")


'AepAng# = DirektAep# * (100# - (DirektZr! + DirektFr!)) / 100#
AepAng# = ww.WuAEP
txtDirektAepKalk(4).text = Format(AepAng#, "0.00")
txtDirektAepKalk(5).text = Format(AepAng# * DirektLm%, "0.00")

For i% = 0 To 5
    txtDirektAepKalk(i%).Enabled = True
Next i%


If (txtDirektAepKalk(0).Visible) Then
    If (ActiveControl.Name <> txtDirektAepKalk(0).Name) Then
        txtDirektAepKalk(AktTextBox%).SetFocus
    End If
    With ActiveControl
        .SelStart = 0
        .SelLength = Len(.text)
    End With
End If

With frmAction.flxarbeit(0)
    If (.row > 1) Then
        cmdChange(0).Enabled = True
    Else
        cmdChange(0).Enabled = False
    End If
    If (.row < (.Rows - 1)) Then
        cmdChange(1).Enabled = True
    Else
        cmdChange(1).Enabled = False
    End If
End With
        
For i% = 0 To 5
    txtDirektAepKalk(i%).Enabled = True
Next i%

Call DefErrPop
End Sub

