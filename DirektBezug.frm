VERSION 5.00
Begin VB.Form frmDirektBezug 
   Caption         =   "Direktbezug für "
   ClientHeight    =   5415
   ClientLeft      =   3000
   ClientTop       =   705
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   7860
   Begin VB.CheckBox chk1 
      Caption         =   "automat. &Ausdruck"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
   End
   Begin VB.ListBox lst1 
      Height          =   645
      Left            =   2400
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2640
      TabIndex        =   6
      Top             =   3600
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   1200
      TabIndex        =   5
      Top             =   3600
      Width           =   1200
   End
   Begin VB.Label lbl1 
      Height          =   735
      Index           =   2
      Left            =   2280
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label lbl1 
      Caption         =   "&Übertragungsart"
      Height          =   735
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lbl1 
      Caption         =   "Anzahl Positionen"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmDirektBezug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "DIREKTBEZUG.FRM"

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

DirektBezugErg$ = ""
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
Dim h$

h$ = ""
If (chk1.Value) Then h$ = "*"
h$ = h$ + lst1.text
DirektBezugErg$ = h$
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
Dim i%, spBreite%, br%
Dim h$

Call DirektBezugBefuellen

Call wpara.InitFont(Me)

lbl1(0).Left = wpara.LinksX
lbl1(0).Top = wpara.TitelY
lbl1(1).Left = wpara.LinksX
lbl1(1).Top = lbl1(0).Top + lbl1(0).Height + 300
lbl1(2).Left = lbl1(1).Left + lbl1(1).Width + 300
lbl1(2).Top = lbl1(0).Top

With lst1
    .Left = lbl1(2).Left
    .Top = lbl1(1).Top
    .Height = 5 * TextHeight("Äg")
    For i% = 0 To (.ListCount - 1)
        .ListIndex = i%
        br% = TextWidth(.text)
        If (br% > spBreite%) Then spBreite% = br%
    Next i%
    .Width = spBreite% + 300
    .ListIndex = 0
End With

With chk1
    .Left = lbl1(0).Left
    .Top = lst1.Top + lst1.Height + 300
End With

Font.Bold = False   ' True

cmdOk.Top = chk1.Top + chk1.Height + 300
cmdEsc.Top = cmdOk.Top

Me.Width = lst1.Left + lst1.Width + 2 * wpara.LinksX

cmdOk.Width = wpara.ButtonX
cmdOk.Height = wpara.ButtonY
cmdEsc.Width = wpara.ButtonX
cmdEsc.Height = wpara.ButtonY
cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

Call DefErrPop
End Sub

Sub DirektBezugBefuellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DirektBezugBefuellen")
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
        
Call HoleLieferantenDaten(Lieferant%)
lifzus.GetRecord (Lieferant% + 1)

Me.Caption = Me.Caption + LiefName1$
lbl1(2).Caption = Str$(AnzBestellArtikel%)

With lst1
    If (lifzus.DirektBestModemKz) Then .AddItem "Modem: " + Trim(lifzus.DirektBestModem)
'    If (lifzus.DirektBestMailKz) Then .AddItem "eMail: " + Trim(lifzus.DirektBestMail)
'    If (lifzus.DirektBestFaxKz) And (lifzus.DirektBestComputerFaxKz) Then .AddItem "Computer-Fax: " + Trim(lifzus.DirektBestFax)
    If (lifzus.DirektBestDruckKz) Or (lifzus.DirektBestFaxKz) Or (.ListCount = 0) Then .AddItem "Faxfähiger Ausdruck"
End With

chk1.Value = lifzus.DirektBestDruckKz

Call DefErrPop
End Sub

