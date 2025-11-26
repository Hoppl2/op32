VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDirektProtokoll 
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
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   360
      TabIndex        =   1
      Top             =   3000
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxDirektProt 
      Height          =   2280
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4022
      _Version        =   393216
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmDirektProtokoll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LiefInd%(10)
Dim MindDiff#(10)

Private Const DefErrModul = "DIREKTPROTOKOLL.FRM"

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
Dim i%, spBreite%, ind%, iLief%, iRufzeit%, row%, col%, iModus%, maxSp%, iToggle%
Dim DRUCKHANDLE%
Dim h$, h2$

Call wpara.InitFont(Me)

With flxDirektProt
    .Cols = 1
    .Rows = 9
    .FixedRows = 0
    .FixedCols = 0
    
    .Top = wpara.TitelY
    .Left = wpara.LinksX
    .Height = .RowHeight(0) * .Rows + 90
    
    .ColAlignment(0) = flexAlignLeftCenter
    .SelectionMode = flexSelectionByRow
    
    DRUCKHANDLE% = FileOpen("winw\" + GesendetDatei$, "I")
    Do While Not EOF(DRUCKHANDLE%)
        If (i% >= .Rows) Then .Rows = .Rows + 1
        Line Input #DRUCKHANDLE%, h$
        .TextMatrix(i%, 0) = h$
        i% = i% + 1
        spBreite% = TextWidth(h$)
        If (spBreite% > maxSp%) Then maxSp% = spBreite%
    Loop
    Close #DRUCKHANDLE%


    spBreite% = maxSp% + 900
    .ColWidth(0) = spBreite%
    .Width = spBreite% + 90
    
    .row = 0
    .col = 0
End With


Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

Me.Width = flxDirektProt.Left + flxDirektProt.Width + 2 * wpara.LinksX

With cmdEsc
    .Top = flxDirektProt.Top + flxDirektProt.Height + 150 * wpara.BildFaktor
    .Width = wpara.ButtonX%
    .Height = wpara.ButtonY%
    .Left = (ScaleWidth - .Width) / 2
End With

Me.Height = cmdEsc.Top + cmdEsc.Height + wpara.TitelY% + 90 + wpara.FrmCaptionHeight

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2


h$ = "Direktbezug-Automatik: "
iLief% = Val(Left$(GesendetDatei$, 3))
If (iLief% > 0) And (iLief% <= lif.AnzRec) Then
    lif.GetRecord (iLief% + 1)
    h2$ = RTrim$(lif.Name(0))
    
    iRufzeit% = Val(Mid$(GesendetDatei$, 4, 4))
    h2$ = h2$ + "  (" + Format(iRufzeit% \ 100, "00") + ":" + Format(iRufzeit% Mod 100, "00") + ")"
    
'    If (InStr(GesendetDatei$, "m.") > 0) Then h2$ = h2$ + "  manuell"
    h$ = h$ + h2$
End If
Caption = h$

Call DefErrPop
End Sub

