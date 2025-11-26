VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDirektPartner 
   Caption         =   "Lagerstand Partner-Apos"
   ClientHeight    =   3555
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   4875
   Icon            =   "DirektPartner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4875
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   3240
      TabIndex        =   2
      Top             =   2760
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   2760
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid flxDirektPartner 
      Height          =   2280
      Left            =   0
      TabIndex        =   0
      Top             =   240
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
      HighLight       =   2
      GridLines       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmDirektPartner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AufteilungPzn$

Private Const DefErrModul = "DIREKTPARTNER.FRM"

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
Dim i%, iLief%
Dim h$

h$ = ""
With flxDirektPartner
    For i% = 1 To (.Rows - 1)
        If (.TextMatrix(i%, 0) <> "") Then
            h$ = h$ + Format(Val(.TextMatrix(i%, 2)), "000") + ","
        End If
    Next i%
End With

OpDirektPartner$ = h$

Unload Me

Call DefErrPop
End Sub

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

Private Sub flxDirektPartner_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxDirektPartner_Click")
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

With flxDirektPartner
    If (.row > 0) Then
        SendKeys " ", True
    End If
End With

Call DefErrPop
End Sub

Private Sub flxDirektPartner_KeyDown(KeyCode As Integer, Shift As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxDirektPartner_KeyDown")
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

If (KeyCode = vbKeySpace) Then
    With flxDirektPartner
        If (.TextMatrix(.row, 0) <> "") Then
            h$ = ""
        Else
            h$ = Chr$(214)
        End If
        .TextMatrix(.row, 0) = h$
        
        If (.row < .Rows - 1) Then
            .row = .row + 1
        End If
    End With
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
Dim i%, spBreite%

Call wpara.InitFont(Me)

AufteilungPzn$ = KorrPzn$

Caption = "Folgende PartnerApotheken berücksichtigen"

Font.Bold = False   ' True

With flxDirektPartner
    .Cols = 5
    .Rows = 2
    .FixedRows = 1
    .FixedCols = 0
    
    .FormatString = "|>Sort#|>Profil#|>Lief#|<Name"
    .Rows = 1
    .SelectionMode = flexSelectionByRow
    
    .ColWidth(0) = TextWidth("XX")
    .ColWidth(1) = 0    'TextWidth("999999")
    .ColWidth(2) = 0    'TextWidth("999999")
    .ColWidth(3) = TextWidth("99999999")
    .ColWidth(4) = TextWidth(String(35, "X"))

    .Top = wpara.TitelY
    .Left = wpara.LinksX
    .Height = .RowHeight(0) * 6 + 90
    
    spBreite% = 0
    For i% = 0 To .Cols - 1
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
    .Rows = 1
End With

Font.Bold = False   ' True

Me.Width = flxDirektPartner.Width + 2 * wpara.LinksX

With cmdOk
    .Width = wpara.ButtonX
    .Height = wpara.ButtonY
    .Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
    .Top = flxDirektPartner.Top + flxDirektPartner.Height + 150
    .Visible = True
End With
With cmdEsc
    .Width = wpara.ButtonX
    .Height = wpara.ButtonY
    .Left = cmdOk.Left + cmdEsc.Width + 300
    .Top = cmdOk.Top
End With

Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2

Call DirektPartnerBefuellen

Call DefErrPop
End Sub

Private Sub DirektPartnerBefuellen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DirektPartnerBefuellen")
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
Dim i%, iProfilNr%, iSortNr%
Dim OrgBm%, ActBm%, iSum%, iBm%, iRest%, MaxRow%, MaxRest%, iHeute%
Dim multi!, bm!
Dim h$, h2$, sName$, sLief$

Set OpPartnerRec = OpPartnerDB.OpenRecordset("PartnerProfile", dbOpenTable)
OpPartnerRec.Index = "Unique"
If (OpPartnerRec.RecordCount > 0) Then
    OpPartnerRec.MoveFirst
End If


'OpPartnerLiefs$ = h$

With flxDirektPartner
    .Rows = 1
    
    Do
        If (OpPartnerRec.EOF) Then Exit Do
        
        If (OpPartnerRec!BeiDirektbezug) Then
            iSortNr% = OpPartnerRec!IntSortNr
            iProfilNr% = OpPartnerRec!ProfilNr
            sLief$ = Format(OpPartnerRec!IntLiefNr, "0")
            sName$ = OpPartnerRec!Name
            
            h$ = Chr$(214)
            h$ = h$ + vbTab + Format(iSortNr%, "0")
            h$ = h$ + vbTab + Format(iProfilNr, "0")
            h$ = h$ + vbTab + sLief$
            h$ = h$ + vbTab + sName$
            .AddItem h$
        End If
        
        OpPartnerRec.MoveNext
    Loop
    
    .row = 1
    .col = 0
    .RowSel = .Rows - 1
    .ColSel = .col
    .Sort = 5
    
    .Height = .RowHeight(0) * .Rows + 90
    .ZOrder 0
    
    .FillStyle = flexFillRepeat
    .row = 1
    .col = 0
    .RowSel = .Rows - 1
    .ColSel = .col
    .CellFontName = "Symbol"
    .FillStyle = flexFillSingle
    
    .row = 1
    .col = 0
    .ColSel = .Cols - 1
End With
    
Call DefErrPop
End Sub

