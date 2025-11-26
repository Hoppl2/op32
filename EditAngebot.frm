VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEditAngebot 
   BorderStyle     =   0  'Kein
   Caption         =   " "
   ClientHeight    =   3210
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   4935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtEditAngebot 
      Height          =   495
      Index           =   3
      Left            =   3480
      TabIndex        =   4
      Text            =   "99.99"
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2160
      ScaleHeight     =   465
      ScaleWidth      =   825
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picLabel 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2520
      ScaleHeight     =   465
      ScaleWidth      =   825
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtEditAngebot 
      Height          =   495
      Index           =   2
      Left            =   2400
      TabIndex        =   3
      Text            =   "99.99"
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtEditAngebot 
      Height          =   495
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      Text            =   "999"
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtEditAngebot 
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Text            =   "999"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   390
      Left            =   3600
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   3600
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   840
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxEditAngebot 
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1931
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483633
      BackColorBkg    =   -2147483633
      Enabled         =   0   'False
      FocusRect       =   0
      HighLight       =   2
      GridLines       =   0
      ScrollBars      =   0
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmEditAngebot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BarPreis%
Dim OrgAep#

Private Const DefErrModul = "EDITANGEBOT.FRM"

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

AnzLocalAngebote% = AnzLocalAngebote% - 1
EditErg% = False
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
Dim i%, ind%, uhr%, st%, min%, falsch%
Dim h2$

If (AngebotTemporaer = 0) Then Call AngebotSpeichern
EditErg% = True
Unload Me

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

If (ActiveControl.Name = txtEditAngebot(0).Name) Then
    If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
    If (EditModus% <> 1) Then
        If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And ((EditModus% <> 4) Or (Chr$(KeyAscii) <> ".")) Then
            Beep
            KeyAscii = 0
        End If
    End If
End If

Call clsError.DefErrPop
End Sub

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
Dim i%, MaxWi%, wi%, wi1%, wi2%
Dim zr!
Dim c As Control

Call wPara1.InitFont(Me)

With flxEditAngebot
    .Cols = 1
    .Rows = 16
    .FixedRows = 1
    .FixedCols = 0
    
    .Top = 0
    .Left = 0
    .Height = .RowHeight(0) * 15 + 90   '16
    
    .RowHeight(12) = 0
    
    If (AngebotNeu%) Then
        .FormatString = "^Neu"
    Else
        .FormatString = "^Edit"
    End If
    .SelectionMode = flexSelectionByColumn
    .HighLight = flexHighlightAlways
    
    .FillStyle = flexFillSingle
    
    For i% = 0 To .Cols - 1
        .ColWidth(i%) = TextWidth("WWWWWWWW ")
        .ColAlignment(i%) = flexAlignRightCenter
    Next i%

    .Width = .ColWidth(0) + 45
    .row = 1
    .RowSel = .Rows - 1
End With

On Error Resume Next
For Each c In Controls
    If (TypeOf c Is TextBox) Then
        c.Width = TextWidth(c.text) + 90
'        c.text = ""
    End If
Next
On Error GoTo DefErr

With txtEditAngebot(1)
    .Left = flxEditAngebot.Width - 30 - TextWidth(" Stk") - .Width
    .Top = flxEditAngebot.RowPos(2)
    .text = Format(AngebotEditNm%, "0")
End With
With txtEditAngebot(0)
    .Left = flxEditAngebot.Width - TextWidth("  +  999  Stk") - .Width
    .Top = flxEditAngebot.RowPos(2)
    .text = Format(AngebotEditBm%, "0")
End With
With txtEditAngebot(2)
    .Left = txtEditAngebot(0).Left
    .Top = flxEditAngebot.RowPos(5)
    zr! = Abs(AngebotEditZr!)
    If (zr! = 123.45!) Then zr! = 0!
    .text = Format(zr!, "0.00")
    If (AngebotMitZr = 0) Then .Visible = False
    If (AEPorg# = 0#) Then .Visible = False
End With
With txtEditAngebot(3)
    .Left = txtEditAngebot(0).Left
    .Top = flxEditAngebot.RowPos(6) + 30
    .text = "0"
    If (AngebotMitZr = 0) Then .Visible = False
End With

AnzLocalAngebote% = AnzLocalAngebote% + 1

With frmAngebote
    Me.Left = .Left + wPara1.FrmBorderHeight + .flxAngebote.Left
    Me.Top = .Top + wPara1.FrmBorderHeight + wPara1.FrmCaptionHeight + .flxAngebote.Top
    
    If (AngebotNeu%) Then
        If (.flxAngebote.Cols > 1) Then
            .flxAngebote.ColWidth(1) = .flxAngebote.ColWidth(1) * 2
            .flxAngebote.LeftCol = 1
        End If
        
        Me.Left = Me.Left + .flxAngebote.ColPos(1) + 45
    Else
        Me.Left = Me.Left + .flxAngebote.ColPos(.flxAngebote.col) + 45
    End If
End With
            
cmdOk.Left = 10000
cmdEsc.Left = 10000

BarPreis% = 0

If (iNewLine) Then
    With flxEditAngebot
        .ScrollBars = flexScrollBarNone
        .BorderStyle = 0
        .Width = .Width - 45
        .Height = .Height - 90
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridFlat
        .GridColorFixed = .GridColor
        .BackColor = vbWhite
        .BackColorFixed = RGB(199, 176, 123)
        .BackColorSel = RGB(232, 217, 172)
        .ForeColorSel = vbBlack
        
'        .Left = .Left + iAdd
'        .Top = .Top + iAdd
    End With
    Me.Left = Me.Left - 45
End If
Me.Width = flxEditAngebot.Width
Me.Height = flxEditAngebot.Height

Call clsError.DefErrPop
End Sub

Private Sub txtEditAngebot_Change(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtEditAngebot_Change")
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
Dim i%
Dim AngAep#, OrgAep#, zr#
Dim h$

On Error Resume Next
If (ActiveControl.Index <> Index) Then Call clsError.DefErrPop: Exit Sub
On Error GoTo DefErr

BarPreis% = 0
If (Index = 3) Then
    With flxEditAngebot
        h$ = Trim(.TextMatrix(6, 0))
        If (h$ = "") Then h$ = Trim(.TextMatrix(4, 0))
        OrgAep# = clsOpTool.xVal(h$)
        AngAep# = clsOpTool.xVal(Trim(txtEditAngebot(3).text))
        If (OrgAep# <> AngAep#) Then
'            h$ = Trim(.TextMatrix(4, 0))
'            OrgAep# = clsOpTool.xVal(h$)
            If (OrgAep# > 0#) Then
                zr# = 100 - (AngAep# / OrgAep#) * 100
            Else
                zr# = 321.12
            End If
            txtEditAngebot(2).text = Format(zr# + 0.005, "0.00")
            
            BarPreis% = True
        End If
    End With
End If


Call AufrufAngebotZeigen

If (Index = 2) Then
    With flxEditAngebot
        h$ = Trim(.TextMatrix(6, 0))
        If (h$ = "") Then h$ = Trim(.TextMatrix(4, 0))
        txtEditAngebot(3).text = h$
    End With
End If

Call clsError.DefErrPop
End Sub

Private Sub txtEditAngebot_KeyPress(Index As Integer, KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtEditAngebot_KeyPress")
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

If (Chr$(KeyAscii) = ".") Then KeyAscii = Asc(",")

Call clsError.DefErrPop
End Sub

Private Sub txtEditAngebot_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtEditAngebot_GotFocus")
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
Dim i%, wi%, ind%
Dim h$, h2$

With picLabel
    .Visible = False
    
    If (Index = 0) Then
        h$ = "BM"
    ElseIf (Index = 1) Then
        h$ = "NR"
    ElseIf (Index = 2) Then
        h$ = "Zeilenrab. in %"
    Else
        h$ = "Ang.Preis"
    End If
    
    .Font.Name = wPara1.FontName(0)
    .Font.Size = wPara1.FontSize(0)
    .Width = Me.Width
'    .Height = flxarbeit(0).Height
    .Cls
    .CurrentY = 30
    .CurrentX = 30
    picLabel.Print h$;
    wi% = TextWidth(h$)
    .Width = .CurrentX + 60 ' wi% + 60
    picLabel.Print
    .Height = .CurrentY + 30 ' 150
    .Left = txtEditAngebot(Index).Left
    .Top = txtEditAngebot(Index).Top - .Height
    
    If (Index <> 3) Then .Visible = True
End With

With picInfo
    .Visible = False
    
    If (Index = 2) Then
        h$ = "ZeilRab." + vbCr + "FaktRab." + vbCr + "-" + vbCr + "GesRab."
    Else
        h$ = ""
    End If
    
    .Font.Name = wPara1.FontName(0)
    .Font.Size = wPara1.FontSize(0)
    .Width = Me.Width
    .Height = flxEditAngebot.Height
    .Cls
    .CurrentY = 30
    
    Do
        If (h$ = "") Then Exit Do
        
        ind% = InStr(h$, vbCr)
        If (ind% > 0) Then
            h2$ = Left$(h$, ind% - 1)
            h$ = Mid$(h$, ind% + 1)
        Else
            h2$ = h$
            h$ = ""
        End If
        
        .CurrentX = 30
        If (Left$(h2$, 1) = "-") Then
            picInfo.Line (.CurrentX, .CurrentY)-(.Width - 60, .CurrentY)
'            picInfo.Print
        Else
            picInfo.Print h2$
        End If
    Loop
    .Width = flxEditAngebot.Width
    .Height = .CurrentY + 30
    .Left = flxEditAngebot.Left
    .Top = flxEditAngebot.Top + flxEditAngebot.Height
    
    If (Index = 2) Then
        Me.Height = .Top + .Height
        .Visible = True
    Else
        Me.Height = flxEditAngebot.Top + flxEditAngebot.Height
    End If
End With

Call AufrufAngebotZeigen
With flxEditAngebot
    h$ = Trim(.TextMatrix(6, 0))
    If (h$ = "") Then h$ = Trim(.TextMatrix(4, 0))
    txtEditAngebot(3).text = h$
End With

With txtEditAngebot(Index)
'    h$ = .text
'    For i% = 1 To Len(h$)
'        If (Mid$(h$, i%, 1) = ",") Then Mid$(h$, i%, 1) = "."
'    Next i%
'    .text = h$
    .SelStart = 0
    .SelLength = Len(.text)
    If (Index = 3) Then
        EditModus% = 4
    Else
        EditModus% = 0
    End If
End With

Call clsError.DefErrPop
End Sub

Private Sub AufrufAngebotZeigen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("AufrufAngebotZeigen")
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
Dim ind%
Dim AepKalk#
                        
If (txtEditAngebot(0).Visible) Then
    With LocalAngebotRec(AnzLocalAngebote% - 1)
        .gh = AngebotNeuLief%
        .st = "M"
        .bm = Val(txtEditAngebot(0).text)
        .mp = Val(txtEditAngebot(1).text)
        .zr = clsOpTool.xVal(txtEditAngebot(2).text)
        If (AngebotEditZr! < 0) Then
            If (.zr = 0) Then .zr = 123.45!
            .zr = -.zr
        End If
        .bmOrg = .bm
        .mpOrg = .mp
        .LaufNr = AnzManuelle% + 1
        .IstManuell = 1
        .ghBest = AngebotGhBest%
        
        If (BarPreis%) Then
            .AepAngOrg = clsOpTool.xVal(txtEditAngebot(3).text)
            .zr = 200
            If (AngebotEditZr! < 0) Then
                .zr = -.zr
            End If
        Else
            .AepAngOrg = 0
        End If
        
        If ((.bm + .mp) > 0) Then
            AepKalk# = AngebotAuswerten#(AnzLocalAngebote% - 1)
            Call AngebotZeigen(AnzLocalAngebote% - 1, AepKalk#, flxEditAngebot, True, picInfo, ActiveControl.Index)
        End If
    End With
End If

Call clsError.DefErrPop
End Sub

Sub AngebotSpeichern()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("AngebotSpeichern")
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
Dim ind%, FabsErrf%, bm%
Dim FabsRecno&, mp&, lRecs&
Dim AepKalk#
                        
With clsManuelleAngebote1
    .pzn = AngebotPzn$
    
    .bm = Val(txtEditAngebot(0).text)
    .mp = Val(txtEditAngebot(1).text)
    .gh = AngebotNeuLief%
    .st = "M"
    .zr = clsOpTool.xVal(txtEditAngebot(2).text)
    If (AngebotEditZr! <= 0) Then
        If (.zr = 0!) Then .zr = 123.45!
        .zr = -.zr
    End If
    
    If (BarPreis%) Then
        .BarPreis = clsOpTool.xVal(txtEditAngebot(3).text)
        .zr = 200
        If (AngebotEditZr! < 0) Then
            .zr = -.zr
        End If
    Else
        .BarPreis = 0
    End If
    
    .Saisonal = 0
    .rest = String(Len(.rest), 0)
    
    If (AngeboteDbOk) Then
        If (AngebotNeu%) Then
            SQLStr = "INSERT INTO Angebote (Pzn,Gh,St,Bm,Mp,Zr,Saisonal,BarPreis)"
            SQLStr = SQLStr + " VALUES (" + .pzn + "," + CStr(.gh) + ",'" + .st + "'," + CStr(.bm) + "," + CStr(.mp) + "," + clsOpTool.uFormat(.zr, "0.00") + "," + CStr(.Saisonal) + "," + clsOpTool.uFormat(.BarPreis, "0.00")
            SQLStr = SQLStr + ")"
        Else
            SQLStr = "Update Angebote SET Pzn=" + .pzn + ",Gh=" + CStr(.gh) + ",St='" + .st + "',Bm=" + CStr(.bm) + ",Mp=" + CStr(.mp) + ",Zr=" + clsOpTool.uFormat(.zr, "0.00") + ",Saisonal=" + CStr(.Saisonal) + ",BarPreis=" + clsOpTool.uFormat(.BarPreis, "0.00")
        End If
        ManuellAngeboteDB1.ActiveConn.CommandTimeout = 120
        Call ManuellAngeboteDB1.ActiveConn.Execute(SQLStr, lRecs, adExecuteNoRecords)
    Else
        If (AngebotNeu%) Then
            FabsErrf% = .IndexInsert(0, AngebotPzn$, FabsRecno&)
            If (FabsErrf% = 0) Then
                .PutRecord (FabsRecno& + 1)
            End If
        Else
            .PutRecord (AngebotEditRecno& + 1)
        End If
    End If
End With

Call clsError.DefErrPop
End Sub

