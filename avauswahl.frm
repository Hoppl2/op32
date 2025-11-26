VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlButton.ocx"
Begin VB.Form frmAvAuswahl 
   AutoRedraw      =   -1  'True
   Caption         =   "A+V  Parameter"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   9360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   9360
   Begin VB.PictureBox picToolTip 
      Appearance      =   0  '2D
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
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   825
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtOVP 
      Enabled         =   0   'False
      Height          =   615
      Left            =   5880
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "avauswahl.frx":0000
      Top             =   4560
      Width           =   2055
   End
   Begin VB.ListBox lstSortierung 
      Height          =   450
      Left            =   4680
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   8880
      Picture         =   "avauswahl.frx":0006
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   8640
      Picture         =   "avauswahl.frx":00BF
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   8400
      Picture         =   "avauswahl.frx":0173
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2640
      TabIndex        =   7
      Top             =   3600
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   1200
      TabIndex        =   6
      Top             =   3600
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxAvAuswahl 
      Height          =   2700
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   840
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   4763
      _Version        =   393216
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483633
      BackColorBkg    =   -2147483633
      FocusRect       =   0
      HighLight       =   2
      GridLines       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin MSFlexGridLib.MSFlexGrid flxAvAuswahl 
      Height          =   2640
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4657
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
   Begin MSFlexGridLib.MSFlexGrid flxAvAuswahl 
      Height          =   2640
      Index           =   2
      Left            =   5640
      TabIndex        =   3
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4657
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
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   495
      Left            =   2520
      TabIndex        =   13
      Top             =   4200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   495
      Left            =   1080
      TabIndex        =   14
      Top             =   4200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin VB.Label lblOVP 
      Caption         =   "&OVP"
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblAvAuswahl 
      Caption         =   "&Verordnung"
      Height          =   375
      Index           =   2
      Left            =   6360
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblAvAuswahl 
      Caption         =   "Ver&einbarung  (Leertaste ... beitrittspflichtigen ('roten') Vereinbarungen/Pauschalen beitreten)"
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblAvAuswahl 
      Caption         =   "&Bundesland"
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmAvAuswahl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iEditModus%

Private Const DefErrModul = "AVAUSWAHL.FRM"

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
Dim i%
Dim l&
Dim h$, h2$, LeftLif$
        
With flxAvAuswahl(0)
    ActBundesland% = Val(.TextMatrix(.row, 1))
'    ActBundeslandInd% = .row
End With
With flxAvAuswahl(1)
'    ActKasse% = Val(.TextMatrix(.row, 1))
    If (InStr(UCase(.TextMatrix(.row, 0)), "PAUSCHALE:") > 0) Then
        ActPauschaleNr = Val(.TextMatrix(.row, 1))
        ActVebNr = 0
    Else
        ActVebNr = Val(.TextMatrix(.row, 1))
        ActPauschaleNr = 0
    End If
    
    If (InStr(UCase(.TextMatrix(.row, 0)), "BERUFSGENOSSENSCHAFT") > 0) Then
        Call ActProgram.MachGebFrei(.TextMatrix(.row, 0))
    ElseIf (InStr(UCase(.TextMatrix(.row, 0)), "BUNDESWEHR") > 0) Then
        Call ActProgram.MachGebFrei
    End If
    
'    h = ""
'    For i = 0 To (.Rows - 1)
'        If (Val(.TextMatrix(i, 2)) > 0) Then
'            If (Val(.TextMatrix(i, 3)) > 0) Then
'                h2 = .TextMatrix(i, 1)
'                If (InStr(UCase(.TextMatrix(i, 0)), "PAUSCHALE:") > 0) Then
'                    h2 = "-" + h2
'                End If
'                h = h + h2 + ","
'            End If
'        End If
'    Next i
'    l& = WritePrivateProfileString("Rezeptkontrolle", "Beigetreten", h$, INI_DATEI)
'    BeigetreteneVereinbarungen$ = "," + h$
End With
With flxAvAuswahl(2)
    ActVerordnung% = Val(.TextMatrix(.row, 1))
End With

Unload Me

Call DefErrPop
End Sub

Private Sub flxAvAuswahl_Click(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxAvAuswahl_Click")
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

If (index = 0) Then
    Dim i%, j%, iLand%, ind%
    Dim IstVebNr&, SollVebNr&
    Dim h$, h2$
    
    With frmAvAuswahl.flxAvAuswahl(0)
        iLand = Val(.TextMatrix(.row, 1))
    End With
        
    With frmAvAuswahl.flxAvAuswahl(1)
        .Redraw = False
        .Rows = 0
        .Cols = 6
    '    For i% = 0 To UBound(AvKassen$)
    '        h$ = Mid$(AvKassen$(i%), 3) + vbTab + Left$(AvKassen$(i%), 2)
    '        .AddItem h$
    '    Next i%
    '    .row = 0
    '    For i% = 0 To (.Rows - 1)
    '        If (Val(.TextMatrix(i%, 1)) = ActKasse%) Then
    '            .row = i%
    '            Exit For
    '        End If
    '    Next i%
    '    ActKasseInd% = .row
    
        lstSortierung.Clear
    
    
        If (AutIdemIk& > 0) Then    'And (ActVebNr = 0) And (ActPauschaleNr = 0) Then
            SQLStr = "SELECT VdbVVI.VebNr FROM VdbVVI LEFT JOIN VdbVVL ON VdbVVI.VebNr=VdbVVL.VebNr"
            SQLStr = SQLStr + " WHERE VdbVVI.Ik=" + CStr(AutIdemIk)
            SQLStr = SQLStr + " AND (VdbVVL.LandNr=" + CStr(iLand) + " OR VdbVVL.LandNr=18)"
        Else
            SQLStr = "SELECT * FROM VdbVVL"
        '    SQLStr = SQLStr + " LEFT JOIN VdbVereinbarungen ON VdbVVL.VebNr=VdbVereinbarungen.VebNr"
            SQLStr = SQLStr + " WHERE VdbVVL.LandNr=" + CStr(iLand) + " OR VdbVVL.LandNr=18"
        End If
        SQLStr = SQLStr + " ORDER BY VebNr"
        FabsErrf = AplusVDB.OpenRecordset(VdbPznRec2, SQLStr)
        Do
            If (VdbPznRec2.EOF) Then
                Exit Do
            End If
            
            SollVebNr = CheckNullLong(VdbPznRec2!VebNr)
            h$ = "(leer)" + vbTab + "0"
            For j% = 0 To UBound(AvVereinbarungen$)
                h2 = AvVereinbarungen(j)
                IstVebNr = Val(Left(h2, 10))
                If (IstVebNr = SollVebNr) Then
'                    h$ = Trim(Mid$(h2, 12)) + vbTab + CStr(IstVebNr) + vbTab + Mid(h2, 11, 1) + vbTab
'                    h = h + CStr(Abs(InStr(BeigetreteneVereinbarungen, "," + CStr(IstVebNr) + ",") > 0))
'                    .AddItem h$
                    h$ = Trim(Mid$(h2, 12, 100)) + vbTab + CStr(IstVebNr) + vbTab + Mid(h2, 11, 1) + vbTab
                    h = h + CStr(Abs(InStr(BeigetreteneVereinbarungen, "," + CStr(IstVebNr) + ",") > 0)) + vbTab
                    h = h + Trim(Mid$(h2, 112)) + Chr(10) + Left(h2, 10)
                    lstSortierung.AddItem h
                    
'                    For i% = 0 To UBound(AvPauschalen$)
'                        h2 = AvPauschalen(i)
'                        IstVebNr = Val(Mid(h2, 112, 10))
'                        If (IstVebNr = SollVebNr) Then
'                            h$ = Space(3) + Trim(Mid$(h2, 12, 100)) + vbTab + CStr(Val(Left(h2, 10))) + vbTab + Mid(h2, 11, 1) + vbTab
'                            h = h + CStr(Abs(InStr(BeigetreteneVereinbarungen, "," + "-" + CStr(Val(Left(h2, 10))) + ",") > 0))
'                            .AddItem h$
'                        End If
'                    Next i%
                    
                    Exit For
                End If
            Next j%
            
            VdbPznRec2.MoveNext
        Loop
        
        For j = 0 To (lstSortierung.ListCount - 1)
            h = lstSortierung.List(j)
            ind = InStr(h, Chr(10))
            SollVebNr = Val(Mid(h, ind + 1))
            h = Left(h, ind - 1)
            .AddItem h$
                    
            For i% = 0 To UBound(AvPauschalen$)
                h2 = AvPauschalen(i)
                IstVebNr = Val(Mid(h2, 112, 10))
                If (IstVebNr = SollVebNr) Then
                    h$ = Space(3) + Trim(Mid$(h2, 12, 100)) + vbTab + CStr(Val(Left(h2, 10))) + vbTab + Mid(h2, 11, 1) + vbTab
                    h = h + CStr(Abs(InStr(BeigetreteneVereinbarungen, "," + "-" + CStr(Val(Left(h2, 10))) + ",") > 0))
                    .AddItem h$
                End If
            Next i%
        Next j

'        If (VdbPznRec.RecordCount <> 0) Then
'            VdbPznRec.MoveFirst
'
'            Do
'                If (VdbPznRec.EOF) Then
'                    Exit Do
'                End If
'
'                SollVebNr = CheckNullLong(VdbPznRec!VebNr)
'                h$ = "(leer)" + vbTab + "0"
'                For i% = 0 To UBound(AvPauschalen$)
'                    h2 = AvPauschalen(i)
'                    IstVebNr = Val(Mid(h2, 111, 10))
'                    If (IstVebNr = SollVebNr) Then
'                        h$ = Trim(Mid$(h2, 11, 100)) + vbTab + CStr(Val(Left(h2, 10)))
'                        .AddItem h$
'                    End If
'                Next i%
'
'                VdbPznRec.MoveNext
'            Loop
'        End If
        
        VdbPznRec2.Close
        
        If (.Rows = 0) Then
            .AddItem "keine Vereinbarung gefunden" + vbTab
        End If
        
        .FillStyle = flexFillRepeat
        For i = 0 To (.Rows - 1)
            .row = i
            .col = 0
            .RowSel = .row
            .ColSel = .Cols - 1
'            .CellFontItalic = (Val(.TextMatrix(i, 2)) > 0)
            If (Val(.TextMatrix(i, 2)) > 0) And (Val(.TextMatrix(i, 3)) = 0) Then
                .CellForeColor = vbRed
            End If
        Next i
        .FillStyle = flexFillSingle
        .Redraw = True
        
        
        .row = 0
        For i% = 0 To (.Rows - 1)
            If (Val(.TextMatrix(i%, 1)) = ActVebNr) Then
                .row = i%
                Exit For
            End If
        Next i%
        
'        ActKasseInd% = .row
    End With
End If

Call DefErrPop
End Sub

Private Sub flxAvauswahl_DblClick(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxAvAuswahl_DblClick")
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
cmdOk.Value = True
Call DefErrPop
End Sub

Private Sub flxAvAuswahl_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxAvAuswahl_GotFocus")
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
Dim iRow%

With flxAvAuswahl(index)
    .HighLight = flexHighlightAlways
    
    iRow% = .row
    .Redraw = False
    .FillStyle = flexFillRepeat
    .row = 0
    .col = 0
    .RowSel = .Rows - 1
    .ColSel = .Cols - 1
    .CellFontBold = False
    .FillStyle = flexFillSingle
    .Redraw = True
    .row = iRow%
    .col = 0
    .ColSel = .Cols - 1
End With

Call DefErrPop
End Sub

Private Sub flxAvAuswahl_lostFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxAvAuswahl_lostFocus")
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

With flxAvAuswahl(index)
    .HighLight = flexHighlightNever
    
    .FillStyle = flexFillRepeat
    .col = 0
    .RowSel = .row
    .ColSel = .Cols - 1
    .CellFontBold = True
    .FillStyle = flexFillSingle
End With

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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%, lief%, MaxHe%, rHe%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%, FeldInd%
Dim iAdd%, iAdd2%
Dim h$, h2$, FormStr$
Dim c As Control
Dim LastOVP As Date

Call wpara.InitFont(Me)

iEditModus = 1

Call ActProgram.FlxAvAuswahlBefuellen
Call flxAvAuswahl_Click(0)

rHe% = flxAvAuswahl(0).RowHeight(0)
MaxHe% = frmAction.ScaleHeight - 2 * wpara.FrmCaptionHeight
MaxHe% = MaxHe% - wpara.FrmCaptionHeight - wpara.ButtonY - 3 * wpara.TitelY
If (para.Newline) Then
    MaxHe = MaxHe - wpara.ButtonY
End If
MaxHe% = ((MaxHe% - 90) \ rHe%) * rHe% + 90

For j% = 0 To 2
    With flxAvAuswahl(j%)
        Breite1% = 0
        If (j = 1) Then
            Breite1 = TextWidth(String(60, "X"))
        End If
        For i% = 0 To (.Rows - 1)
            Breite2% = TextWidth(.TextMatrix(i%, 0))
            If (Breite2% > Breite1%) Then Breite1% = Breite2%
        Next i%
        .ColWidth(0) = Breite1% + 150
        If (j = 1) Then
            .ColWidth(1) = 0 ' TextWidth(String(8, "0"))
        Else
            .ColWidth(1) = TextWidth("00000")
        End If
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        If (j = 1) Then
            .ColWidth(4) = TextWidth(String(9, "0"))
        Else
            .ColWidth(4) = 0
        End If
        .ColWidth(5) = 0
        
        .Height = .RowHeight(0) * .Rows + 90
        If (.Height > MaxHe%) Or (j = 1) Then
            .Height = MaxHe%
            .ColWidth(5) = wpara.FrmScrollHeight
        End If
        
        .Width = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5) + 90
        
        lblAvAuswahl(j%).Top = wpara.TitelY
        
        If (j% = 0) Then
            .Top = lblAvAuswahl(j%).Top + lblAvAuswahl(j%).Height + 60
            .Left = wpara.LinksX
        Else
            .Top = flxAvAuswahl(j% - 1).Top
            .Left = flxAvAuswahl(j% - 1).Left + flxAvAuswahl(j% - 1).Width + 600
        End If
        
        .SelectionMode = flexSelectionByRow
        .col = 0
        .ColSel = .Cols - 1
'        .row = 0
        Call flxAvAuswahl_lostFocus(j%)
        
        lblAvAuswahl(j%).Left = .Left
    End With
Next j%

With lblOVP
    .Left = lblAvAuswahl(2).Left
    .Top = flxAvAuswahl(2).Top + flxAvAuswahl(2).Height + 900
End With

LastOVP = DateValue("01.01.1980")
h = Space(255)
l = GetPrivateProfileString("LetzteAktionen", "OvpDownload", "", h, 255, CurDir + "\SYSMANAG.INI")
If l > 0 Then
    h = Left(h, l)
    If IsDate(h) Then
        LastOVP = h
    End If
End If

With txtOVP
    .Left = flxAvAuswahl(2).Left
    .Top = lblOVP.Top + lblOVP.Height + 60
    .Width = flxAvAuswahl(2).Height
    .Height = .Height * 2 + 90
    .text = "Letzte Einspielung:" + vbCrLf + Format(LastOVP, "DD.MM.YYYY HH:MM")
    If (Year(LastOVP) < 2000) Then
        .Visible = False
        .tag = "0"
        lblOVP.Visible = False
    End If
End With


Font.Bold = False   ' True

cmdOk.Top = flxAvAuswahl(1).Top + flxAvAuswahl(1).Height + 150
cmdEsc.Top = cmdOk.Top

Me.Width = flxAvAuswahl(2).Left + flxAvAuswahl(2).Width + 2 * wpara.LinksX

cmdOk.Width = wpara.ButtonX
cmdOk.Height = wpara.ButtonY
cmdEsc.Width = wpara.ButtonX
cmdEsc.Height = wpara.ButtonY
cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

If (para.Newline) Then
    iAdd = wpara.NlFlexBackY
    iAdd2 = wpara.NlCaptionY
    
    For j% = 0 To 2
        With flxAvAuswahl(j%)
            If (j <> 1) Then
                .ScrollBars = flexScrollBarNone
            End If
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
    Next j
    
    cmdOk.Top = cmdOk.Top + 2 * iAdd
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
        .Top = flxAvAuswahl(1).Top + flxAvAuswahl(1).Height + iAdd + 600
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .Default = cmdEsc.Default
        .Cancel = cmdEsc.Cancel
        .Visible = True
    End With
    cmdEsc.Visible = False

    With nlcmdOk
        .Init
'        .Left = Me.ScaleWidth - .Width - 150
        .Top = nlcmdEsc.Top
        .Caption = cmdOk.Caption
        .TabIndex = cmdOk.TabIndex
        .Enabled = cmdOk.Enabled
        .Default = cmdOk.Default
        .Cancel = cmdOk.Cancel
        .Visible = True
    End With
    cmdOk.Visible = False

    nlcmdOk.Left = (Me.Width - (nlcmdOk.Width * 2 + 300)) / 2
    nlcmdEsc.Left = nlcmdOk.Left + nlcmdEsc.Width + 300

    Me.Height = nlcmdEsc.Top + nlcmdEsc.Height + wpara.FrmCaptionHeight + 450

    Call wpara.NewLineWindow(Me, nlcmdEsc.Top)
    For i = 0 To 2
        With flxAvAuswahl(i)
            RoundRect hdc, (.Left - iAdd) / Screen.TwipsPerPixelX, (.Top - iAdd) / Screen.TwipsPerPixelY, (.Left + .Width + iAdd) / Screen.TwipsPerPixelX, (.Top + .Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
        End With
    Next i

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
    nlcmdOk.Visible = False
    nlcmdEsc.Visible = False
End If
'''''''''

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2

Call DefErrPop
End Sub

Private Sub flxavauswahl_KeyPress(index As Integer, KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxAvAuswahl_KeyPress")
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
Dim i%, row%, gef%, col%, ind%
Dim lColor&, l&
Dim ch$, h$, h2

ch$ = UCase$(Chr$(KeyAscii))

If (index = 1) And (ch = " ") Then
    With flxAvAuswahl(1)
        If (Val(.TextMatrix(.row, 2)) > 0) Then
            If (Val(.TextMatrix(.row, 3)) > 0) Then
                ch = "0"
                lColor = vbRed
            Else
                ch = "1"
                lColor = .ForeColor
            End If
            .TextMatrix(.row, 3) = ch
            .FillStyle = flexFillRepeat
            .col = 0
            .RowSel = .row
            .ColSel = .Cols - 1
            .CellForeColor = lColor
            .FillStyle = flexFillSingle
            
            h2 = .TextMatrix(.row, 1)
            If (InStr(UCase(.TextMatrix(.row, 0)), "PAUSCHALE:") > 0) Then
                h2 = "-" + h2
            End If
            ind = InStr(BeigetreteneVereinbarungen, "," + h2 + ",")
            If (ch = "0") And (ind > 0) Then
                BeigetreteneVereinbarungen = Left(BeigetreteneVereinbarungen, ind) + Mid(BeigetreteneVereinbarungen, ind + Len(h2) + 2)
            ElseIf (ch = "1") And (ind <= 0) Then
                BeigetreteneVereinbarungen = BeigetreteneVereinbarungen + h2 + ","
            End If
            
            If (BeigetreteneVereinbarungen <> "") Then
                If (Left(BeigetreteneVereinbarungen, 1) = ",") Then
                    BeigetreteneVereinbarungen = Mid(BeigetreteneVereinbarungen, 2)
                End If
                If (Right(BeigetreteneVereinbarungen, 1) = ",") Then
                    BeigetreteneVereinbarungen = Left(BeigetreteneVereinbarungen, Len(BeigetreteneVereinbarungen) - 1)
                End If
            End If
            
            l& = WritePrivateProfileString("Rezeptkontrolle", "Beigetreten", BeigetreteneVereinbarungen, INI_DATEI)
            BeigetreteneVereinbarungen$ = "," + BeigetreteneVereinbarungen + ","
        End If
    End With
ElseIf (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890", ch$) > 0) Then
    gef% = False
    With flxAvAuswahl(index)
        row% = .row
        For i% = (row% + 1) To (.Rows - 1)
            If (UCase(Left$(.TextMatrix(i%, 0), 1)) = ch$) Then
                .row = i%
                gef% = True
                Exit For
            End If
        Next i%
        If (gef% = False) Then
            For i% = 1 To (row% - 1)
                If (UCase(Left$(.TextMatrix(i%, 0), 1)) = ch$) Then
                    .row = i%
                    gef% = True
                    Exit For
                End If
            Next i%
        End If
        If (gef% = True) Then
'            If (.row < .TopRow) Then .TopRow = .row
            .TopRow = .row
        End If
    End With
End If

Call DefErrPop
End Sub

Private Sub flxavauswahl_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxavauswahl_MouseMove")
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
Dim i%, ind%
Dim h$

If (index = 1) Then
    h = ""
    On Error Resume Next
    With flxAvAuswahl(index)
        ind = y \ .RowHeight(0) + .TopRow
        h = .TextMatrix(ind, 1)
    End With
    On Error GoTo DefErr
    
    'flxInfo(0).TextMatrix(2, 0) = h
    
    If (h <> "") Then
        With picToolTip
            .Width = .TextWidth(h$ + "x")
            .Height = .TextHeight(h$) + 45
            .Left = x + flxAvAuswahl(index).Left + 300
            .Top = y + flxAvAuswahl(0).Top + 150
            .Visible = True
            .Cls
            .CurrentX = 2 * Screen.TwipsPerPixelX
            .CurrentY = 0
            picToolTip.Print h$
        End With
    End If
End If

Call DefErrPop
End Sub

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

If (picToolTip.Visible = True) Then
    picToolTip.Visible = False
End If

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

Private Sub nlcmdOk_Click()
Call cmdOk_Click
End Sub

Private Sub nlcmdEsc_Click()
Call cmdEsc_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If (para.Newline) Then
    If (KeyAscii = 13) Then
        Call nlcmdOk_Click
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





