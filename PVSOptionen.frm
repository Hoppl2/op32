VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPVSOptionen 
   Caption         =   "Optionen"
   ClientHeight    =   5985
   ClientLeft      =   540
   ClientTop       =   735
   ClientWidth     =   8565
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8565
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2280
      TabIndex        =   8
      Top             =   6240
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   840
      TabIndex        =   7
      Top             =   6240
      Width           =   1200
   End
   Begin TabDlg.SSTab tabOptionen 
      Height          =   5325
      Left            =   240
      TabIndex        =   9
      Top             =   360
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   9393
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&1 - Allgemein"
      TabPicture(0)   =   "PVSOptionen.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fmeOptionen(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&2 - A+V Taxierung"
      TabPicture(1)   =   "PVSOptionen.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fmeOptionen(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&3 - Tätigkeiten"
      TabPicture(2)   =   "PVSOptionen.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fmeOptionen(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&4 - Abrechnungsdaten"
      TabPicture(3)   =   "PVSOptionen.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fmeOptionen(3)"
      Tab(3).ControlCount=   1
      Begin VB.Frame fmeOptionen 
         Height          =   4215
         Index           =   3
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   8895
         Begin MSFlexGridLib.MSFlexGrid flxOptionen1 
            Height          =   2700
            Index           =   2
            Left            =   840
            TabIndex        =   16
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
      End
      Begin VB.Frame fmeOptionen 
         Height          =   4215
         Index           =   2
         Left            =   -75000
         TabIndex        =   13
         Top             =   840
         Width           =   8895
         Begin MSFlexGridLib.MSFlexGrid flxOptionen1 
            Height          =   2700
            Index           =   1
            Left            =   840
            TabIndex        =   14
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
      End
      Begin VB.Frame fmeOptionen 
         Height          =   4215
         Index           =   1
         Left            =   -74520
         TabIndex        =   11
         Top             =   480
         Width           =   8895
         Begin MSFlexGridLib.MSFlexGrid flxOptionen1 
            Height          =   2700
            Index           =   0
            Left            =   840
            TabIndex        =   12
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
      End
      Begin VB.Frame fmeOptionen 
         Height          =   4215
         Index           =   0
         Left            =   720
         TabIndex        =   10
         Top             =   480
         Width           =   8895
         Begin VB.TextBox txtOptionen0 
            Height          =   495
            Index           =   1
            Left            =   5280
            MaxLength       =   38
            TabIndex        =   3
            Text            =   "WWW9999"
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CheckBox chkOptionen0 
            Caption         =   "&Rezepturen mit Kassenrabatt taxieren"
            Height          =   495
            Left            =   120
            TabIndex        =   6
            Top             =   2280
            Width           =   6015
         End
         Begin VB.TextBox txtOptionen0 
            Height          =   495
            Index           =   2
            Left            =   5280
            TabIndex        =   5
            Text            =   "999,999"
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox txtOptionen0 
            Height          =   495
            Index           =   0
            Left            =   5280
            MaxLength       =   7
            TabIndex        =   1
            Text            =   "WWW9999"
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblOptionen0 
            Caption         =   "Rezept-&Text"
            Height          =   375
            Index           =   1
            Left            =   480
            TabIndex        =   2
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label lblOptionen0 
            Caption         =   "&Kassen-Rabatt  (%)"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   4
            Top             =   1920
            Width           =   3615
         End
         Begin VB.Label lblOptionen0 
            Caption         =   "&Instituts-Kennzeichen"
            Height          =   375
            Index           =   0
            Left            =   480
            TabIndex        =   0
            Top             =   480
            Width           =   2175
         End
      End
   End
End
Attribute VB_Name = "frmPVSOptionen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OrgRezApoDruckName$


Private Const DefErrModul = "REZKOPTIONEN.FRM"

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

Private Sub flxOptionen1_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxOptionen1_GotFocus")
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

With flxOptionen1(index)
    If (index = 0) Then
        .col = 0
        .ColSel = .Cols - 1
    End If
    .HighLight = flexHighlightAlways
End With

Call DefErrPop
End Sub

Private Sub flxOptionen1_LostFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxOptionen1_LostFocus")
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

With flxOptionen1(index)
    .HighLight = flexHighlightNever
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
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%
Dim h$, h2$, h3$, FormStr$
Dim c As Object


Call wpara.InitFont(Me)

txtOptionen0(1).Text = String(38, "A")

On Error Resume Next
For Each c In Controls
    If (TypeOf c Is TextBox) Then
        c.Width = TextWidth(c.Text) + 90
        c.Text = ""
    End If
Next
On Error GoTo DefErr

'''''''''''''''''''''''''''''''''''
tabOptionen.Tab = 0

lblOptionen0(0).Left = wpara.LinksX
lblOptionen0(0).Top = 2 * wpara.TitelY
txtOptionen0(0).Left = lblOptionen0(1).Left + lblOptionen0(1).Width + 300
txtOptionen0(0).Top = lblOptionen0(0).Top + (lblOptionen0(0).Height - txtOptionen0(0).Height) / 2

For i% = 1 To 2
    lblOptionen0(i%).Left = lblOptionen0(0).Left
    lblOptionen0(i%).Top = lblOptionen0(i% - 1).Top + lblOptionen0(i% - 1).Height + 300
    txtOptionen0(i%).Left = txtOptionen0(0).Left
    txtOptionen0(i%).Top = lblOptionen0(i%).Top + (lblOptionen0(i%).Height - txtOptionen0(i%).Height) / 2
Next i%

chkOptionen0.Left = lblOptionen0(0).Left
chkOptionen0.Top = lblOptionen0(2).Top + lblOptionen0(2).Height + 600

txtOptionen0(0).Text = OrgRezApoNr$
OrgRezApoDruckName$ = RezApoDruckName$
txtOptionen0(1).Text = Left$(RezApoDruckName$ + Space$(38), 38)
'txtOptionen0(2).text = Format(VmRabattFaktor#, "0.0000")
txtOptionen0(2).Text = Format(100# - ((1# / VmRabattFaktor#) * 100#), "0.00")

chkOptionen0.Value = Abs(RezepturMitFaktor%)

'''''''''''''''''''''''''''''''''''
tabOptionen.Tab = 1
Call ActProgram.FlxOptionenBefuellen
With flxOptionen1(0)
    Breite1% = 0
    For i% = 0 To (.Rows - 1)
        Breite2% = TextWidth(.TextMatrix(i%, 0))
        If (Breite2% > Breite1%) Then Breite1% = Breite2%
    Next i%
    .ColWidth(0) = Breite1% + 150
    .ColWidth(1) = TextWidth("00000")
    .ColWidth(2) = wpara.FrmScrollHeight

    .Width = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + 90
    .Height = .RowHeight(0) * 11 + 90
    
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
    
    .SelectionMode = flexSelectionByRow
    .col = 0
    .ColSel = .Cols - 1
End With

'''''''''''''''''''''''''''''''''''
tabOptionen.Tab = 2
With flxOptionen1(1)
    .Cols = 3
    .Rows = 2
    .FixedRows = 1
    .Rows = 1
    
    .FormatString = "^Tätigkeit|^Personal|^ "
   
    .ColWidth(0) = TextWidth(String(20, "A")) + 150
    .ColWidth(1) = TextWidth(String(30, "A"))
    .ColWidth(2) = 0

    .Width = .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + 90
    .Height = .RowHeight(0) * 11 + 90
    
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
    
    .SelectionMode = flexSelectionFree
    
    h$ = "Rezeptspeicher" + vbTab
    
    For i% = 0 To (AnzTaetigkeiten% - 1)
'        h$ = RTrim$(Taetigkeiten(i%).Taetigkeit)
'        h$ = h$ + vbTab
        h2$ = ""
        For k% = 0 To 49
            If (Taetigkeiten(i%).pers(k%) > 0) Then
                h2$ = h2$ + Mid$(Str$(Taetigkeiten(i%).pers(k%)), 2) + ","
            Else
                Exit For
            End If
        Next k%
        If (k% = 1) Then
            ind% = Taetigkeiten(i%).pers(0)
            h3$ = RTrim$(para.Personal(ind%))
        ElseIf (k% > 1) Then
            h3$ = "mehrere (" + Mid$(Str$(k%), 2) + ")"
        End If
        h$ = h$ + h3$ + vbTab + h2$
    Next i%
    
    .AddItem h$
    .row = 1
End With
'''''''''''''''''''''''''''''''''''
tabOptionen.Tab = 3
With flxOptionen1(2)
    .Rows = 13
    .Cols = 2
    .FixedRows = 1
    .FixedCols = 1
    .TextMatrix(0, 0) = "Monat"
    .TextMatrix(0, 1) = "Datum"
    .ColWidth(0) = TextWidth(String$(10, "A"))
    .ColWidth(1) = TextWidth(String$(10, "0"))
    AbrechDatenRec.MoveFirst
    For i% = 1 To 12
        .TextMatrix(i%, 0) = Format(CDate("01." + CStr(i%) + ".2002"), "MMMM")
        .TextMatrix(i%, 1) = Left(AbrechDatenRec!datum, 2) + "." + Mid(AbrechDatenRec!datum, 3, 2) + "." + Mid(AbrechDatenRec!datum, 5, 2)
        AbrechDatenRec.MoveNext
    Next i%
    
    .Width = .ColWidth(0) + .ColWidth(1) + 90
    .Height = .RowHeight(0) * 13 + 90
    
    .Left = wpara.LinksX
    .Top = 2 * wpara.TitelY
    
    .SelectionMode = flexSelectionByRow
    .col = 0
    .ColSel = .Cols - 1
End With

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

With fmeOptionen(0)
    .Left = wpara.LinksX
    .Top = 3 * wpara.TitelY
'    .Width = txtOptionen0(1).Left + txtOptionen0(1).Width + 900
    .Width = flxOptionen1(1).Left + flxOptionen1(1).Width + 900
    .Height = flxOptionen1(2).Top + flxOptionen1(2).Height + 300
End With
For i% = 1 To 3
    With fmeOptionen(i%)
        .Left = fmeOptionen(0).Left
        .Top = fmeOptionen(0).Top
        .Width = fmeOptionen(0).Width
        .Height = fmeOptionen(0).Height
    End With
Next i%

With tabOptionen
    .Left = wpara.LinksX
    .Top = wpara.TitelY
    .Width = fmeOptionen(0).Left + fmeOptionen(0).Width + wpara.LinksX
    .Height = fmeOptionen(0).Top + fmeOptionen(0).Height + wpara.TitelY
End With



cmdOk.Top = tabOptionen.Top + tabOptionen.Height + 150
cmdEsc.Top = cmdOk.Top

Me.Width = tabOptionen.Width + 2 * wpara.LinksX

cmdOk.Width = wpara.ButtonX
cmdOk.Height = wpara.ButtonY
cmdEsc.Width = wpara.ButtonX
cmdEsc.Height = wpara.ButtonY
cmdOk.Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
cmdEsc.Left = cmdOk.Left + cmdEsc.Width + 300

Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

Breite1% = frmAction.Left + (frmAction.Width - Me.Width) / 2
If (Breite1% < 0) Then Breite1% = 0
Me.Left = Breite1%
Hoehe1% = frmAction.Top + (frmAction.Height - Me.Height) / 2
If (Hoehe1% < 0) Then Hoehe1% = 0
Me.Top = Hoehe1%

tabOptionen.Tab = 0
Call TabDisable
Call TabEnable(tabOptionen.Tab)

Call DefErrPop
End Sub

Private Sub tabOptionen_Click(PreviousTab As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("tabOptionen_Click")
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
If (tabOptionen.Visible = False) Then Call DefErrPop: Exit Sub

Call TabDisable
Call TabEnable(tabOptionen.Tab)

Call DefErrPop
End Sub

Private Sub flxOptionen1_DblClick(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxOptionen1_DblClick")
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
Call cmdOk_Click
Call DefErrPop
End Sub

Sub TabDisable()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TabDisable")
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

For i% = 0 To 3
    fmeOptionen(i%).Visible = False
Next i%

Call DefErrPop
End Sub

Sub TabEnable(hTab%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("TabEnable")
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

fmeOptionen(hTab%).Visible = True

If (hTab% = 0) Then
    If (txtOptionen0(0).Visible) Then txtOptionen0(0).SetFocus
ElseIf (hTab% = 1) Then
    flxOptionen1(0).SetFocus
ElseIf (hTab% = 2) Then
    flxOptionen1(1).col = 1
    flxOptionen1(1).SetFocus
ElseIf (hTab% = 3) Then
    flxOptionen1(2).col = 1
    flxOptionen1(2).SetFocus
End If

Call DefErrPop
End Sub

Private Sub txtOptionen0_GotFocus(index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtOptionen0_GotFocus")
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
Dim h$

With txtOptionen0(index)
    h$ = .Text
    For i% = 1 To Len(h$)
        If (Mid$(h$, i%, 1) = ",") Then Mid$(h$, i%, 1) = "."
    Next i%
    .Text = h$
    .SelStart = 0
    .SelLength = Len(.Text)
End With

Call DefErrPop
End Sub

Private Sub txtOptionen0_KeyPress(index As Integer, KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtOptionen0_KeyPress")
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

If (index <> 1) Then
    If (Chr$(KeyAscii) = ",") Then KeyAscii = Asc(".")
    If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) And ((index <> 2) Or (Chr$(KeyAscii) <> ".")) Then
        Beep
        KeyAscii = 0
    End If
End If

Call DefErrPop
End Sub

Sub EditOptionenLstMulti()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EditOptionenLstMulti")
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
Dim i%, j%, l%, hTab%, row%, col%, ind%, aRow%
Dim s$, h$, BetrLief$, Lief2$

hTab% = tabOptionen.Tab
row% = 1
col% = 1
                
With flxOptionen1(1)
    aRow% = .row
    .row = 0
    .CellFontBold = True
    .row = aRow%
End With
            
With frmEdit.lstMultiEdit
    .Clear
    .AddItem "(keiner)"
    For i% = 1 To 50
        h$ = para.Personal(i%)
        .AddItem h$
    Next i%

    For i% = 0 To (.ListCount - 1)
        .Selected(i%) = False
    Next i%

    
    Load frmEdit
    
     .ListIndex = 0
     
     BetrLief$ = LTrim$(RTrim$(flxOptionen1(1).TextMatrix(row%, 2)))
     
     For i% = 0 To 19
         If (BetrLief$ = "") Then Exit For
         
         ind% = InStr(BetrLief$, ",")
         If (ind% > 0) Then
             Lief2$ = RTrim$(Left$(BetrLief$, ind% - 1))
             BetrLief$ = LTrim$(Mid$(BetrLief$, ind% + 1))
         Else
             Lief2$ = BetrLief$
             BetrLief$ = ""
         End If
         
         If (Lief2$ <> "") Then
            ind% = Val(Lief2$)
            .Selected(ind%) = True
         End If
     Next i%
     
    With frmEdit
        .Left = tabOptionen.Left + fmeOptionen(2).Left + flxOptionen1(1).Left + flxOptionen1(1).ColPos(col%) + 45
        .Left = .Left + Me.Left + wpara.FrmBorderHeight
        .Top = tabOptionen.Top + fmeOptionen(2).Top + flxOptionen1(1).Top + flxOptionen1(1).RowHeight(0)
        .Top = .Top + Me.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
        .Width = flxOptionen1(1).ColWidth(col%)
        .Height = flxOptionen1(1).Height - flxOptionen1(1).RowHeight(0)
    End With
    With frmEdit.lstMultiEdit
        .Height = frmEdit.ScaleHeight
        frmEdit.Height = .Height
        .Width = frmEdit.ScaleWidth
        .Left = 0
        .Top = 0
        
        .Visible = True
    End With


    frmEdit.Show 1

    With flxOptionen1(1)
        aRow% = .row
        .row = 0
        .CellFontBold = False
        .row = aRow%
    End With
            

    If (EditErg%) Then
            
        flxOptionen1(1).TextMatrix(row%, col% + 1) = EditTxt$
        
        h$ = ""
        If (EditAnzGefunden% = 0) Then
            h$ = ""
        ElseIf (EditAnzGefunden% = 1) Then
            ind% = EditGef%(0)
            h$ = RTrim$(para.Personal(ind%))
        Else
            h$ = "mehrere (" + Mid$(Str$(EditAnzGefunden%), 2) + ")"
        End If

        With flxOptionen1(1)
            .TextMatrix(row%, col%) = h$
            If (.col < .Cols - 2) Then .col = .col + 1
            If (.col < .Cols - 2) And (.ColWidth(.col) = 0) Then .col = .col + 1
        End With
        
    End If

End With

Call DefErrPop
End Sub
                            
Sub AuslesenFlexTaetigkeiten()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuslesenFlexTaetigkeiten")
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
Dim i%, j%, k%, ind%
Dim l&
Dim h$, BetrLief$, Lief2$, Key$

With flxOptionen1(1)
    j% = 0
    For i% = 1 To (.Rows - 1)
        h$ = RTrim$(.TextMatrix(i%, 0))
        If (h$ <> "") And (RTrim$(.TextMatrix(i%, 1)) <> "") Then
            Taetigkeiten(j%).Taetigkeit = h$
            BetrLief$ = LTrim$(RTrim$(.TextMatrix(i%, 2)))
            For k% = 0 To 49
                If (BetrLief$ = "") Then Exit For
                
                ind% = InStr(BetrLief$, ",")
                If (ind% > 0) Then
                    Lief2$ = RTrim$(Left$(BetrLief$, ind% - 1))
                    BetrLief$ = LTrim$(Mid$(BetrLief$, ind% + 1))
                Else
                    Lief2$ = BetrLief$
                    BetrLief$ = ""
                End If
                If (Lief2$ <> "") Then
                    Taetigkeiten(j%).pers(k%) = Val(Lief2$)
                End If
            Next k%
            Do
                If (k% > 49) Then Exit Do
                Taetigkeiten(j%).pers(k%) = 0
                k% = k% + 1
            Loop
            j% = j% + 1
        End If
    Next i%
End With
AnzTaetigkeiten = j%

i% = 1
If (i% <= AnzTaetigkeiten) Then
    h$ = RTrim$(Taetigkeiten(i% - 1).Taetigkeit)
    For j% = 0 To 49
        If (Taetigkeiten(i% - 1).pers(j%) > 0) Then
            h$ = h$ + "," + Mid$(Str$(Taetigkeiten(i% - 1).pers(j%)), 2)
        Else
            Exit For
        End If
    Next j%
Else
    h$ = ""
End If

Key$ = "Taetigkeit" + Format(i%, "00")
l& = WritePrivateProfileString("Rezeptkontrolle", Key$, h$, INI_DATEI)

Call DefErrPop
End Sub

