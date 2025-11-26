VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAngebote 
   Caption         =   "GH-Angebote"
   ClientHeight    =   4380
   ClientLeft      =   3645
   ClientTop       =   2205
   ClientWidth     =   6075
   Icon            =   "frmAngebote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6075
   Begin VB.CommandButton cmdF5 
      Caption         =   "Lösen (F5)"
      Height          =   450
      Left            =   960
      TabIndex        =   1
      Top             =   3480
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Caption         =   "ESC"
      Height          =   450
      Left            =   3600
      TabIndex        =   3
      Top             =   3480
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Zuordnen"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   3480
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid flxAngebote 
      Height          =   2280
      Left            =   840
      TabIndex        =   0
      Top             =   960
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
      ScrollBars      =   1
      SelectionMode   =   1
   End
   Begin VB.Label lblAngeboteWert 
      Caption         =   "999999.99"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   9
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblAngebote 
      Caption         =   "TaxeAEP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   8
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblAngeboteWert 
      Caption         =   "999999.99"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblAngebote 
      Caption         =   "AEP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblAngeboteWert 
      Caption         =   "999999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblAngebote 
      Caption         =   "BMopt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmAngebote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEsc_Click()

AngebotInd% = -1
Unload Me

End Sub

Private Sub cmdF5_Click()

AngebotInd% = 0
Unload Me

End Sub

Private Sub cmdOk_Click()

AngebotInd% = flxAngebote.col
Unload Me

End Sub

Private Sub flxAngebote_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i%, j%

If (KeyCode = vbKeyF5) Then
    cmdF5.Value = True
End If

End Sub

Private Sub Form_Load()
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%, maxSp%
Dim MenuHeight&, ScrollHeight&
Dim h$, h2$, FormStr$

Call wPara1.InitFont(Me)

Font.Bold = False   ' True

For i% = 0 To 2
    lblAngebote(i%).Top = wPara1.TitelY
    lblAngeboteWert(i%).Caption = ""
    lblAngeboteWert(i%).Top = wPara1.TitelY
Next i%

lblAngebote(0).Left = wPara1.LinksX
lblAngeboteWert(0).Left = lblAngebote(0).Left + lblAngebote(0).Width + 150
lblAngebote(1).Left = lblAngeboteWert(0).Left + lblAngeboteWert(0).Width + 300
lblAngeboteWert(1).Left = lblAngebote(1).Left + lblAngebote(1).Width + 150
lblAngebote(2).Left = lblAngeboteWert(1).Left + lblAngeboteWert(1).Width + 300
lblAngeboteWert(2).Left = lblAngebote(2).Left + lblAngebote(2).Width + 150



With flxAngebote
    .Cols = 2
    .Rows = 12
    .FixedRows = 1
    .FixedCols = 1
    
    .Top = lblAngebote(0).Top + lblAngebote(0).Height + 300
    .Left = wPara1.LinksX
    .Height = .RowHeight(0) * 13 + 90
    
    FormStr$ = "|^1|^2|^3|^4|^5|^6|^7|^8|^9|^10|^11|^12|^13|^14|^15;"
    FormStr$ = FormStr$ + "|Lieferant|Angebot|AEP|Rabatt|Preis|NR|aliLager|aliBest|Staffel|kalkAEP|Gespart|%"
    .FormatString = FormStr$
    .SelectionMode = flexSelectionByColumn
    
    .FillStyle = flexFillRepeat
    .row = 10
    .col = 0
    .RowSel = .Rows - 1
    .ColSel = .Cols - 1
    .CellBackColor = vbWhite
    .FillStyle = flexFillSingle
    
    Call AngeboteBefuellen


    For i% = 0 To .Cols - 1
        .ColWidth(i%) = TextWidth("WWWWWWWW")
        .ColAlignment(i%) = flexAlignCenterCenter
    Next i%

'    maxSp% = (frmAction.ScaleWidth - (2 * wPara1.LinksX) - 900) \ .ColWidth(0)
    maxSp% = (Screen.Width - (2 * wPara1.LinksX) - 900) \ .ColWidth(0)
    If (.Cols <= maxSp%) Then
        maxSp% = .Cols
    Else
'        .Height = .Height + .RowHeight(0)
        .Height = .Height + wPara1.FrmScrollHeight
    End If
    
    spBreite% = 0
    For i% = 0 To maxSp% - 1
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 180    '90
    
    Do
        If (.col >= (.LeftCol + maxSp% - 1)) Then
            .LeftCol = .LeftCol + 1
        Else
            Exit Do
        End If
    Loop
End With

'

Font.Name = wPara1.FontName(1)
Font.Size = wPara1.FontSize(1)


cmdOk.Width = wPara1.ButtonX
cmdOk.Height = wPara1.ButtonY
cmdEsc.Width = cmdOk.Width
cmdEsc.Height = cmdOk.Height
cmdF5.Width = TextWidth(cmdF5.Caption) + 150
cmdF5.Height = cmdOk.Height

cmdOk.Top = flxAngebote.Top + flxAngebote.Height + 150
cmdEsc.Top = cmdOk.Top
cmdF5.Top = cmdOk.Top


Breite1% = flxAngebote.Width + 2 * wPara1.LinksX
If (AngebotModus% = 0) Then
    Breite2% = 0
Else
    Breite2% = cmdF5.Width + 900 + cmdOk.Width + 300 + cmdEsc.Width + 2 * wPara1.LinksX
End If

If (Breite2% > Breite1%) Then
    Me.Width = Breite2%
Else
    Me.Width = Breite1%
End If


If (AngebotModus% = 0) Then
    cmdOk.Caption = "OK"
    cmdOk.Default = True
    cmdOk.Cancel = True
    cmdOk.Visible = True
    cmdEsc.Visible = False
    cmdF5.Visible = False
    
    cmdOk.Left = (Me.ScaleWidth - cmdOk.Width) / 2
Else
    cmdOk.Caption = "Binden"
    cmdOk.Default = True
    cmdOk.Cancel = False
    cmdOk.Visible = True
    cmdEsc.Cancel = True
    cmdEsc.Visible = True
    cmdF5.Visible = True
    
    cmdF5.Left = (Me.ScaleWidth - (cmdF5.Width + 900 + cmdOk.Width + 300 + cmdEsc.Width)) / 2
    cmdOk.Left = cmdF5.Left + cmdF5.Width + 900
    cmdEsc.Left = cmdOk.Left + cmdOk.Width + 300
End If


Me.Height = cmdOk.Top + cmdOk.Height + wPara1.TitelY + 90 + wPara1.FrmCaptionHeight

'Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
'Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2


End Sub

Sub AngeboteBefuellen()

Call Angebote

End Sub
    
Private Sub lblAbholerWert_Click(Index As Integer)

End Sub
