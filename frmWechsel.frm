VERSION 5.00
Begin VB.Form frmWechsel 
   Caption         =   "Wechsel"
   ClientHeight    =   1980
   ClientLeft      =   1785
   ClientTop       =   3195
   ClientWidth     =   2835
   Icon            =   "frmWechsel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   2835
   Begin VB.ListBox lstWechsel 
      Height          =   1635
      Left            =   0
      Style           =   1  'Kontrollkästchen
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   390
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   1200
   End
End
Attribute VB_Name = "frmWechsel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEsc_Click()
Dim i%

With lstWechsel
    WechselErg% = -1
    WechselStart% = 0
    For i% = 0 To .ListCount - 1
        If (.Selected(i%)) Then
            WechselStart% = i%
            Exit For
        End If
    Next i%
End With

Unload Me
End Sub

Private Sub cmdOk_Click()
Dim i%

With lstWechsel
    WechselErg% = .ListIndex
    WechselStart% = 0
    For i% = 0 To .ListCount - 1
        If (.Selected(i%)) Then
            WechselStart% = i%
            Exit For
        End If
    Next i%
End With

Unload Me
End Sub

Private Sub Form_Load()
Dim i%, MaxWi%, wi%, AnzRows%
Dim h$
                
Call wPara1.InitFont(Me)

With lstWechsel
    .Clear
    MaxWi% = 0
    AnzRows% = UBound(WechselZeilen$) + 1
    For i% = 0 To AnzRows% - 1
        h$ = WechselZeilen$(i%)
        .AddItem h$
        wi% = TextWidth(h$)
        If (wi% > MaxWi%) Then
            MaxWi% = wi%
        End If
    Next i%
    .Selected(WechselStart%) = True
    .ListIndex = WechselStart%
End With

Width = MaxWi% + 900
Height = AnzRows% * (TextHeight("Äg") * 1.3) + 90 + 2 * wPara1.FrmBorderHeight + wPara1.FrmCaptionHeight
'Left = frmMatchcode.Left + (frmMatchcode.Width - Width) / 2
'Top = frmMatchcode.Top + (frmMatchcode.Height - Height) / 2
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2
    
With lstWechsel
    .Height = ScaleHeight
    .Width = ScaleWidth
    .Left = 0
    .Top = 0
    Height = .Height + 2 * wPara1.FrmBorderHeight + wPara1.FrmCaptionHeight
End With

'Icon = frmMatchcode.imgToolbar(0).ListImages(3).ExtractIcon
        
'        .AddItem "(keiner)"
'        .AddItem String$(50, "-")
'        For i% = 1 To AnzLiefNamen%
'            h$ = LiefNamen$(i% - 1)
'            .AddItem h$
'            If (bek.lief > 0) And (lInd% < 0) Then
'                ind% = InStr(h$, "(")
'                h$ = Mid$(h$, ind% + 1)
'                If (bek.lief = Val(Left$(h$, Len(h$) - 1))) Then
'                    lInd% = i% + 1
'                End If
'            End If
'        Next i%
'        If (lInd% < 0) Then
'            .ListIndex = 0
'        Else
'            .ListIndex = lInd%
'        End If
'        .Visible = True
'    End With

End Sub

Private Sub lstWechsel_ItemCheck(Item As Integer)
Dim i%
Static Aktiv%

If (Aktiv%) Then Exit Sub

Aktiv% = True

With lstWechsel
    For i% = 0 To .ListCount - 1
        If (i% <> Item) Then
            .Selected(i%) = False
        End If
    Next i%
End With

Aktiv% = False

End Sub
