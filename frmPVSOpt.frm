VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptionen 
   Caption         =   "Optionen"
   ClientHeight    =   7395
   ClientLeft      =   -630
   ClientTop       =   240
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   9975
   Begin VB.CommandButton cmdF5 
      Caption         =   "Entfernen (F5)"
      Height          =   450
      Left            =   8160
      TabIndex        =   10
      Top             =   6000
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdF2 
      Caption         =   "Einfügen (F2)"
      Height          =   450
      Left            =   6600
      TabIndex        =   9
      Top             =   6000
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.ListBox lstOptionenMulti 
      Height          =   450
      Left            =   4680
      MultiSelect     =   1  '1 -Einfach
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtOptionen 
      Height          =   375
      Left            =   4920
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstOptionen 
      Height          =   450
      Left            =   7080
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   3360
      TabIndex        =   12
      Top             =   6240
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   1920
      TabIndex        =   11
      Top             =   6240
      Width           =   1200
   End
   Begin TabDlg.SSTab tabOptionen 
      Height          =   5325
      Left            =   0
      TabIndex        =   13
      Top             =   600
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   9393
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
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
      TabCaption(0)   =   "&1 - Kontrollen"
      TabPicture(0)   =   "frmPVSOpt.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "flxOptionen(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&2 - Zuordnungen"
      TabPicture(1)   =   "frmPVSOpt.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "flxOptionen(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&3 - Rufzeiten"
      TabPicture(2)   =   "frmPVSOpt.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "flxOptionen(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&4 - Automatik"
      TabPicture(3)   =   "frmPVSOpt.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lblOptionenAutomatik(1)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblOptionenAutomatik(0)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lblOptionenAutomatikMinuten(0)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "lblOptionenAutomatikMinuten(1)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "fmeOptionenAutomatikBestVors"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtOptionenAutomatik(1)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "txtOptionenAutomatik(0)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).ControlCount=   7
      Begin VB.TextBox txtOptionenAutomatik 
         Alignment       =   2  'Zentriert
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   7920
         TabIndex        =   3
         ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtOptionenAutomatik 
         Alignment       =   2  'Zentriert
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   7920
         TabIndex        =   4
         ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
         Top             =   1800
         Width           =   495
      End
      Begin VB.Frame fmeOptionenAutomatikBestVors 
         Caption         =   "Bestell&vorschlag"
         Height          =   2055
         Left            =   480
         TabIndex        =   14
         Top             =   2640
         Width           =   9015
         Begin VB.CheckBox chkOptionenAutomatikBestVors 
            Caption         =   "&Komplett vor jedem eingetragenen Sendeauftrag :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   5
            Top             =   480
            Width           =   5775
         End
         Begin VB.TextBox txtOptionenAutomatikBestVors 
            Alignment       =   2  'Zentriert
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   6600
            TabIndex        =   6
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   480
            Width           =   495
         End
         Begin VB.CheckBox chkOptionenAutomatikBestVors 
            Caption         =   "&Periodisch im Hintergrund, ein Durchlauf dauert"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   360
            TabIndex        =   7
            Top             =   1080
            Width           =   5535
         End
         Begin VB.TextBox txtOptionenAutomatikBestVors 
            Alignment       =   2  'Zentriert
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   6600
            TabIndex        =   8
            ToolTipText     =   "Obergrenze, bis zu der eine Liste automatisch aufgeblendet wird"
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label lblOptionenAutomatikBestVorsMinuten 
            Caption         =   "Minuten davor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   7200
            TabIndex        =   16
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblOptionenAutomatikBestVorsMinuten 
            Caption         =   "Minuten "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   7200
            TabIndex        =   15
            Top             =   1080
            Width           =   1695
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flxOptionen 
         Height          =   3600
         Index           =   0
         Left            =   -74640
         TabIndex        =   0
         Top             =   1260
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   6350
         _Version        =   393216
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
      Begin MSFlexGridLib.MSFlexGrid flxOptionen 
         Height          =   3600
         Index           =   1
         Left            =   -74760
         TabIndex        =   1
         Top             =   1020
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   6350
         _Version        =   393216
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
      Begin MSFlexGridLib.MSFlexGrid flxOptionen 
         Height          =   3600
         Index           =   2
         Left            =   -74760
         TabIndex        =   2
         Top             =   1140
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   6350
         _Version        =   393216
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
      Begin VB.Label lblOptionenAutomatikMinuten 
         Caption         =   "Minuten "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   8400
         TabIndex        =   20
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lblOptionenAutomatikMinuten 
         Caption         =   "Minuten davor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   8400
         TabIndex        =   19
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblOptionenAutomatik 
         Caption         =   "Wieviele Minuten vor dem &Senden soll der Hinweis erscheinen :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   18
         Top             =   1200
         Width           =   6855
      End
      Begin VB.Label lblOptionenAutomatik 
         Caption         =   "Wieviele Minuten soll auf &Anruf gewartet werden :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   17
         Top             =   1800
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmOptionen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "frmOptionen.frm"

Private Sub cmdEsc_Click()
Unload Me
End Sub

Private Sub cmdF2_Click()
Dim i%, j%
        
With flxOptionen(tabOptionen.Tab)
    For j% = (.Rows - 2) To .row Step -1
        For i% = 0 To .Cols - 1
            .TextMatrix(j% + 1, i%) = .TextMatrix(j%, i%)
        Next i%
    Next j%
    For i% = 0 To .Cols - 1
        .TextMatrix(.row, i%) = ""
    Next i%
End With

End Sub

Private Sub cmdF5_Click()
Dim i%

With flxOptionen(tabOptionen.Tab)
    For i% = 0 To .Cols - 1
        .TextMatrix(.row, i%) = ""
    Next i%
End With

End Sub

Private Sub cmdOk_Click()
Dim i%, j%, l%, hTab%, row%, col%, Anzgefunden%, ind%, Gef%(19)
Dim uhr%, st%, min%, MultiAuswahl%, LiefGef%, BmMulti%
Dim h$, h2$, s$, BetrLief$, Lief2$

If (ActiveControl.Name = cmdOk.Name) Then
    Call AuslesenFlexKontrollen
    Call SpeicherIniKontrollen
    Call SpeicherIniZuordnungen
    Call SpeicherIniRufzeiten
    Call frmAction.SpeicherIniWerte
    
    Call frmAction.AuslesenBestellung(True, False, True)
    Unload Me

Else
    hTab% = tabOptionen.Tab
    If (hTab% < 3) Then
        row% = flxOptionen(hTab%).row
        col% = flxOptionen(hTab%).col
        If (lstOptionen.Visible = True) Then
            h$ = lstOptionen.Text
            lstOptionen.Visible = False
            
            If ((h$ = "Eingabe") Or (h$ = "?-fach")) Then
                With txtOptionen
                    BmMulti% = False
                    If (h$ = "?-fach") Then BmMulti% = True
                    
                    .Top = tabOptionen.Top + flxOptionen(hTab%).Top + (row% - flxOptionen(hTab%).TopRow + 1) * flxOptionen(hTab%).RowHeight(0)
                    .Height = flxOptionen(hTab%).RowHeight(1)
                    .Left = tabOptionen.Left + flxOptionen(hTab%).Left + flxOptionen(hTab%).ColPos(col%) + 15
                    .Width = flxOptionen(hTab%).ColWidth(col%)
                    
                    h2$ = flxOptionen(hTab%).TextMatrix(row%, col%)
                    If (BmMulti%) Then
                        ind% = InStr(h2$, "*")
                        If (ind% > 0) Then
                            .Text = Left$(h2$, ind% - 1)
                        Else
                            .Text = h2$
                        End If
                    Else
                        ind% = InStr(h2$, "Tage")
                        If (ind% > 0) Then
                            .Text = Left$(h2$, ind% - 1)
                        Else
                            .Text = h2$
                        End If
                    End If
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .BackColor = vbRed
                    .Visible = True
                    .TabStop = True
                    .SetFocus
                End With
            End If
        ElseIf (txtOptionen.Visible = True) Then
            h$ = RTrim(txtOptionen.Text)
            txtOptionen.Visible = False
            tabOptionen.Enabled = True
            With flxOptionen(hTab%)
                If (hTab% < 2) Then
                    If (BmMulti% = True) Then
                        .TextMatrix(.row, .col) = h$ + "* BMopt"
                    Else
                        If (col% = 2) And (.TextMatrix(row%, 0) = "Besorger") Then h$ = h$ + " Tage"
                        .TextMatrix(.row, .col) = h$
                    End If
                Else
                    Select Case .col
                        Case 1
                            uhr% = Val(h$)
                            st% = uhr% \ 100
                            min% = uhr% Mod 100
                            If (st% >= 0) And (st% < 24) And (min% >= 0) And (min% < 60) Then
                                .TextMatrix(.row, .col) = Format(st%, "00") + ":" + Format(min%, "00")
                            End If
                        Case 4, 5
                            .TextMatrix(.row, .col) = UCase$(h$)
                    End Select
                End If
                .SetFocus
                If (.col < .Cols - 2) Then .col = .col + 1
                If (.col < .Cols - 2) And (.ColWidth(.col) = 0) Then .col = .col + 1
            End With
        Else
            s$ = RTrim$(flxOptionen(hTab%).TextMatrix(row%, 0))
            
            Select Case flxOptionen(hTab%).col
                Case 0
                    Call EditOptionenLst
                    Exit Sub
                Case 1
                    If (hTab% < 2) Then
                        Call EditOptionenLst
                    Else
                        Call EditOptionenTxt
                    End If
                Case 2
                    If (hTab% < 2) Then
                        If (ZeilenTyp%(s$) = 1) Then
                            If (UCase(s$) = "LIEFERANT") Then
                                Call EditOptionenLst
                            Else
                                Call EditOptionenTxt
                            End If
                        Else
                            Call EditOptionenTxt
                        End If
                    Else
                        Call EditOptionenTxt
                    End If
                Case 3
                    If (hTab% = 0) Then
                        Call EditOptionenLst
                    Else
                        Call EditOptionenLstMulti
                    End If
                Case 5, 6
                    Call EditOptionenLst
                Case 7
                    If (hTab% = 2) Then
                        Call EditOptionenLst
                    End If
            End Select
        End If
    End If
End If

End Sub

Private Sub flxOptionen_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i%, j%

If (Index < 3) Then
    If (KeyCode = vbKeyF2) Then
        cmdF2.Value = True
    ElseIf (KeyCode = vbKeyF5) Then
        cmdF5.Value = True
    End If
End If

End Sub

Private Sub Form_Load()
Dim i%, j%, l%, k%, lInd%, wi%, MaxWi%, spBreite%, ind%
Dim Breite1%, Breite2%, Hoehe1%, Hoehe2%
Dim h$, h2$, FormStr$


Call wpara.InitFont(Me)

tabOptionen.Left = wpara.LinksX
tabOptionen.Top = wpara.TitelY

For j% = 0 To 2
    With flxOptionen(j%)
        If (j% = 2) Then
            .Cols = 9
        Else
            .Cols = 5
        End If
        
        .Rows = 2
        .FixedRows = 1
        
        .Top = 900   'TitelY%
        .Left = wpara.LinksX
        .Height = .RowHeight(0) * 11 + 90
        
        If (j% = 2) Then
            FormStr$ = "^Lieferant|^Rufzeit|^Liefzeit|^Wochentag(e)||^Erg|^Art|<Aktiv|"
        Else
            FormStr$ = "^Wert1|^Bedingung|^Wert2|"
            If (j% = 0) Then
                FormStr$ = FormStr$ + "<Unkontrollierte"
            Else
                FormStr$ = FormStr$ + "<Lieferant(en)"
            End If
            FormStr$ = FormStr$ + "|^ "
        End If
        
        .FormatString = FormStr$
        .Rows = 1
        .SelectionMode = flexSelectionFree
    End With
Next j%
        


Font.Bold = False   ' True

tabOptionen.Tab = 2
With flxOptionen(2)
    .ColWidth(0) = TextWidth("WWWWWW (999)")
    .ColWidth(1) = TextWidth("WWWWWW")
    .ColWidth(2) = TextWidth("WWWWWW")
'    .ColWidth(2) = TextWidth(String$(15, "W"))
    .ColWidth(3) = TextWidth("Wwwwwwwww, Wwwwwwwww")
    .ColWidth(4) = 0
    .ColWidth(5) = TextWidth("WWW")
    .ColWidth(6) = TextWidth("WWW")
    .ColWidth(7) = TextWidth("WWWWW")
    .ColWidth(8) = 0
    
    spBreite% = 0
    For i% = 0 To .Cols - 1
        spBreite% = spBreite% + .ColWidth(i%)
    Next i%
    .Width = spBreite% + 90
    .Rows = 1
End With

For j% = 0 To 1
    tabOptionen.Tab = j%
    With flxOptionen(j%)
        .Width = flxOptionen(2).Width
        .ColWidth(0) = 0
        .ColWidth(4) = 0
        For i% = 1 To (.Cols - 2)
            .ColWidth(i%) = .Width / 4
        Next i%

        spBreite% = 0
        For i% = 1 To .Cols - 2
            spBreite% = spBreite% + .ColWidth(i%)
        Next i%
        .ColWidth(0) = .Width - spBreite%
    End With
Next j%




        
Call OptionenBefuellen
'

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

cmdF2.Width = TextWidth(cmdF5.Caption) + 150
cmdF2.Height = wpara.ButtonY
cmdF2.Left = tabOptionen.Left + flxOptionen(0).Left + flxOptionen(0).Width + 150
cmdF2.Top = tabOptionen.Top + flxOptionen(0).Top

cmdF5.Width = cmdF2.Width
cmdF5.Height = cmdF2.Height
cmdF5.Left = cmdF2.Left
cmdF5.Top = cmdF2.Top + cmdF2.Height + 90


Hoehe1% = flxOptionen(2).Top + flxOptionen(2).Height + 180
Breite1% = cmdF2.Left + cmdF2.Width + wpara.LinksX - tabOptionen.Left

'------------------------



tabOptionen.Tab = 3


txtOptionenAutomatik(0).Top = 900   'TitelY%
For i% = 1 To 1
    txtOptionenAutomatik(i%).Top = txtOptionenAutomatik(i% - 1).Top + txtOptionenAutomatik(i% - 1).Height + 90
Next i%

lblOptionenAutomatik(0).Left = wpara.LinksX
lblOptionenAutomatik(0).Top = txtOptionenAutomatik(0).Top
For i% = 1 To 1
    lblOptionenAutomatik(i%).Left = lblOptionenAutomatik(i% - 1).Left
    lblOptionenAutomatik(i%).Top = txtOptionenAutomatik(i%).Top
Next i%

MaxWi% = 0
For i% = 0 To 1
    wi% = lblOptionenAutomatik(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

txtOptionenAutomatik(0).Left = lblOptionenAutomatik(0).Left + MaxWi% + 300
For i% = 1 To 1
    txtOptionenAutomatik(i%).Left = txtOptionenAutomatik(i% - 1).Left
Next i%

lblOptionenAutomatikMinuten(0).Left = txtOptionenAutomatik(0).Left + txtOptionenAutomatik(0).Width + 150
lblOptionenAutomatikMinuten(0).Top = txtOptionenAutomatik(0).Top
For i% = 1 To 1
    lblOptionenAutomatikMinuten(i%).Left = lblOptionenAutomatikMinuten(i% - 1).Left
    lblOptionenAutomatikMinuten(i%).Top = txtOptionenAutomatik(i%).Top
Next i%



fmeOptionenAutomatikBestVors.Left = lblOptionenAutomatik(0).Left
fmeOptionenAutomatikBestVors.Top = lblOptionenAutomatik(1).Top + lblOptionenAutomatik(1).Height + 450

txtOptionenAutomatikBestVors(0).Top = 2 * wpara.TitelY
For i% = 1 To 1
    txtOptionenAutomatikBestVors(i%).Top = txtOptionenAutomatikBestVors(i% - 1).Top + txtOptionenAutomatikBestVors(i% - 1).Height + 90
Next i%

chkOptionenAutomatikBestVors(0).Left = wpara.LinksX
chkOptionenAutomatikBestVors(0).Top = txtOptionenAutomatikBestVors(0).Top
For i% = 1 To 1
    chkOptionenAutomatikBestVors(i%).Left = chkOptionenAutomatikBestVors(i% - 1).Left
    chkOptionenAutomatikBestVors(i%).Top = txtOptionenAutomatikBestVors(i%).Top
Next i%

MaxWi% = 0
For i% = 0 To 1
    wi% = chkOptionenAutomatikBestVors(i%).Width
    If (wi% > MaxWi%) Then
        MaxWi% = wi%
    End If
Next i%

txtOptionenAutomatikBestVors(0).Left = txtOptionenAutomatik(0).Left - fmeOptionenAutomatikBestVors.Left
For i% = 1 To 1
    txtOptionenAutomatikBestVors(i%).Left = txtOptionenAutomatikBestVors(i% - 1).Left
Next i%

lblOptionenAutomatikBestVorsMinuten(0).Left = txtOptionenAutomatikBestVors(0).Left + txtOptionenAutomatikBestVors(0).Width + 150
lblOptionenAutomatikBestVorsMinuten(0).Top = txtOptionenAutomatikBestVors(0).Top
For i% = 1 To 1
    lblOptionenAutomatikBestVorsMinuten(i%).Left = lblOptionenAutomatikBestVorsMinuten(i% - 1).Left
    lblOptionenAutomatikBestVorsMinuten(i%).Top = txtOptionenAutomatikBestVors(i%).Top
Next i%

fmeOptionenAutomatikBestVors.Width = lblOptionenAutomatikBestVorsMinuten(0).Left + lblOptionenAutomatikBestVorsMinuten(0).Width + 2 * wpara.LinksX
fmeOptionenAutomatikBestVors.Height = txtOptionenAutomatikBestVors(1).Top + txtOptionenAutomatikBestVors(1).Height + 2 * wpara.TitelY

Hoehe2% = fmeOptionenAutomatikBestVors.Top + fmeOptionenAutomatikBestVors.Height + 180
Breite2% = fmeOptionenAutomatikBestVors.Width + 2 * wpara.LinksX

If (Hoehe1% > Hoehe2%) Then
    tabOptionen.Height = Hoehe1%
Else
    tabOptionen.Height = Hoehe2%
End If
If (Breite1% > Breite2%) Then
    tabOptionen.Width = Breite1%
Else
    tabOptionen.Width = Breite2%
End If


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

Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2


tabOptionen.Tab = 0
Call TabDisable
flxOptionen(0).Visible = True
cmdF2.Visible = True
cmdF5.Visible = True

End Sub

Private Sub tabOptionen_Click(PreviousTab As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("tabOptionen_Click")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
If (tabOptionen.Visible = False) Then Call DefErrPop: Exit Sub

Call TabDisable
Call TabEnable(tabOptionen.Tab)

'Select Case tabOptionen.Tab
'    Case 0
'        flxOptionen(0).SetFocus
'    Case 1
'        flxOptionen(1).SetFocus
'    Case 2
'        flxOptionen(2).SetFocus
'    Case 3
'        txtOptionenAutomatik(0).SetFocus
'End Select

Call DefErrPop
End Sub

Sub OptionenBefuellen()
Dim i%, j%, k%, l%, ind%
Dim h$, h2$, h3$

lstOptionen.Visible = False
lstOptionenMulti.Visible = False
txtOptionen.Visible = False

For j% = 0 To 2
    flxOptionen(j%).Rows = 1
    If (j% = 0) Then
        For i% = 0 To (AnzKontrollen% - 1)
            h$ = RTrim$(Kontrollen(i%).bedingung.wert1)
            h$ = h$ + vbTab + RTrim$(Kontrollen(i%).bedingung.op)
            h$ = h$ + vbTab + RTrim$(Kontrollen(i%).bedingung.wert2)
            If (Kontrollen(i%).Send = "J") Then
                h$ = h$ + vbTab + "Senden"
            Else
                h$ = h$ + vbTab + "Nicht Senden"
            End If
            flxOptionen(j%).AddItem h$
        Next i%
        flxOptionen(j%).Rows = MAX_KONTROLLEN
    ElseIf (j% = 1) Then
        For i% = 0 To (AnzZuordnungen% - 1)
            h$ = RTrim$(Zuordnungen(i%).bedingung.wert1)
            h$ = h$ + vbTab + RTrim$(Zuordnungen(i%).bedingung.op)
            h$ = h$ + vbTab + RTrim$(Zuordnungen(i%).bedingung.wert2)

            h$ = h$ + vbTab
            h2$ = ""
            For k% = 0 To 19
                If (Zuordnungen(i%).lief(k%) > 0) Then
                    h2$ = h2$ + Mid$(Str$(Zuordnungen(i%).lief(k%)), 2) + ","
                Else
                    Exit For
                End If
            Next k%
            If (k% = 1) Then
                If (Zuordnungen(i%).lief(0) = 255) Then
                    h3$ = "Naechstliefernder"
                Else
                    lif.GetRecord (Zuordnungen(i%).lief(0) + 1)
                    h3$ = RTrim$(lif.kurz)
                    Call OemToChar(h3$, h3$)
                End If
                h$ = h$ + h3$
            ElseIf (k% > 1) Then
                h$ = h$ + "mehrere (" + Mid$(Str$(k%), 2) + ")"
            End If
            h$ = h$ + vbTab + h2$
            flxOptionen(j%).AddItem h$
        Next i%
        flxOptionen(j%).Rows = MAX_ZUORDNUNGEN
    Else
        For i% = 0 To (AnzRufzeiten% - 1)
            lif.GetRecord (Rufzeiten(i%).Lieferant + 1)
            h$ = RTrim$(lif.kurz)
            Call OemToChar(h$, h$)
            h$ = h$ + " (" + Mid$(Str$(Rufzeiten(i%).Lieferant), 2) + ")"

            h$ = h$ + vbTab + Format$(Rufzeiten(i%).RufZeit / 100, "00")
            h$ = h$ + ":" + Format$(Rufzeiten(i%).RufZeit Mod 100, "00")
            h$ = h$ + vbTab + Format$(Rufzeiten(i%).LieferZeit / 100, "00")
            h$ = h$ + ":" + Format$(Rufzeiten(i%).LieferZeit Mod 100, "00")

            h$ = h$ + vbTab
            h2$ = ""
            For k% = 0 To 6
                ind% = Rufzeiten(i%).WoTag(k%)
                If (ind% > 0) Then
                    h2$ = h2$ + Mid$(Str$(ind%), 2) + ","
                Else
                    Exit For
                End If
            Next k%
            If (k% = 0) Then
            ElseIf (k% = 1) Then
                ind% = Rufzeiten(i%).WoTag(0)
                h$ = h$ + WochenTag$(ind% - 1)
            ElseIf (k% < 4) Then
                For l% = 0 To (k% - 1)
                    ind% = Rufzeiten(i%).WoTag(l%)
                    h$ = h$ + WochenTag$(ind% - 1) + ","
                Next l%
            Else
                For l% = 0 To (k% - 1)
                    ind% = Rufzeiten(i%).WoTag(l%)
                    h$ = h$ + Left$(WochenTag$(ind% - 1), 2) + ","
                Next l%
            End If
            h$ = h$ + vbTab + h2$
            h$ = h$ + vbTab + Rufzeiten(i%).AuftragsErg
            h$ = h$ + vbTab + Rufzeiten(i%).AuftragsArt
            If (Rufzeiten(i%).Aktiv = "J") Then
                h$ = h$ + vbTab + "ja"
            Else
                h$ = h$ + vbTab + "nein"
            End If
            flxOptionen(j%).AddItem h$
        Next i%
        flxOptionen(j%).Rows = MAX_RUFZEITEN
    End If
    flxOptionen(j%).row = 1
    flxOptionen(j%).col = 0
Next j%

txtOptionenAutomatik(0).Text = Str$(AnzMinutenWarnung%)
txtOptionenAutomatik(1).Text = Str$(AnzMinutenWarten%)

chkOptionenAutomatikBestVors(0).Value = Abs(BestVorsKomplett%)
chkOptionenAutomatikBestVors(1).Value = Abs(BestVorsPeriodisch%)
txtOptionenAutomatikBestVors(0).Text = Str$(BestVorsKomplettMinuten%)
txtOptionenAutomatikBestVors(0).Visible = Abs(BestVorsKomplett%)
txtOptionenAutomatikBestVors(1).Text = Str$(BestVorsPeriodischMinuten%)
txtOptionenAutomatikBestVors(1).Visible = Abs(BestVorsPeriodisch%)
lblOptionenAutomatikBestVorsMinuten(0).Visible = Abs(BestVorsKomplett%)
lblOptionenAutomatikBestVorsMinuten(1).Visible = Abs(BestVorsPeriodisch%)

End Sub

Private Sub chkOptionenAutomatikBestVors_Click(Index As Integer)

If (Index = 0) Then
    BestVorsKomplett% = chkOptionenAutomatikBestVors(0).Value
Else
    BestVorsPeriodisch% = chkOptionenAutomatikBestVors(1).Value
End If
txtOptionenAutomatikBestVors(0).Visible = Abs(BestVorsKomplett%)
txtOptionenAutomatikBestVors(1).Visible = Abs(BestVorsPeriodisch%)
lblOptionenAutomatikBestVorsMinuten(0).Visible = Abs(BestVorsKomplett%)
lblOptionenAutomatikBestVorsMinuten(1).Visible = Abs(BestVorsPeriodisch%)

End Sub

Private Sub lstOptionen_DblClick()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("lstOptionen_DblClick")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call cmdOk_Click
Call DefErrPop
End Sub

Private Sub flxOptionen_DblClick(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("flxOptionen_DblClick")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call cmdOk_Click
Call DefErrPop
End Sub

Private Sub txtOptionenAutomatik_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtOptionenAutomatik_GotFocus")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'If (tabOptionen.Tab <> 3) Then
'    cmdOk.SetFocus
'End If

With txtOptionenAutomatik(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With

Call DefErrPop

End Sub

Private Sub txtOptionenAutomatikbestvors_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtOptionenAutomatikBestVors_GotFocus")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'If (tabOptionen.Tab <> 3) Then
'    flxOptionen(tabOptionen.Tab).SetFocus
'End If

With txtOptionenAutomatikBestVors(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With

Call DefErrPop


End Sub

Private Sub AuslesenFlexKontrollen()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("AuslesenFlexKontrollen")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i%, j%, k%, ind%
Dim h$, Send$, BetrLief$, Lief2$, BetrTage$, tag$, Aktiv$

With flxOptionen(0)
    j% = 0
    For i% = 1 To (.Rows - 1)
        h$ = RTrim$(.TextMatrix(i%, 0))
        If (h$ <> "") Then
            Kontrollen(j%).bedingung.wert1 = h$
            Kontrollen(j%).bedingung.op = RTrim$(.TextMatrix(i%, 1))
            
            h$ = RTrim$(.TextMatrix(i%, 2))
            Kontrollen(j%).bedingung.wert2 = h$
            
            Send$ = RTrim$(.TextMatrix(i%, 3))
            If (Send$ = "Nicht Senden") Then
                Kontrollen(j%).Send = "N"
            Else
                Kontrollen(j%).Send = "J"
            End If
            j% = j% + 1
        End If
    Next i%
End With
AnzKontrollen% = j%

With flxOptionen(1)
    j% = 0
    For i% = 1 To (.Rows - 1)
        h$ = RTrim$(.TextMatrix(i%, 0))
        If (h$ <> "") Then
            Zuordnungen(j%).bedingung.wert1 = h$
            Zuordnungen(j%).bedingung.op = RTrim$(.TextMatrix(i%, 1))
            Zuordnungen(j%).bedingung.wert2 = RTrim$(.TextMatrix(i%, 2))
            BetrLief$ = LTrim$(RTrim$(.TextMatrix(i%, 4)))
            For k% = 0 To 19
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
                    Zuordnungen(j%).lief(k%) = Val(Lief2$)
                End If
            Next k%
            Do
                If (k% > 19) Then Exit Do
                Zuordnungen(j%).lief(k%) = 0
                k% = k% + 1
            Loop
            j% = j% + 1
        End If
    Next i%
End With
AnzZuordnungen% = j%

With flxOptionen(2)
    j% = 0
    For i% = 1 To (.Rows - 1)
        h$ = LTrim$(RTrim$(.TextMatrix(i%, 0)))
        If (h$ <> "") Then
            ind% = InStr(h$, "(")
            If (ind% > 0) Then
                h$ = Mid$(h$, ind% + 1)
                ind% = InStr(h$, ")")
                h$ = Left$(h$, ind% - 1)
                Rufzeiten(j%).Lieferant = Val(h$)
            End If
            
            h$ = RTrim$(.TextMatrix(i%, 1))
            If (h$ <> "") Then
                h$ = Left$(h$, 2) + Mid$(h$, 4)
            End If
            Rufzeiten(j%).RufZeit = Val(h$)
            
            h$ = RTrim$(.TextMatrix(i%, 2))
            If (h$ <> "") Then
                h$ = Left$(h$, 2) + Mid$(h$, 4)
            End If
            Rufzeiten(j%).LieferZeit = Val(h$)
            
            BetrTage$ = LTrim$(RTrim$(.TextMatrix(i%, 4)))
            For k% = 0 To 6
                If (BetrTage$ = "") Then Exit For
                
                ind% = InStr(BetrTage$, ",")
                If (ind% > 0) Then
                    tag$ = RTrim$(Left$(BetrTage$, ind% - 1))
                    BetrTage$ = LTrim$(Mid$(BetrTage$, ind% + 1))
                Else
                    tag$ = BetrTage$
                    BetrTage$ = ""
                End If
                If (tag$ <> "") Then
                    Rufzeiten(j%).WoTag(k%) = Val(tag$)
                End If
            Next k%
            Do
                If (k% > 6) Then Exit Do
                Rufzeiten(j%).WoTag(k%) = 0
                k% = k% + 1
            Loop
            
            h$ = Left$(LTrim$(RTrim$(.TextMatrix(i%, 5))) + Space$(2), 2)
            Rufzeiten(j%).AuftragsErg = h$
            
            h$ = Left$(LTrim$(RTrim$(.TextMatrix(i%, 6))) + Space$(2), 2)
            Rufzeiten(j%).AuftragsArt = h$
            
            Aktiv$ = RTrim$(.TextMatrix(i%, 7))
            If (Aktiv$ = "ja") Then
                Rufzeiten(j%).Aktiv = "J"
            Else
                Rufzeiten(j%).Aktiv = "N"
            End If
            
            j% = j% + 1
        End If
    Next i%
End With
AnzRufzeiten% = j%



AnzMinutenWarnung% = Val(txtOptionenAutomatik(0).Text)
AnzMinutenWarten% = Val(txtOptionenAutomatik(1).Text)
BestVorsKomplett% = chkOptionenAutomatikBestVors(0).Value
BestVorsKomplettMinuten% = Val(txtOptionenAutomatikBestVors(0).Text)
BestVorsPeriodisch% = chkOptionenAutomatikBestVors(1).Value
BestVorsPeriodischMinuten% = Val(txtOptionenAutomatikBestVors(1).Text)

Call DefErrPop
End Sub

Sub AnzeigeContainer()
Dim h$, h2$
Dim c As Control

On Error Resume Next
If (prot% = True) Then
    For Each c In Controls
        h$ = c.Name
        h2$ = ""
        h2$ = RTrim(Format(c.Index, "00"))
        If (h2$ <> "") Then
            h$ = h$ + "(" + h2$ + ")"
        End If
        h$ = Left$(h$ + Space$(40), 40)
        h$ = h$ + c.Container.Name
        h2$ = ""
        h2$ = RTrim(Format(c.Container.Index, "00"))
        If (h2$ <> "") Then
            h$ = h$ + "(" + h2$ + ")"
        End If
        Print #PROTOKOLL%, h$
    Next c
End If

End Sub

Sub TabDisable()
Dim i%

For i% = 0 To 2
    flxOptionen(i%).Visible = False
Next i%

cmdF2.Visible = False
cmdF5.Visible = False

For i% = 0 To 1
    lblOptionenAutomatik(i%).Visible = False
    txtOptionenAutomatik(i%).Visible = False
    lblOptionenAutomatikMinuten(i%).Visible = False
Next i%
fmeOptionenAutomatikBestVors.Visible = False

End Sub

Sub TabEnable(hTab%)
Dim i%

If (hTab% < 3) Then
    flxOptionen(hTab%).Visible = True
    flxOptionen(hTab%).SetFocus
    
    cmdF2.Visible = True
    cmdF5.Visible = True
Else
    For i% = 0 To 1
        lblOptionenAutomatik(i%).Visible = True
        txtOptionenAutomatik(i%).Visible = True
        lblOptionenAutomatikMinuten(i%).Visible = True
    Next i%
    fmeOptionenAutomatikBestVors.Visible = True
    txtOptionenAutomatik(0).SetFocus
End If

End Sub

Sub EditOptionenLst()
Dim i%, hTab%, row%, col%, xBreit%, ind%, aRow%
Dim s$, h$

hTab% = tabOptionen.Tab
row% = flxOptionen(hTab%).row
col% = flxOptionen(hTab%).col
s$ = RTrim$(flxOptionen(hTab%).TextMatrix(row%, 0))

xBreit% = False
                
With flxOptionen(hTab%)
    aRow% = .row
    .row = 0
    .CellFontBold = True
    .row = aRow%
End With
            

With frmEdit.lstEdit
    .Clear
    Select Case col%
        Case 0
            If (hTab% < 2) Then
                .AddItem "Absagen"
                .AddItem "Anfragen"
                .AddItem "Besorger"
                .AddItem "def.Lagerartikel"
                .AddItem "Manuelle"
                .AddItem "Lagerartikel"
                .AddItem "Sonderangebote"
                .AddItem String$(50, "-")
                .AddItem "BM"
                .AddItem "EK"
                .AddItem "VK"
                .AddItem "Zeilenwert"
                .AddItem String$(50, "-")
                .AddItem "Hersteller"
                .AddItem "Lieferant"
    '                            .AddItem "Zwischenbestellung"
            Else
                For i% = 1 To AnzLiefNamen%
                    h$ = LiefNamen$(i% - 1)
                    .AddItem h$
                Next i%
            End If
            
        Case 1
            Select Case ZeilenTyp%(s$)
                Case 0
                    .AddItem "vor"
                    .AddItem "nach"
                    If (s$ = "Besorger") Then
                        .AddItem "älter als"
                    End If
                
                Case 1
                    .AddItem "="
                    .AddItem "<>"

                Case Else
                    .AddItem "<"
                    .AddItem "<="
                    .AddItem "="
                    .AddItem "<>"
                    .AddItem ">="
                    .AddItem ">"
    '                        .AddItem "?-fach"
            End Select
    
        Case 2
            If (s$ = "Lieferant") Then
                For i% = 1 To AnzLiefNamen%
                    h$ = LiefNamen$(i% - 1)
                    .AddItem h$
                Next i%
            ElseIf (s$ = "Besorger") Then
                .AddItem "1 Tag"
                .AddItem "2 Tage"
                .AddItem "3 Tage"
                .AddItem "4 Tage"
                .AddItem "5 Tage"
                .AddItem "Eingabe"
            Else
                .AddItem "1"
                .AddItem "2"
                .AddItem "3"
                .AddItem "4"
                .AddItem "5"
                .AddItem "Eingabe"
                If (s$ = "BM") Then
                    .AddItem "BMopt"
                    .AddItem "?-fach"
                End If
            End If
    
        Case 3
            .AddItem "Nicht Senden"
            .AddItem "Senden"
    
        Case 5
            If (hTab% = 2) Then
                .AddItem "Zustellung Heute (ZH)"
                .AddItem "Zustellung Morgen (ZM)"
                .AddItem "Heute kein Auftrag (KA)"
                xBreit% = True
            End If
        
        Case 6
            If (hTab% = 2) Then
                .AddItem "Normalauftrag (  )"
                .AddItem "Inventurliste (IN)"
                .AddItem "Lochkarten (LK)"
                .AddItem "Rückkauf-Anfrage (RK)"
                .AddItem "SBL-Auftrag (SB)"
                .AddItem "Sonder-Auftrag (SO)"
                .AddItem "Stapel-Auftrag (ST)"
                .AddItem "Test-Auftrag (TE)"
                .AddItem "Verfalldatenliste (VD)"
                .AddItem "Vorratskauf (VR)"
                .AddItem "10er-Auftrag (ZE)"
                .AddItem "Zeit-Auftrag (ZT)"
                .AddItem "Zeilenwert-Auftrag (ZW)"
                xBreit% = True
            End If
        
        Case 7
            If (hTab% = 2) Then
                .AddItem "ja"
                .AddItem "nein"
            End If
    End Select


    .ListIndex = 0
    s$ = RTrim$(flxOptionen(hTab%).TextMatrix(row%, col%))
    If (InStr(s$, "*") > 0) Then s$ = "?-fach"
    If (s$ <> "") Then
        If (col% = 5) Or (col% = 6) Then
            s$ = "(" + s$ + ")"
        End If
        For i% = 0 To (.ListCount - 1)
            .ListIndex = i%
            h$ = .Text
            If (col% = 4) Or (col% = 5) Then
                If (InStr(h$, s$) > 0) Then
                    Exit For
                End If
            ElseIf (s$ = h$) Then
                Exit For
            End If
        Next i%
    End If
    
    Load frmEdit
    
    With frmEdit
        .Left = tabOptionen.Left + flxOptionen(hTab%).Left + flxOptionen(hTab%).ColPos(col%) + 45
        .Left = .Left + Me.Left + wpara.FrmBorderHeight
        .Top = tabOptionen.Top + flxOptionen(hTab%).Top + flxOptionen(hTab%).RowPos(1)
        .Top = .Top + Me.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
        If (xBreit%) Then
            .Width = TextWidth("Zeilenwert-Auftrag (ZW)") + wpara.FrmScrollHeight + 90
        Else
            .Width = flxOptionen(hTab%).ColWidth(col%)
        End If
        .Height = flxOptionen(hTab%).Height - flxOptionen(hTab%).RowPos(1)
    End With
    With frmEdit.lstEdit
        .Height = frmEdit.ScaleHeight
        frmEdit.Height = .Height
        .Width = frmEdit.ScaleWidth
        .Left = 0
        .Top = 0
        
        .Visible = True
    End With
    
    frmEdit.Show 1
    
    With flxOptionen(hTab%)
        aRow% = .row
        .row = 0
        .CellFontBold = False
        .row = aRow%
    End With
            

    If (EditErg%) Then

        h$ = EditTxt$
'        If ((h$ = "Eingabe") Or (h$ = "?-fach")) Then
'            With txtOptionen
'                BmMulti% = False
'                If (h$ = "?-fach") Then BmMulti% = True
'
'                .Top = tabOptionen.Top + flxOptionen(hTab%).Top + (row% - flxOptionen(hTab%).TopRow + 1) * flxOptionen(hTab%).RowHeight(0)
'                .Height = flxOptionen(hTab%).RowHeight(1)
'                .Left = tabOptionen.Left + flxOptionen(hTab%).Left + flxOptionen(hTab%).ColPos(col%) + 15
'                .Width = flxOptionen(hTab%).ColWidth(col%)
'
'                h2$ = flxOptionen(hTab%).TextMatrix(row%, col%)
'                If (BmMulti%) Then
'                    ind% = InStr(h2$, "*")
'                    If (ind% > 0) Then
'                        .text = Left$(h2$, ind% - 1)
'                    Else
'                        .text = h2$
'                    End If
'                Else
'                    ind% = InStr(h2$, "Tage")
'                    If (ind% > 0) Then
'                        .text = Left$(h2$, ind% - 1)
'                    Else
'                        .text = h2$
'                    End If
'                End If
'                .SelStart = 0
'                .SelLength = Len(.text)
'                .BackColor = vbRed
'                .Visible = True
'                .TabStop = True
'                .SetFocus
'            End With
'        Else
            With flxOptionen(hTab%)
                If (Left$(h$, 1) <> "-") Then
                    If (col% = 5) Or (col% = 6) Then
                        ind% = InStr(h$, "(")
                        If (ind% > 0) Then
                            h$ = Mid$(h$, ind% + 1, 2)
                        End If
                    End If
                    .TextMatrix(row%, col%) = h$
                    If (.col < .Cols - 2) Then .col = .col + 1
                    If (.col < .Cols - 2) And (.ColWidth(.col) = 0) Then .col = .col + 1
                End If
            End With
'        End If
        
    End If





End With

End Sub
                            
                            
Sub EditOptionenLstMulti()
Dim i%, j%, l%, hTab%, row%, col%, ind%, aRow%
Dim s$, h$, BetrLief$, Lief2$

hTab% = tabOptionen.Tab
row% = flxOptionen(hTab%).row
col% = flxOptionen(hTab%).col
s$ = RTrim$(flxOptionen(hTab%).TextMatrix(row%, 0))
                
With flxOptionen(hTab%)
    aRow% = .row
    .row = 0
    .CellFontBold = True
    .row = aRow%
End With
            
With frmEdit.lstMultiEdit
    .Clear
    Select Case col%
        Case 3
            If (hTab% = 2) Then
                .AddItem "(keiner)"
                .AddItem "Montag"
                .AddItem "Dienstag"
                .AddItem "Mittwoch"
                .AddItem "Donnerstag"
                .AddItem "Freitag"
                .AddItem "Samstag"
                .AddItem "Sonntag"
                
            ElseIf (hTab% > 0) Then
                .AddItem "(keiner)"
                .AddItem String$(50, "-")
                .AddItem "Nächstliefernder (255)"
                For i% = 1 To AnzLiefNamen%
                    h$ = LiefNamen$(i% - 1)
                    .AddItem h$
                Next i%
            End If
    End Select

    For i% = 0 To (.ListCount - 1)
        .Selected(i%) = False
    Next i%

    
    Load frmEdit
    
     .ListIndex = 0
     
     If (hTab% < 2) Then
         BetrLief$ = LTrim$(RTrim$(flxOptionen(hTab%).TextMatrix(row%, 4)))
     Else
         BetrLief$ = LTrim$(RTrim$(flxOptionen(hTab%).TextMatrix(row%, 3)))
     End If
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
             If (hTab% < 2) Then
                 h$ = "(" + Mid$(Str$(ind%), 2) + ")"
                 For j% = 0 To (.ListCount - 1)
                     .ListIndex = j%
                     s$ = .Text
                     If (InStr(s$, h$) > 0) Then
                         .Selected(j%) = True
                         Exit For
                     End If
                 Next j%
             Else
                 .Selected(ind%) = True
             End If
         End If
     Next i%
     
    With frmEdit
        .Left = tabOptionen.Left + flxOptionen(hTab%).Left + flxOptionen(hTab%).ColPos(col%) + 45
        .Left = .Left + Me.Left + wpara.FrmBorderHeight
        .Top = tabOptionen.Top + flxOptionen(hTab%).Top + flxOptionen(hTab%).RowPos(1)
        .Top = .Top + Me.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
        .Width = flxOptionen(hTab%).ColWidth(col%)
        .Height = flxOptionen(hTab%).Height - flxOptionen(hTab%).RowPos(1)
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

    With flxOptionen(hTab%)
        aRow% = .row
        .row = 0
        .CellFontBold = False
        .row = aRow%
    End With
            

    If (EditErg%) Then
            
        flxOptionen(hTab%).TextMatrix(row%, col% + 1) = EditTxt$
        
        h$ = ""
        If (hTab% = 1) Then
            If (EditGef%(0) = 255) Then
                flxOptionen(hTab%).TextMatrix(row%, col% + 1) = "255"
                h$ = "Naechstliefernder"
            Else
                If (EditAnzGefunden% = 0) Then
                    h$ = ""
                ElseIf (EditAnzGefunden% = 1) Then
                    ind% = EditGef%(0)
                    lif.GetRecord (ind% + 1)
                    h$ = lif.kurz
                    Call OemToChar(h$, h$)
                Else
                    h$ = "mehrere (" + Mid$(Str$(EditAnzGefunden%), 2) + ")"
                End If
            End If
        Else
            If (EditAnzGefunden% = 0) Then
            ElseIf (EditAnzGefunden% = 1) Then
                ind% = EditGef%(0)
                h$ = h$ + WochenTag$(ind% - 1)
            ElseIf (EditAnzGefunden% < 4) Then
                For l% = 0 To (EditAnzGefunden% - 1)
                    ind% = EditGef%(l%)
                    h$ = h$ + WochenTag$(ind% - 1) + ","
                Next l%
            Else
                For l% = 0 To (EditAnzGefunden% - 1)
                    ind% = EditGef%(l%)
                    h$ = h$ + Left$(WochenTag$(ind% - 1), 2) + ","
                Next l%
            End If
        End If

        With flxOptionen(hTab%)
            .TextMatrix(row%, col%) = h$
            If (.col < .Cols - 2) Then .col = .col + 1
            If (.col < .Cols - 2) And (.ColWidth(.col) = 0) Then .col = .col + 1
        End With
        
    End If

End With

End Sub
                            
Sub EditOptionenTxt()
Dim i%, hTab%, row%, col%, aRow%, uhr%, st%, min%, iZeilenTyp%
Dim s$, h$

EditModus% = 2

hTab% = tabOptionen.Tab
row% = flxOptionen(hTab%).row
col% = flxOptionen(hTab%).col
s$ = RTrim$(flxOptionen(hTab%).TextMatrix(row%, 0))
iZeilenTyp% = ZeilenTyp%(s$)
                
With flxOptionen(hTab%)
    aRow% = .row
    .row = 0
    .CellFontBold = True
    .row = aRow%
End With
            
With frmEdit.txtEdit
    Select Case col%
        Case 1
            If (hTab% = 2) Then
                .MaxLength = 4
                h$ = flxOptionen(hTab%).TextMatrix(row%, col%)
                .Text = Right$("0000" + Left$(h$, 2) + Mid$(h$, 4), 4)
            End If
        Case 2
            If (hTab% = 2) Then
                .MaxLength = 4
                h$ = flxOptionen(hTab%).TextMatrix(row%, col%)
                .Text = Right$("0000" + Left$(h$, 2) + Mid$(h$, 4), 4)
            Else
                If (iZeilenTyp% = 0) Then
                    .MaxLength = 4
                    h$ = flxOptionen(hTab%).TextMatrix(row%, col%)
                    .Text = Right$("0000" + Left$(h$, 2) + Mid$(h$, 4), 4)
                ElseIf (iZeilenTyp% = 1) Then
                    .MaxLength = 5
                    h$ = flxOptionen(hTab%).TextMatrix(row%, col%)
                    .Text = Right$("     " + h$, 5)
                    EditModus% = 1
                Else
                    .MaxLength = 4
                    h$ = flxOptionen(hTab%).TextMatrix(row%, col%)
                    .Text = Right$("    " + h$, 4)
                End If
            End If
    End Select


    
    Load frmEdit
    
    With frmEdit
        .Left = tabOptionen.Left + flxOptionen(hTab%).Left + flxOptionen(hTab%).ColPos(col%) + 45
        .Left = .Left + Me.Left + wpara.FrmBorderHeight
        .Top = tabOptionen.Top + flxOptionen(hTab%).Top + (row% - flxOptionen(hTab%).TopRow + 1) * flxOptionen(hTab%).RowHeight(0)
        .Top = .Top + Me.Top + wpara.FrmBorderHeight + wpara.FrmCaptionHeight
        .Width = flxOptionen(hTab%).ColWidth(col%)
        .Height = frmEdit.txtEdit.Height 'flxarbeit(0).RowHeight(1)
    End With
    With frmEdit.txtEdit
        .Width = frmEdit.ScaleWidth
'            .Height = frmEdit.ScaleHeight
        .Left = 0
        .Top = 0
        .BackColor = vbWhite
        .Visible = True
    End With
    
    
    frmEdit.Show 1
    
    With flxOptionen(hTab%)
        aRow% = .row
        .row = 0
        .CellFontBold = False
        .row = aRow%
    End With
            

    If (EditErg%) Then

        If (col% = 1) Or ((col% = 2) And ((hTab% = 2) Or (iZeilenTyp% = 0))) Then
            uhr% = Val(EditTxt$)
            st% = uhr% \ 100
            min% = uhr% Mod 100
            h$ = Format(st%, "00") + ":" + Format(min%, "00")
        ElseIf (iZeilenTyp% = 1) Then
            h$ = UCase$(Trim$(EditTxt$))
        Else
            h$ = Mid$(Str(Val(EditTxt$)), 2)
        End If
        With flxOptionen(hTab%)
            flxOptionen(hTab%).TextMatrix(row%, col%) = h$
            If (.col < .Cols - 2) Then .col = .col + 1
            If (.col < .Cols - 2) And (.ColWidth(.col) = 0) Then .col = .col + 1
        End With
    End If

End With

End Sub

