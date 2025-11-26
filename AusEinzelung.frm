VERSION 5.00
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlbutton.ocx"
Begin VB.Form frmAusEinzelung 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Auseinzelung"
   ClientHeight    =   7695
   ClientLeft      =   240
   ClientTop       =   1545
   ClientWidth     =   14205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   14205
   Begin VB.PictureBox picfmeAuseinzelung 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   3855
      Left            =   7920
      ScaleHeight     =   3855
      ScaleWidth      =   5055
      TabIndex        =   9
      Tag             =   "2.Wirkstärke"
      Top             =   3240
      Width           =   5055
      Begin VB.TextBox txtAuseinzelung 
         Height          =   375
         Index           =   3
         Left            =   3360
         MaxLength       =   8
         TabIndex        =   11
         Text            =   "9999999"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtAuseinzelung 
         Height          =   375
         Index           =   4
         Left            =   3480
         MaxLength       =   5
         TabIndex        =   14
         Text            =   "99999"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtAuseinzelung 
         Height          =   375
         Index           =   5
         Left            =   3600
         MaxLength       =   7
         TabIndex        =   17
         Text            =   "9999999"
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblAuseinzelung2 
         Height          =   375
         Index           =   3
         Left            =   4680
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblAuseinzelung2 
         Caption         =   "Anzahl ausgeeinzelt (von 99999)"
         Height          =   375
         Index           =   4
         Left            =   4560
         TabIndex        =   15
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblAuseinzelung2 
         Caption         =   "Cent"
         Height          =   375
         Index           =   5
         Left            =   4680
         TabIndex        =   18
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblAuseinzelung 
         Caption         =   "PZN/&Name:"
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   10
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label lblAuseinzelung 
         Caption         =   "&Menge:"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label lblAuseinzelung 
         Caption         =   "&Preis:"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   3255
      End
   End
   Begin VB.CheckBox chkAuseinzelung 
      Caption         =   "2.Wirkstärke:"
      Height          =   255
      Index           =   0
      Left            =   9120
      TabIndex        =   8
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   7200
      Picture         =   "AusEinzelung.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   6960
      Picture         =   "AusEinzelung.frx":00B9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   6720
      Picture         =   "AusEinzelung.frx":016D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txtAuseinzelung 
      Height          =   375
      Index           =   0
      Left            =   3720
      MaxLength       =   8
      TabIndex        =   1
      Text            =   "9999999"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtAuseinzelung 
      Height          =   375
      Index           =   1
      Left            =   3840
      MaxLength       =   5
      TabIndex        =   3
      Text            =   "99999"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtAuseinzelung 
      Height          =   375
      Index           =   2
      Left            =   3960
      MaxLength       =   7
      TabIndex        =   6
      Text            =   "9999999"
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   3240
      TabIndex        =   19
      Top             =   2520
      Width           =   1200
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   4800
      TabIndex        =   20
      Top             =   2520
      Width           =   1200
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   495
      Left            =   4800
      TabIndex        =   25
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   495
      Left            =   3360
      TabIndex        =   26
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "nlCommand"
   End
   Begin VB.Label lblchkAuseinzelung 
      Caption         =   "AAAA"
      Height          =   375
      Index           =   0
      Left            =   11040
      TabIndex        =   27
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblAuseinzelung2 
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   21
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblAuseinzelung 
      Caption         =   "PZN/&Name:"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label lblAuseinzelung 
      Caption         =   "&Menge:"
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label lblAuseinzelung2 
      Caption         =   "Anzahl ausgeeinzelt (von 99999)"
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblAuseinzelung 
      Caption         =   "&Preis:"
      Height          =   375
      Index           =   2
      Left            =   600
      TabIndex        =   5
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label lblAuseinzelung2 
      Caption         =   "Cent"
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   7
      Top             =   1440
      Width           =   1695
   End
End
Attribute VB_Name = "frmAusEinzelung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iEditModus%

Dim OrgMenge&(1)
Dim OrgPreis&(1)

Private Const DefErrModul = "AUSEINZELUNG.FRM"

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
Dim i%, k%, IstPzn%, iAdd%, ind%, ind2%, iOk%
Dim pzn$, txt$, mErg$

If (ActiveControl.Name = txtAuseinzelung(0).Name) Then
    If (ActiveControl.Index Mod 3 = 0) Then
        iAdd = IIf(ActiveControl.Index < 3, 0, 3)
        ind2 = IIf(ActiveControl.Index < 3, 0, 1)
        pzn = ""
        txt = Trim(txtAuseinzelung(0 + iAdd).text)
        IstPzn = (txt <> "")
        For i% = 1 To Len(txt)
            If (InStr("0123456789", Mid$(txt$, i%, 1)) = 0) Then
                IstPzn = 0
                Exit For
            End If
        Next i%
        If (IstPzn) Then
            pzn = PznString(Val(txt)) ' Right("0000000" + txt, 7)
            mErg = pzn
        Else
            mErg$ = MatchCode(0, pzn$, txt$, (txt$ <> ""), True)
        End If
        If (mErg <> "") Then
            SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + pzn$
            'Set TaxeRec = TaxeDB.OpenRecordset(SQLStr$)
            On Error Resume Next
            TaxeRec.Close
            Err.Clear
            On Error GoTo DefErr
            TaxeRec.open SQLStr, taxeAdoDB.ActiveConn
            If (TaxeRec.EOF) Then
                Call DefErrPop: Exit Sub
            End If
            
            txtAuseinzelung(0 + iAdd).text = PznString(TaxeRec!pzn)
'            OrgMenge = Mid$(TaxeRec!menge, 3)
            OrgMenge(ind2) = TaxeRec!StdMenge
            txtAuseinzelung(1 + iAdd).text = CStr(OrgMenge(ind2))
            lblAuseinzelung2(1 + iAdd).Caption = "Anzahl ausgeeinzelt (von " + CStr(OrgMenge(ind2)) + ")"
            
            If (TaxeRec!FestKz) And (TaxeRec!FESTBETRAG < TaxeRec!vk) Then
                OrgPreis(ind2) = TaxeRec!FESTBETRAG
            Else
                OrgPreis(ind2) = TaxeRec!vk
            End If
            txtAuseinzelung(2 + iAdd).text = CStr(OrgPreis(ind2))
            
            If (TaxeRec!BtmKz) Then
                AuseinzelungBtm = (iMsgBox("BTM-Gebühr hinzufügen?", vbYesNo Or vbDefaultButton1) = vbYes)
            End If
        End If
    End If
    MySendKeys "{TAB}", True
Else
    ind = 0
    iOk = 0
    
    For k = 0 To 1
        iAdd = IIf(k = 0, 0, 3)
        
        IstPzn = True
        For i = 0 To 2
            If (Trim(txtAuseinzelung(i + iAdd)) = "") Then
                IstPzn = 0
                Exit For
            End If
        Next i
        If (IstPzn) Then
            AuseinzelungPzn(ind) = txtAuseinzelung(0 + iAdd).text
            AuseinzelungFaktor(ind) = xVal(txtAuseinzelung(1 + iAdd).text) / OrgMenge(k) * 1000
            AuseinzelungPreis(ind) = xVal(txtAuseinzelung(2 + iAdd).text) / 100#
            AuseinzelungPreisGesamt = AuseinzelungPreisGesamt + AuseinzelungPreis(ind)
            ind = ind + 1
            iOk = True
        End If
    Next k
    If (AuseinzelungBtm) Then
        AuseinzelungPreisGesamt = AuseinzelungPreisGesamt + 3.58
    End If
    If (iOk) Then
        Unload Me
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
Dim i%, k%, Breite%, MaxWi%, wi%, diff%, FormVersatzY%
Dim iAdd%, iAdd2%
Dim c As Control

AuseinzelungPzn$(0) = ""
AuseinzelungPreisGesamt = 0
OrgMenge(0) = 0
OrgMenge(1) = 0

iEditModus = 1

EditErg% = 0

Call wpara.InitFont(Me)

On Error Resume Next
For Each c In Controls
    If (TypeOf c Is TextBox) Then
        c.Width = TextWidth(c.text) + 450   ' 90
        c.text = ""
    End If
Next
On Error GoTo DefErr

'txtImpAlternativ(0).text = "0101" + Format(Now, "YY")
'txtImpAlternativ(1).text = "1"

With chkAuseinzelung(0)
    .Top = wpara.TitelY
End With
For k = 0 To 1
    iAdd = IIf(k = 0, 0, 3)
    
    txtAuseinzelung(0 + iAdd).Top = IIf(k = 0, chkAuseinzelung(0).Top + chkAuseinzelung(0).Height + 300, chkAuseinzelung(0).Top) ' 2 * wpara.TitelY
    For i% = 1 To 2
        txtAuseinzelung(i% + iAdd).Width = txtAuseinzelung(0).Width
        txtAuseinzelung(i% + iAdd).Top = txtAuseinzelung(i% + iAdd - 1).Top + txtAuseinzelung(i% + iAdd - 1).Height + 90
    Next i%
    
    diff% = (txtAuseinzelung(0 + iAdd).Height - lblAuseinzelung(0 + iAdd).Height) / 2
    lblAuseinzelung(0 + iAdd).Left = wpara.LinksX
    lblAuseinzelung(0 + iAdd).Top = txtAuseinzelung(0 + iAdd).Top + diff%
    For i% = 1 To 2
        lblAuseinzelung(i% + iAdd).Left = lblAuseinzelung(0 + iAdd).Left + lblAuseinzelung(0 + iAdd).Width - lblAuseinzelung(i% + iAdd).Width
        lblAuseinzelung(i% + iAdd).Top = txtAuseinzelung(i% + iAdd).Top + diff%
    Next i%
    
    MaxWi% = 0
    For i% = 0 To 2
        wi% = lblAuseinzelung(i% + iAdd).Width
        If (wi% > MaxWi%) Then
            MaxWi% = wi%
        End If
    Next i%
    
    txtAuseinzelung(0 + iAdd).Left = lblAuseinzelung(0 + iAdd).Left + MaxWi% + 300
    For i% = 1 To 2
        txtAuseinzelung(i% + iAdd).Left = txtAuseinzelung(i% + iAdd - 1).Left
    Next i%
    
    lblAuseinzelung2(0 + iAdd).Left = txtAuseinzelung(0 + iAdd).Left + txtAuseinzelung(0 + iAdd).Width + 150
    lblAuseinzelung2(0 + iAdd).Top = txtAuseinzelung(0 + iAdd).Top
    For i% = 1 To 2
        lblAuseinzelung2(i% + iAdd).Left = lblAuseinzelung2(i% + iAdd - 1).Left
        lblAuseinzelung2(i% + iAdd).Top = txtAuseinzelung(i% + iAdd).Top
    Next i%
    lblAuseinzelung2(1 + iAdd).Caption = ""
Next k
With chkAuseinzelung(0)
    .Left = lblAuseinzelung2(1).Left + lblAuseinzelung2(1).Width + 900
    .Value = 0
End With
With picfmeAuseinzelung
    .Left = chkAuseinzelung(0).Left
    .Top = chkAuseinzelung(0).Top + chkAuseinzelung(0).Height + 90 ' lblAuseinzelung(0).Top
    .Width = lblAuseinzelung2(4).Left + lblAuseinzelung2(4).Width + 2 * wpara.LinksX '+ 1500
    .Height = lblAuseinzelung2(5).Top + lblAuseinzelung2(5).Height + 300
    .Visible = False
End With

Font.Name = wpara.FontName(1)
Font.Size = wpara.FontSize(1)

'Me.Width = fmeImpAlternativ.Left + fmeImpAlternativ.Width + 2 * wpara.LinksX
Me.Width = lblAuseinzelung2(2).Left + lblAuseinzelung2(2).Width + 2 * wpara.LinksX + 1500
Me.Width = lblAuseinzelung2(4).Left + lblAuseinzelung2(1).Width + 2 * wpara.LinksX '+ 1500
Me.Width = picfmeAuseinzelung.Left + picfmeAuseinzelung.Width + 2 * wpara.LinksX '+ 1500

With cmdOk
'    .Top = fmeImpAlternativ.Top + fmeImpAlternativ.Height + 150 * wpara.BildFaktor
    .Top = txtAuseinzelung(2).Top + txtAuseinzelung(2).Height + 900
    .Width = wpara.ButtonX
    .Height = wpara.ButtonY
    .Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
End With
With cmdEsc
    .Top = cmdOk.Top
    .Width = wpara.ButtonX
    .Height = wpara.ButtonY
    .Left = cmdOk.Left + cmdEsc.Width + 300
End With

Me.Height = cmdOk.Top + cmdOk.Height + wpara.TitelY + 90 + wpara.FrmCaptionHeight

If (para.Newline) Then
    iAdd = wpara.NlFlexBackY
    iAdd2 = wpara.NlCaptionY
    
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
        .Top = txtAuseinzelung(2).Top + txtAuseinzelung(2).Height + iAdd + 900
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
            ElseIf (TypeOf c Is CheckBox) Then
                c.Height = 0
                c.Width = c.Height
                If (c.Name = "chkAuseinzelung") Then
                    If (c.Index > 0) Then
                        Load lblchkAuseinzelung(c.Index)
                    End If
                    With lblchkAuseinzelung(c.Index)
                        .BackStyle = 0 'duchsichtig
                        .Caption = c.Caption
                        .Left = c.Left + 300
                        .Top = c.Top
                        .Width = TextWidth(.Caption) + 90
                        .TabIndex = c.TabIndex
                        .Visible = True
                    End With
                End If
            End If
        End If
    
        If (Left(c.Name, 6) = "picfme") Then
            With c
                If (para.Newline) Then
                    Dim y%
                    Dim lColor&(1)

                    lColor(0) = GetPixel(.Container.hdc, c.Left / Screen.TwipsPerPixelX - 2, c.Top / Screen.TwipsPerPixelY)
                    lColor(1) = GetPixel(.Container.hdc, c.Left / Screen.TwipsPerPixelX - 2, (c.Top + c.Height) / Screen.TwipsPerPixelY)
    '                Call wpara.FillGradient(c, 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, RGB(177, 177, 177), RGB(225, 225, 225))
                    Call wpara.FillGradient(c, 0, 0, .ScaleWidth / Screen.TwipsPerPixelX, .ScaleHeight / Screen.TwipsPerPixelY, lColor(0), lColor(1))
                Else
                    .BackColor = Me.BackColor
                End If
                
                .ForeColor = vbBlack ' RGB(150, 150, 150)
                .FillStyle = vbFSTransparent ' vbSolid
                .FillColor = vbWhite
            
                y = 30 + .TextHeight(.tag) / 2
                RoundRect .hdc, 30 / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY, (.Width - 30) / Screen.TwipsPerPixelX, (.Height - 30) / Screen.TwipsPerPixelY, 20, 20
                
                c.Line (240, y)-(300 + .TextWidth(.tag) + 90, y), RGB(180, 180, 180)
                
                .CurrentX = 300
                .CurrentY = 30
                c.Print .tag
            End With
        End If
    Next
    On Error GoTo DefErr
    
Else
    nlcmdOk.Visible = False
    nlcmdEsc.Visible = False
End If
'''''''''

Me.Left = frmRezSpeicher.Left + (frmRezSpeicher.Width - Me.Width) / 2
If (Me.Left < 0) Then
    Me.Left = 0
End If

Me.Top = frmRezSpeicher.Top + (frmRezSpeicher.Height - Me.Height) / 2
If (Me.Top < 0) Then
    Me.Top = 0
End If

Call DefErrPop
End Sub

Private Sub txtAuseinzelung_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtAuseinzelung_GotFocus")
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

With txtAuseinzelung(Index)
'    h$ = .text
'    For i% = 1 To Len(h$)
'        If (Mid$(h$, i%, 1) = ",") Then Mid$(h$, i%, 1) = "."
'    Next i%
'    .text = h$
    .SelStart = 0
    .SelLength = Len(.text)
End With

Call DefErrPop
End Sub

Private Sub txtAuseinzelung_KeyPress(Index As Integer, KeyAscii As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtAuseinzelung_KeyPress")
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

If (Index Mod 3 > 0) Then
    If (KeyAscii >= 32) And ((KeyAscii < 48) Or (KeyAscii > 57)) Then
        Beep
        KeyAscii = 0
    End If
End If

Call DefErrPop
End Sub

Private Sub txtAuseinzelung_Change(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("txtAuseinzelung_Change")
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
Dim iAdd%, ind2%
Dim dFaktor#

If (Me.Visible) Then
    If (Index Mod 3 = 1) Then
        iAdd = IIf(ActiveControl.Index < 3, 0, 3)
        ind2 = IIf(ActiveControl.Index < 3, 0, 1)
        If (OrgMenge(ind2) > 0) Then
            dFaktor = xVal(txtAuseinzelung(1 + iAdd).text) / OrgMenge(ind2) '* 1000
            txtAuseinzelung(2 + iAdd).text = CStr(CLng(OrgPreis(ind2) * dFaktor))
        End If
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

Private Sub picControlBox_Click(Index As Integer)

If (Index = 0) Then
    Me.WindowState = vbMinimized
ElseIf (Index = 1) Then
    Me.WindowState = vbNormal
Else
    Unload Me
End If

End Sub

Private Sub chkAuseinzelung_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("chkAuseinzelung_Click")
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

picfmeAuseinzelung.Visible = chkAuseinzelung(0).Value

Call DefErrPop
End Sub


Private Sub lblchkAuseinzelung_Click(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("lblchkAuseinzelung_Click")
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

With chkAuseinzelung(Index)
    If (.Enabled) Then
        If (.Value) Then
            .Value = 0
        Else
            .Value = 1
        End If
        .SetFocus
    End If
End With

Call DefErrPop
End Sub

Private Sub chkAuseinzelung_GotFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("chkAuseinzelung_GotFocus")
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

Call nlCheckBox(chkAuseinzelung(Index).Name, Index)

Call DefErrPop
End Sub

Private Sub chkAuseinzelung_LostFocus(Index As Integer)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("chkAuseinzelung_LostFocus")
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

Call nlCheckBox(chkAuseinzelung(Index).Name, Index, 0)

Call DefErrPop
End Sub

Sub nlCheckBox(sCheckBox$, Index As Integer, Optional GotFocus% = True)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("nlCheckBox")
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
Dim ok%
Dim Such$
Dim c As Object

Such = "lbl" + sCheckBox

On Error Resume Next
For Each c In Controls
    If (c.Name = Such) Then
        ok = True
        If (Index >= 0) Then
            ok = (c.Index = Index)
        End If
        If (ok) Then
            If (GotFocus) Then
'                c.Font.underline = True
'                c.ForeColor = vbHighlight
                c.BackStyle = 1
                c.BackColor = vbHighlight
                c.ForeColor = vbWhite
            Else
'                c.Font.underline = 0
'                c.ForeColor = vbBlack
                c.BackStyle = 0
                c.BackColor = vbHighlight
                c.ForeColor = vbBlack
            End If
        End If
    End If
Next
On Error GoTo DefErr

Call DefErrPop
End Sub






