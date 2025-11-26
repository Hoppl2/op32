VERSION 5.00
Object = "{65ED66DD-4DAA-499D-95C5-98B8A92C2A2B}#63.0#0"; "nlbutton.ocx"
Begin VB.Form frmRFID 
   Caption         =   "Benutzer-Signatur"
   ClientHeight    =   9660
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   12300
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   12300
   Begin VB.Timer tmrRFID 
      Interval        =   200
      Left            =   4080
      Top             =   5520
   End
   Begin VB.PictureBox picRFID 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3300
      Left            =   960
      ScaleHeight     =   3270
      ScaleWidth      =   4350
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   4380
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   0
      Left            =   3360
      Picture         =   "RFID.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   1
      Left            =   3600
      Picture         =   "RFID.frx":00A9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picControlBox 
      BorderStyle     =   0  'Kein
      Height          =   240
      Index           =   2
      Left            =   3840
      Picture         =   "RFID.frx":015D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   1920
      TabIndex        =   2
      Top             =   5520
      Width           =   1200
   End
   Begin VB.TextBox txtSignatur 
      Appearance      =   0  '2D
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4200
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   360
      TabIndex        =   1
      Top             =   5520
      Width           =   1200
   End
   Begin nlCommandButton.nlCommand nlcmdEsc 
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   6120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin nlCommandButton.nlCommand nlcmdOk 
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   6120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "nlCommand"
   End
   Begin VB.Label lblRFID 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Lesen Sie bitte Ihren Personal-Chip ein ...."
      Height          =   615
      Left            =   6000
      TabIndex        =   11
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label lblODER 
      Alignment       =   2  'Zentriert
      Caption         =   "O D E R"
      Height          =   615
      Left            =   6240
      TabIndex        =   10
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label lblSignatur 
      BackStyle       =   0  'Transparent
      Caption         =   "Bitte geben Sie Ihr &Passwort ein:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   4680
      Width           =   3735
   End
End
Attribute VB_Name = "frmRFID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "RFID.FRM"

Dim RFID_Tabelle As String
Dim RFID_Comp As Integer
Dim SQLStr$
Dim lRecs&

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
Dim iBenutzerInd%, okAktiv%

'If (iNewLine) Then
'    okAktiv = (ActiveControl.Name = nlcmdOk.Name)
'Else
'    okAktiv = (ActiveControl.Name = cmdOk.Name)
'End If
'If (okAktiv) Then
    iBenutzerInd% = CheckPasswort%
    If (iBenutzerInd% > 0) Then
        SignaturEingabeErg% = iBenutzerInd%
        Unload Me
    Else
        txtSignatur.SetFocus
    End If
'Else
'    clsSI.MySendKeys "{TAB}", True
'End If

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
Dim iAdd%, iAdd2%, x%, y%, wi%, ydiff%, IconDa%

SignaturEingabeErg% = 0

Call wPara1.InitFont(Me)

IconDa = (clsOpTool.FileExist(CurDir + "\opchip.jpg"))

With lblRFID
    .Left = wPara1.LinksX
    .Top = 4 * wPara1.TitelY
End With
If (IconDa) Then
    lblRFID.Visible = False
    With picRFID
        .Picture = LoadPicture(CurDir + "\opchip.jpg")
        .Left = lblRFID.Left
        .Top = lblRFID.Top
        .Visible = True
    End With
    With lblODER
        .Left = wPara1.LinksX
        .Top = picRFID.Top + picRFID.Height + 600
    End With
Else
    With lblODER
        .Left = wPara1.LinksX
        .Top = lblRFID.Top + lblRFID.Height + 600
    End With
End If

With lblSignatur
    .Left = wPara1.LinksX
    .Top = lblODER.Top + lblODER.Height + 600
End With
With txtSignatur
    ydiff% = (.Height - lblSignatur.Height) / Screen.TwipsPerPixelY
    ydiff% = (ydiff% \ 2) * Screen.TwipsPerPixelY
    .Top = lblSignatur.Top - ydiff%
    .Left = lblSignatur.Left + lblSignatur.Width + 300
    .Width = TextWidth(String(15, "X"))
End With

Me.Width = txtSignatur.Left + txtSignatur.Width + 3 * wPara1.LinksX

If (IconDa) Then
    With picRFID
        .Left = (Me.ScaleWidth - .Width) / 2
    End With
Else
    With lblRFID
        .Left = (Me.ScaleWidth - .Width) / 2
    End With
End If
With lblODER
    .Left = (Me.ScaleWidth - .Width) / 2
End With

With cmdOk
    .Width = wPara1.ButtonX
    .Height = wPara1.ButtonY
    .Left = (Me.Width - (cmdOk.Width * 2 + 300)) / 2
    .Top = lblSignatur.Top + lblSignatur.Height + 600
End With

With cmdEsc
    .Width = cmdOk.Width
    .Height = cmdOk.Height
    .Left = cmdOk.Left + cmdEsc.Width + 300
    .Top = cmdOk.Top
End With

If (SignaturEingabeModus And 1) Then
Else
    lblODER.Visible = False
    lblSignatur.Visible = False
    txtSignatur.Visible = False
    
    With cmdOk
        .Visible = False
        .Top = lblODER.Top
    End With
    
    With cmdEsc
        .Visible = False
        .Width = cmdOk.Width
        .Height = cmdOk.Height
        .Left = cmdOk.Left + cmdEsc.Width + 300
        .Top = cmdOk.Top
    End With
End If

Me.Height = cmdOk.Top + cmdOk.Height + wPara1.FrmCaptionHeight + 2 * wPara1.TitelY


''''''
If (iNewLine) Then
    iAdd = wPara1.NlFlexBackY
    iAdd2 = wPara1.NlCaptionY
    
    lblRFID.Top = lblRFID.Top + iAdd2
    picRFID.Top = picRFID.Top + iAdd2
    lblODER.Top = lblODER.Top + iAdd2
    
    lblSignatur.Top = lblSignatur.Top + iAdd2
    txtSignatur.Top = txtSignatur.Top + iAdd2
    cmdOk.Top = cmdOk.Top + iAdd2
    cmdEsc.Top = cmdOk.Top
    Height = Height + iAdd2
    
    With nlcmdOk
        .Init
        .Left = (Me.ScaleWidth - (.Width * 2 + 300)) / 2
        .Top = txtSignatur.Top + txtSignatur.Height + iAdd + 600
        .Caption = cmdOk.Caption
        .TabIndex = cmdOk.TabIndex
        .Enabled = cmdOk.Enabled
        .default = cmdOk.default
        .Cancel = cmdOk.Cancel
        .Visible = True
    End With
    cmdOk.Visible = False

    With nlcmdEsc
        .Init
        .Left = nlcmdOk.Left + .Width + 300
        .Top = nlcmdOk.Top
        .Caption = cmdEsc.Caption
        .TabIndex = cmdEsc.TabIndex
        .Enabled = cmdEsc.Enabled
        .default = cmdEsc.default
        .Cancel = cmdEsc.Cancel
        .Visible = True
    End With
    cmdEsc.Visible = False

    If (SignaturEingabeModus And 1) Then
    Else
        lblODER.Visible = False
        lblSignatur.Visible = False
        txtSignatur.Visible = False
        
        With nlcmdOk
            .Visible = False
            .Top = lblODER.Top
        End With
        
        With nlcmdEsc
            .Visible = False
'            .Width = nlcmdOk.Width
'            .Height = nlcmdOk.Height
'            .Left = nlcmdOk.Left + nlcmdEsc.Width + 300
'            .Top = nlcmdOk.Top
        End With
    End If

    Me.Height = nlcmdOk.Top + nlcmdOk.Height + wPara1.FrmCaptionHeight + 450
    
    Call wPara1.NewLineWindow(Me, nlcmdOk.Top)
'    RoundRect hdc, (flxPosPartner.Left - iAdd) / Screen.TwipsPerPixelX, (flxPosPartner.Top - iAdd) / Screen.TwipsPerPixelY, (flxPosPartner.Left + flxPosPartner.Width + iAdd) / Screen.TwipsPerPixelX, (flxPosPartner.Top + flxPosPartner.Height + iAdd) / Screen.TwipsPerPixelY, 20, 20
    With txtSignatur
'        .Appearance = 0
        .BackColor = vbWhite
        Call wPara1.ControlBorderless(txtSignatur, 1, 1)
    End With
    With Me
        .ForeColor = RGB(180, 180, 180) ' vbWhite
        .FillStyle = vbSolid
        .FillColor = vbWhite
        RoundRect .hdc, (txtSignatur.Left - 90) / Screen.TwipsPerPixelX, (txtSignatur.Top - 45) / Screen.TwipsPerPixelY, (txtSignatur.Left + txtSignatur.Width + 90) / Screen.TwipsPerPixelX, (txtSignatur.Top + txtSignatur.Height + 45) / Screen.TwipsPerPixelY, 10, 10
    End With

    'Call picTransparent(picRFID)
'    Me.Left = ProjektForm.Left + (ProjektForm.Width - Me.Width) / 2
'    Me.Top = ProjektForm.Top + (ProjektForm.Height - Me.Height) / 2
Else
    nlcmdOk.Visible = False
    nlcmdEsc.Visible = False
End If

''''''''
RFID_Tabelle = "RFID_Scan"
RFID_Comp = Val(Para1.user) * 10
SQLStr = "DELETE " + RFID_Tabelle + " WHERE Comp=" + CStr(RFID_Comp)
Call ArtikelDB1.ActiveConn.Execute(SQLStr, lRecs, adExecuteNoRecords)

Me.Left = ProjektForm.Left + (ProjektForm.Width - Me.Width) / 2
Me.Top = ProjektForm.Top + (ProjektForm.Height - Me.Height) / 2

Call clsError.DefErrPop
End Sub

Private Sub Form_Paint()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_Paint")
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
Dim i%, spBreite%, ind%, iAnzZeilen%, RowHe%, bis%, bis2%
Dim sp&
Dim h$, h2$
Dim iAdd%, iAdd2%, wi%
Dim c As Control

If (Para1.Newline) Then
    iAdd = wPara1.NlFlexBackY
    iAdd2 = wPara1.NlCaptionY
    
    Call wPara1.NewLineWindow(Me, nlcmdOk.Top, False)
    
    With Me
        .ForeColor = RGB(180, 180, 180) ' vbWhite
        .FillStyle = vbSolid
        .FillColor = vbWhite
        RoundRect .hdc, (txtSignatur.Left - 90) / Screen.TwipsPerPixelX, (txtSignatur.Top - 45) / Screen.TwipsPerPixelY, (txtSignatur.Left + txtSignatur.Width + 90) / Screen.TwipsPerPixelX, (txtSignatur.Top + txtSignatur.Height + 45) / Screen.TwipsPerPixelY, 10, 10
    End With

    Call Form_Resize
End If

Call clsError.DefErrPop
End Sub

Private Sub tmrRFID_Timer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("tmrRFID_Timer")
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

tmrRFID.Enabled = False

SQLStr$ = "SELECT * FROM " + RFID_Tabelle + " WHERE Comp=" + CStr(RFID_Comp)
FabsErrf = ArtikelDB1.OpenRecordset(ArtikelRec2, SQLStr, 0)
If (FabsErrf% = 0) Then
    Do
        If (ArtikelRec2.EOF) Then
            Exit Do
        End If
        
        SignaturEingabeErg = ArtikelRec2!MA
        
        Unload Me
        Call clsError.DefErrPop: Exit Sub
    Loop
End If

tmrRFID.Enabled = True

Call clsError.DefErrPop
End Sub

Private Sub txtSignatur_GotFocus()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("txtSignatur_GotFocus")
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

With txtSignatur
    .SelStart = 0
    .SelLength = Len(.text)
End With

Call clsError.DefErrPop
End Sub

Function CheckPasswort%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("CheckPasswort%")
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
Dim i%, ret%
Dim pass$, GeneralPW$

ret% = 0

pass$ = UCase(Trim(txtSignatur.text))
If (pass$ <> "") Then
    GeneralPW = "@" + Format(Now, "hhnnddmm")        '1.0.77
    If (pass = GeneralPW) Then
'        Para1.Passwort(1) = "OPTIPHARM"
        ret = 1
    Else
        For i% = 1 To 80
            If (pass$ = Para1.Passwort(i%)) Then
                ret% = i%
                Exit For
            End If
        Next i%
    End If
End If

CheckPasswort% = ret%

Call clsError.DefErrPop
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_MouseDown")
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
    
If (y <= wPara1.NlCaptionY) Then
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

Call clsError.DefErrPop
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Form_MouseMove")
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

Call clsError.DefErrPop
End Sub

Private Sub Form_Resize()
If (iNewLine) And (Me.Visible) Then
    CurrentX = wPara1.NlFlexBackY
    CurrentY = (wPara1.NlCaptionY - TextHeight(Caption)) / 2
    ForeColor = vbBlack
    Me.Print Caption
End If
End Sub

Private Sub nlcmdOk_Click()
Call cmdOk_Click
End Sub

Private Sub nlcmdEsc_Click()
Call cmdEsc_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If (iNewLine) Then
    If (KeyAscii = 13) Then
        Call nlcmdOk_Click
    ElseIf (KeyAscii = 27) Then
        Call nlcmdEsc_Click
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

Private Sub picTransparent(cPic As PictureBox)
  Dim lSkin As Long

  With cPic
'    .Visible = True
'    .Left = 0
'    .Top = 0
    .BorderStyle = 0
    .AutoRedraw = True
    .AutoSize = True
    lSkin = picTranz(cPic)
    Call SetWindowRgn(cPic.hWnd, lSkin, True)
  End With
End Sub

Private Function picTranz(cPic As PictureBox) As Long
  Dim lHoch As Long
  Dim lBreit As Long
  Dim lTemp As Long
  Dim lSkin As Long
  Dim lStart As Long
  Dim lZeile As Long
  Dim lSpalte As Long
  Dim lBackColor As Long

  lSkin = CreateRectRgn(0, 0, 0, 0)

  With cPic
    lHoch = .Height / Screen.TwipsPerPixelY
    lBreit = .Width / Screen.TwipsPerPixelX

    lBackColor = GetPixel(.hdc, 0, 0)

    For lZeile = 0 To lHoch - 1
      lSpalte = 0
      Do While lSpalte < lBreit
        Do While lSpalte < lBreit And _
              GetPixel(.hdc, lSpalte, lZeile) = lBackColor
          lSpalte = lSpalte + 1
        Loop

        If lSpalte < lBreit Then
          lStart = lSpalte
          Do While lSpalte < lBreit And _
                GetPixel(.hdc, lSpalte, lZeile) <> lBackColor
            lSpalte = lSpalte + 1
          Loop

          If lSpalte > lBreit Then lSpalte = lBreit
          lTemp = _
              CreateRectRgn(lStart, lZeile, lSpalte, lZeile + 1)
          Call CombineRgn(lSkin, lSkin, lTemp, 2)
          Call DeleteObject(lTemp)
        End If
      Loop
    Next lZeile
  End With

  picTranz = lSkin
End Function




