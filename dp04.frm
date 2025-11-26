VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmDp04 
   Caption         =   "Strichcode zuordnen"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6225
   Begin VB.Timer tmrAutomatik 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3240
      Top             =   1080
   End
   Begin MSCommLib.MSComm comSenden 
      Left            =   4320
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327681
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   450
      Left            =   2640
      TabIndex        =   1
      Top             =   3600
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid flxDp04 
      Height          =   2700
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   4763
      _Version        =   65541
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483633
      BackColorBkg    =   -2147483633
      FocusRect       =   0
      GridLines       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmDp04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DP04_WARTEN = 3

Dim Dp04Deb%
Dim Dp04Para$

Private Const DefErrModul = "DP04.FRM"


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

Call wPara1.InitFont(Me)

Caption = "Kommunikation mit Lesestift"

Call clsOpTool.ModemParameter("DP-0", Dp04Para$, False)

With flxDp04
    .Rows = 1
    .Cols = 1
    .FixedRows = 0
    
    .Top = wPara1.TitelY
    .Left = wPara1.LinksX
    .Height = .RowHeight(0) * 11 + 90
    
    .ColWidth(0) = TextWidth(String(25, "9"))
    
    .Width = .ColWidth(0) + 90
    .Rows = 0
End With

Font.Bold = False   ' True

cmdEsc.Top = flxDp04.Top + flxDp04.Height + 150

Me.Width = flxDp04.Left + flxDp04.Width + 2 * wPara1.LinksX

cmdEsc.Width = wPara1.ButtonX
cmdEsc.Height = wPara1.ButtonY
cmdEsc.Left = (Me.Width - cmdEsc.Width) / 2

Font.Name = wPara1.FontName(1)
Font.Size = wPara1.FontSize(1)

Me.Height = cmdEsc.Top + cmdEsc.Height + wPara1.TitelY + 90 + wPara1.FrmCaptionHeight

Me.Left = ProjektForm.Left + (ProjektForm.Width - Me.Width) / 2
Me.Top = ProjektForm.Top + (ProjektForm.Height - Me.Height) / 2

tmrAutomatik.Enabled = True

Call clsError.DefErrPop
End Sub

Private Sub tmrAutomatik_Timer()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("tmrAutomatik_Timer")
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
Dim erg%
Dim s$

tmrAutomatik.Enabled = False
erg% = clsOpTool.OpenCom(Me, Dp04Para$)
Dp04Deb% = clsDat.FileOpen("winwawi.dp0", "O")
If (erg%) Then
    s$ = String(19, "0") + vbCr
    erg% = Dp04Send%(s$)
    Do
        erg% = Dp04Receive%(s$)
        If (erg% = False) Then Exit Do
    Loop
    comSenden.PortOpen = False
End If
Close (Dp04Deb%)
Unload Me

Call clsError.DefErrPop
End Sub

Function Dp04Send%(SendStr$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Dp04Send%")
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
Dim ret%
Dim l$, deb$

ret% = True
comSenden.Output = SendStr$
deb$ = "> " + SendStr$: Call StatusZeile(deb$)

Dp04Send% = ret%

Call clsError.DefErrPop
End Function

Function Dp04Receive%(RecStr$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("Dp04Receive%")
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
Dim ok%, l%
Dim char$, bcc$, deb$
Dim timeranf
Dim vchar As Variant
Dim bchar() As Byte

Dp04Receive% = False

RecStr$ = ""
ok% = 0
timeranf = Timer
While (ok% = 0)
    
    If (Timer - timeranf > DP04_WARTEN) Then ok% = 2
    
    l% = comSenden.InBufferCount
    If (l% > 0) Then
        char$ = comSenden.Input
'        deb$ = "< " + char$ + Str$(Asc(char$)): Call StatusZeile(deb$)
        RecStr$ = RecStr$ + char$
        If (char$ = vbCr) Then ok% = 1
    Else
        DoEvents
'        If (BestSendenAbbruch% = True) Then
'            ok% = 2
'        End If
    End If
Wend
If (ok% = 1) Then
    deb$ = "< " + RecStr$: Call StatusZeile(deb$)
    ok% = Dp04Send(RecStr$)
    Dp04Receive% = True
End If

Call clsError.DefErrPop
End Function

Sub StatusZeile(h$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call clsError.DefErrFnc("StatusZeile")
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
Dim i%, l%
Dim h2$, ch$

h2$ = ""

l% = Len(h$)
For i% = 1 To l%
    ch$ = Mid$(h$, i%, 1)
    If ch$ = Chr$(0) Then           'GS 3.11.00 DP04 spukt gelegentlich Chr(0) statt blank ->
                                    'LINE INPUT bricht dort ab -> ersetzen durch Blank
        Mid$(h$, i%, 1) = " "
    End If
    If (Asc(ch$) < 32) Then
        ch$ = "(" + Mid$(Str$(Asc(ch$)), 2) + ")"
    End If
    h2$ = h2$ + ch$
Next i%

With flxDp04
    .AddItem h2$
    If (.Rows > 10) Then
        .TopRow = .Rows - 10
        If (.FixedRows > 0) Then
            .TopRow = .TopRow + .FixedRows
        End If
    End If
End With

Print #Dp04Deb%, h$
'If (SendeLog%) Then
'    h2$ = h$
'    Call CharToOem(h2$, h2$)
'    Print #LEITUNGBUCH%, Format(Now, "HHMMSS ") + h2$
'End If

DoEvents
Call clsError.DefErrPop
End Sub


