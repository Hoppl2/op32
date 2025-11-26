VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   7440
   Begin MSFlexGridLib.MSFlexGrid flx1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7435
      _Version        =   65541
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Sub BestellVorschlag(Optional HintergrundAktiv% = True)
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Call DefErrFnc("BestellVorschlag")
'On Error GoTo DefErr
'GoTo DefErrEnd
'DefErr:
'Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
'Case vbRetry
'  Resume
'Case vbIgnore
'  Resume Next
'End Select
'End
'DefErrEnd:
''DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Dim i%, sMax%, VonSatz%, BisSatz%
'Dim Prozent!, StartZeit!, Dauer!, GesamtDauer!, RestDauer!
'Dim lSatz&
'Dim h$
'
'AnzArtikelDazu% = 0
'
'h$ = Format(Now, "DDMMYY")
'xheute% = iDate(h$)
'
'Call EinlesenWÜ
'
'If (para.BVRetour = "J") Then
'    Call EinlesenNB
'End If
'
'Call EinlesenBestellung
'
'If (InStr(para.Benutz, "Y") > 0) Then
'    Call AbholerTesten
'    AbholerMenge% = 0 'damit normale Bestellung nicht falsch wird
'End If
'
'Call ass.GetRecord(1)
'sMax% = ass.erstmax
'
'If (sMax% <= 0) Then Call DefErrPop: Exit Sub
'
'If (HintergrundAktiv%) Then
'
'    VonSatz% = HintergrundSsatz%    '+ 1
'    BisSatz% = VonSatz% + HintergrundAnz% - 1
''    If (VonSatz% = 1) Then VonSatz% = 2
'    If (VonSatz% = 0) Then VonSatz% = 1
'
'    For ssatz% = VonSatz% To BisSatz%
'        If (ssatz% > sMax%) Then Exit For
'
'        lSatz& = ssatz%
'        Call ass.GetRecord(lSatz& + 1)
'
'        If (InStr(AbholerPzns$, " " + ass.pzn) = 0) Then
'            Call ArtikelInBestellung
'        End If
'    Next ssatz%
'
'    HintergrundSsatz% = HintergrundSsatz% + HintergrundAnz%
'    If (HintergrundSsatz% > sMax%) Then HintergrundSsatz% = 0
'
'Else
'
'    BestvorsAbbruch% = False
'
'    frmBestVors!prgBestVors.Max = sMax%
'    StartZeit! = Timer
'
'    ass.GetRecord (1)
'
'    For ssatz% = 1 To sMax%
'        ass.GetRecord
'
'        If (InStr(AbholerPzns$, " " + ass.pzn) = 0) Then
'            Call ArtikelInBestellung
'        End If
'
'        If (ssatz% Mod 100 = 0) Then
'            frmBestVors!lblBestVorsStatusWert(0).Caption = ssatz%
'            frmBestVors!lblBestVorsStatusWert(1).Caption = AnzArtikelDazu%
'            Dauer! = Timer - StartZeit!
'            frmBestVors!lblBestVorsDauerWert(0).Caption = Format$(Dauer! \ 60, "##0") + ":" + Format$(Dauer! Mod 60, "00")
'            Prozent! = (ssatz% / sMax%) * 100!
'            If (Prozent! > 0) Then
'                GesamtDauer! = (Dauer! / Prozent!) * 100!
'            Else
'                GesamtDauer! = Dauer!
'            End If
'            RestDauer! = GesamtDauer! - Dauer!
'            frmBestVors!lblBestVorsDauerWert(1).Caption = Format$(RestDauer! \ 60, "##0") + ":" + Format$(RestDauer! Mod 60, "00")
'            frmBestVors!prgBestVors.Value = ssatz%
'            frmBestVors!lblBestVorsProzent.Caption = Format$(Prozent!, "##0") + " %"
'
'            h$ = Format$(Prozent!, "##0") + " %"
'            With frmBestVors!picBestVorsProgress
'                .Cls
'                .CurrentX = (.ScaleWidth - .TextWidth(h$)) \ 2
'                .CurrentY = (.ScaleHeight - .TextHeight(h$)) \ 2
'                frmBestVors!picBestVorsProgress.Print h$
'                frmBestVors!picBestVorsProgress.Line (0, 0)-((Prozent! * .ScaleWidth) \ 100, .ScaleHeight), vbHighlight, BF
''                Call BitBlt(.hdc, 0, 0, .ScaleWidth, .ScaleHeight, .hdc, 0, 0, &HCC0020)
'            End With
'
'            DoEvents
'            If (BestvorsAbbruch% = True) Then
'                Exit For
'            End If
'        End If
'    Next ssatz%
'
'End If
'
'Call DefErrPop
'End Sub

Private Sub Form_Load()

ChDir "\vb5src\winwawi"

With flx1
    .Rows = 2
    .Cols = 3
    .FixedRows = 1
    .FormatString = "Dateiname|SubName|Int.Name"
    .ColWidth(0) = .Width / 3
    .ColWidth(1) = .ColWidth(0)
    .ColWidth(2) = .ColWidth(0)
    .Rows = 1
End With

Call ErrorCheck("*.bas")
Call ErrorCheck("*.frm")
Call ErrorCheck("*.cls")

End Sub

Sub ErrorCheck(path$)
Dim DATEI%, ind%
Dim h$, s$, SubName$

DATEI% = FreeFile

h$ = CurDir$

h$ = Dir(path$)
While (h$ <> "")
    h$ = UCase(h$)
    Open h$ For Input As #DATEI%
    While Not (EOF(DATEI%))
        Line Input #DATEI%, s$
        
        s$ = UCase(s$)
        
'        ind% = InStr(s$, "PRIVATE CONST DEFERRMODUL")
'        If (ind% > 0) Then
'            ind% = InStr(s$, Chr$(34))
'            s$ = Mid$(s$, ind% + 1)
'            ind% = InStr(s$, Chr$(34))
'            s$ = Left$(s$, ind% - 1)
'            If (s$ <> h$) Then
'                flx1.AddItem h$ + vbTab + s$
'            End If
'        End If

        ind% = 0
        If (Left$(s$, 3) = "SUB") Then
            ind% = 3
        ElseIf (Left$(s$, 11) = "PRIVATE SUB") Then
            ind% = 11
        ElseIf (Left$(s$, 10) = "PUBLIC SUB") Then
            ind% = 10
        ElseIf (Left$(s$, 8) = "FUNCTION") Then
            ind% = 8
        ElseIf (Left$(s$, 16) = "PRIVATE FUNCTION") Then
            ind% = 16
        ElseIf (Left$(s$, 15) = "PUBLIC FUNCTION") Then
            ind% = 15
        End If
        
        If (ind% > 0) Then
            s$ = Trim(Mid$(s$, ind% + 1))
            ind% = InStr(s$, "(")
            If (ind% > 0) Then
                s$ = Left$(s$, ind% - 1)
            End If
            SubName$ = s$
'            flx1.AddItem h$ + vbTab + s$
        End If
    
        If (Left$(s$, 14) = "CALL DEFERRFNC") Or (Left$(s$, 23) = "CALL CLSERROR.DEFERRFNC") Then
            ind% = InStr(s$, Chr$(34))
            s$ = Mid$(s$, ind% + 1)
            ind% = InStr(s$, Chr$(34))
            s$ = Left$(s$, ind% - 1)
            If (s$ <> SubName$) Then
                flx1.AddItem h$ + vbTab + SubName$ + vbTab + s$
            End If
        End If
        
        ind% = InStr(s$, "EXIT SUB")
        If (ind% = 0) Then
            ind% = InStr(s$, "EXIT FUNCTION")
        End If
        If (ind% > 0) Then
            s$ = Trim(Left$(s$, ind% - 1))
            If (Right$(s$, 10) <> "DEFERRPOP:") Then
                flx1.AddItem h$ + vbTab + SubName$ + vbTab + s$
            End If
        End If
        
    Wend
    Close #DATEI%
    h$ = Dir
Wend

End Sub

