VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Umsatztabelle"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdSchwell 
      Caption         =   "Schwellwert (F2)"
      Height          =   540
      Left            =   5355
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5355
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Beenden (ESC)"
      Height          =   540
      Left            =   7140
      TabIndex        =   4
      Top             =   5355
      Width           =   1695
   End
   Begin VB.Frame Frame 
      Caption         =   "Frame"
      Height          =   5160
      Index           =   2
      Left            =   5985
      TabIndex        =   7
      Top             =   105
      Width           =   2850
      Begin VB.TextBox Schwell 
         BackColor       =   &H8000000F&
         Height          =   330
         Index           =   2
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   4350
         Width           =   1275
      End
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   4005
         Index           =   2
         Left            =   105
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   315
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   7064
         _Version        =   65541
         Enabled         =   0   'False
         ScrollBars      =   0
      End
      Begin VB.Label PLabel 
         Alignment       =   1  'Rechts
         Caption         =   "Prognose:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   105
         TabIndex        =   19
         Top             =   4725
         Width           =   1380
      End
      Begin VB.Label Prognose 
         Alignment       =   1  'Rechts
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1470
         TabIndex        =   16
         Top             =   4725
         Width           =   1275
      End
      Begin VB.Label SLabel 
         Alignment       =   1  'Rechts
         Caption         =   "Schwellwert :"
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   12
         Top             =   4410
         Width           =   1275
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Frame"
      Height          =   5160
      Index           =   1
      Left            =   3045
      TabIndex        =   6
      Top             =   105
      Width           =   2850
      Begin VB.TextBox Schwell 
         BackColor       =   &H8000000F&
         Height          =   330
         Index           =   1
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   4350
         Width           =   1275
      End
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   4005
         Index           =   1
         Left            =   105
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   315
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   7064
         _Version        =   65541
         Enabled         =   0   'False
         ScrollBars      =   0
      End
      Begin VB.Label PLabel 
         Alignment       =   1  'Rechts
         Caption         =   "Prognose:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   105
         TabIndex        =   18
         Top             =   4725
         Width           =   1380
      End
      Begin VB.Label Prognose 
         Alignment       =   1  'Rechts
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1470
         TabIndex        =   15
         Top             =   4725
         Width           =   1275
      End
      Begin VB.Label SLabel 
         Alignment       =   1  'Rechts
         Caption         =   "Schwellwert :"
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   11
         Top             =   4410
         Width           =   1275
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Frame"
      Height          =   5160
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   2850
      Begin VB.TextBox Schwell 
         BackColor       =   &H8000000F&
         Height          =   330
         Index           =   0
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   4350
         Width           =   1275
      End
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   4005
         Index           =   0
         Left            =   105
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   315
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   7064
         _Version        =   65541
         Enabled         =   0   'False
         ScrollBars      =   0
      End
      Begin VB.Label PLabel 
         Alignment       =   1  'Rechts
         Caption         =   "Prognose:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   105
         TabIndex        =   17
         Top             =   4725
         Width           =   1275
      End
      Begin VB.Label Prognose 
         Alignment       =   1  'Rechts
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1470
         TabIndex        =   14
         Top             =   4725
         Width           =   1275
      End
      Begin VB.Label SLabel 
         Alignment       =   1  'Rechts
         Caption         =   "Schwellwert :"
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   10
         Top             =   4410
         Width           =   1275
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAX_LIEF = 200

Dim aLief As Integer
Dim wDir As String
Dim lArray(1 To MAX_LIEF, 1 To 16) As Double
Dim l(1 To 3) As Integer
Dim aDatum As Date
Dim EditMode As Boolean



Private Sub cmdCancel_Click()
Unload Me
End
End Sub

Private Sub cmdCancel_GotFocus()
aLief = -1
SetTableFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i, ret As Integer

If EditMode = True Then
    Select Case KeyCode
    Case vbKeyEscape
        Schwell(aLief).Text = lArray(l(aLief + 1), 16)
        Schwell(aLief).BackColor = vbButtonFace
        Schwell(aLief).Locked = True
        Schwell(aLief).SelStart = 0
        Schwell(aLief).SelLength = 0
        EditMode = False
        cmdCancel.Enabled = True
        cmdSchwell.Enabled = True
        cmdCancel.Cancel = True
        For i = 0 To 2
            If i <> aLief Then
                Schwell(i).TabStop = True
            End If
        Next i
        DrawTable
    Case vbKeyReturn
        If Val(Schwell(aLief).Text) < 1000 And Val(Schwell(aLief).Text) > 0 Then
            ret = MsgBox("Wollen Sie bei diesem Lieferanten" + vbCrLf _
                + "wirklich nur " + Schwell(aLief).Text + " DM" + vbCrLf _
                + "umsetzen ???", vbYesNo + vbQuestion, "Sind Sie sicher ?")
        ElseIf Val(Schwell(aLief).Text) > 1000000 Then
            ret = MsgBox("Wollen Sie bei diesem Lieferanten" + vbCrLf _
                + "wirklich " + Schwell(aLief).Text + " DM" + vbCrLf _
                + "umsetzen ???", vbYesNo + vbQuestion, "Sind Sie sicher ?")
        Else
            ret = vbYes
        End If
        lArray(l(aLief + 1), 16) = Val(Schwell(aLief).Text)
        If ret = vbYes Then
            Schwell(aLief).BackColor = vbButtonFace
            Schwell(aLief).Locked = True
            Schwell(aLief).SelStart = 0
            Schwell(aLief).SelLength = 0
            EditMode = False
            cmdCancel.Enabled = True
            cmdSchwell.Enabled = True
            cmdCancel.Cancel = True
            For i = 0 To 2
                If i <> aLief Then
                    Schwell(i).TabStop = True
                End If
            Next i
            DrawTable
        Else
            Schwell(aLief).SelStart = 0
            Schwell(aLief).SelLength = 32
        End If
    Case Else
        'KeyCode = 0
    End Select
Else
    Select Case KeyCode
    Case vbKeyF2
        If aLief >= 0 And aLief <= 2 Then
            EditMode = True
            cmdCancel.Cancel = False
            cmdCancel.Enabled = False
            cmdSchwell.Enabled = False
            Schwell(aLief).Locked = False
            Schwell(aLief).BackColor = vbWindowBackground
            Schwell(aLief).SelStart = 0
            Schwell(aLief).SelLength = 32
            For i = 0 To 2
                If i <> aLief Then
                    Schwell(i).TabStop = False
                End If
            Next i
        End If
    Case vbKeyReturn
        ret = Val(InputBox("Lieferant", "Lieferant"))
        If ret > 0 And ret <= 200 Then
            l(aLief + 1) = ret
        End If
    Case vbKeyRight
        If aLief = 2 Then
            aLief = 0
        Else
            aLief = aLief + 1
        End If
    Case vbKeyLeft
        If aLief = 0 Then
            aLief = 2
        Else
            aLief = aLief - 1
        End If
    Case Else
        KeyCode = 0
    End Select
    DrawTable
End If
End Sub

Private Sub Form_Load()

Dim aFile As String
Dim ret As Integer
Dim buf As String
Dim buf1 As String * 8
Dim tDatum As Date
Dim tLief As Integer
Dim tValue As Double
Dim tmpDiff As Integer
Dim i, j As Integer
Dim aLief As Integer

wDir = "Z:\USER" 'der Einfachheit halber

aFile = Dir(wDir + "\WUMSATZ.DAT")

If aFile = "" Then
    MsgBox "Datei WUMSATZ.DAT konnte nicht gefunden werden !"
    End
End If

aFile = Dir(wDir + "\LSCHWELL.DAT")

If aFile = "" Then
    MsgBox "Datei LSCHWELL.DAT konnte nicht gefunden werden !"
    End
End If

aDatum = Format(Date, "DD.MM.YYYY")

For i = 0 To 2
    Flex(i).FormatString = ">Monat/Jahr|>Summe"
    Flex(i).ColWidth(0) = 1300
    Flex(i).ColWidth(1) = Flex(i).Width - Flex(i).ColWidth(0) - 90
    Flex(i).Height = 16 * Flex(i).RowHeight(0) + 90
    Schwell(i).Top = Flex(i).Top + Flex(i).Height + 50
    Schwell(i).Width = Flex(i).ColWidth(1) + 45
    Schwell(i).Left = Flex(i).Left + Flex(i).ColPos(1) + 45
    SLabel(i).Top = Schwell(i).Top + 0
Next i

aFile = wDir + "\WUMSATZ.DAT"

Open aFile For Input Access Read Shared As #1
Line Input #1, buf
l(1) = Val(Mid$(buf, 2, 3)) + 1
l(2) = Val(Mid$(buf, 5, 3)) + 1
l(3) = Val(Mid$(buf, 8, 3)) + 1

Do While Not EOF(1)
    Line Input #1, buf
    tDatum = DateValue(Left(buf, 2) + "." + Mid(buf, 3, 2) + "." + Mid(buf, 5, 2))
    tLief = Mid(buf, 8, 3)
    tValue = Val(Mid(buf, 11, 8) + "." + Mid(buf, 19, 2))
    tmpDiff = 13 - Abs(DateDiff("M", aDatum, tDatum))
    
    If tmpDiff > 0 Then
        lArray(tLief, tmpDiff + 1) = lArray(tLief, tmpDiff + 1) + tValue
        If Year(aDatum) = Year(tDatum) Then
            lArray(tLief, 15) = lArray(tLief, 15) + tValue
        Else
            lArray(tLief, 1) = lArray(tLief, 1) + tValue
        End If
    End If
Loop
Close

aFile = wDir + "\LSCHWELL.DAT"

Open aFile For Random Access Read Shared As #1 Len = 8
For i = 1 To MAX_LIEF
    Get #1, i, buf1
    lArray(i, 16) = Val(buf1)
Next i
Close
aLief = 0
Me.Show
DrawTable
End Sub

Sub DrawTable()
Dim i, j As Integer
Dim tmp As Double

Frame(0).Caption = "Lieferant " + Str(l(1))
Frame(1).Caption = "Lieferant " + Str(l(2))
Frame(2).Caption = "Lieferant " + Str(l(3))
For j = 0 To 2
    Flex(j).AddItem "gesamt Vorjahr" + vbTab + Format(lArray(l(j + 1), 1), "# ### ##0.00"), 1
    For i = 2 To 14
        Flex(j).AddItem Format(DateAdd("M", i - 14, aDatum), "MMMM YYYY") + vbTab + Format(lArray(l(j + 1), i), "# ### ##0.00"), i
    Next i
    Flex(j).AddItem "gesamt akt. Jahr" + vbTab + Format(lArray(l(j + 1), 15), "# ### ##0.00"), 15
    Schwell(j).Text = Format(lArray(l(j + 1), 16), "# ### ##0.00")
    tmp = CalcProg(l(j + 1))
    Prognose(j) = Format(tmp / 100, "# ##0 %")
Next j
SetTableFocus
End Sub

Sub SetTableFocus()
Dim j As Integer

For j = 0 To 2
    If j = aLief Then
        Flex(j).BackColor = vbHighlight
        Me.Schwell(j).SetFocus
    Else
        Flex(j).BackColor = vbWindowBackground
    End If
Next j

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim buf1 As String * 8
Dim i As Integer

wDir = "Z:\USER" 'der Einfachheit halber

Open wDir + "\LSCHWELL.DAT" For Random Access Write Lock Read Write As #1 Len = 8
For i = 1 To MAX_LIEF
    buf1 = Format(lArray(i, 16), "@@@@@@@@")
    Put #1, i, buf1
Next i
Close
MsgBox "speichern der Lieferantennummern fehlt noch"
End Sub

Private Sub Schwell_GotFocus(Index As Integer)
aLief = Index
SetTableFocus
End Sub

Function CalcProg(lief As Integer) As Double
Dim aDay, nDay, aMonat As Integer
Dim tmp As Double

If lArray(lief, 16) > 0 Then
    aDay = Day(Date)
    aMonat = Month(Date)
    Select Case aMonat
    Case 1, 3, 5, 7, 9, 10, 12
        nDay = 31
    Case 4, 6, 9, 11
        nDay = 30
    Case Else
        nDay = 28
    End Select
    tmp = (nDay * lArray(lief, 14)) / aDay
    tmp = (tmp * 100) / lArray(lief, 16)
    CalcProg = tmp
Else
    CalcProg = 0
End If
End Function
