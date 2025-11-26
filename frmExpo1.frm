VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "EXPO - Apothekenbesetzung"
   ClientHeight    =   6540
   ClientLeft      =   -750
   ClientTop       =   1455
   ClientWidth     =   10320
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   10320
   Begin VB.CommandButton cmd2 
      Caption         =   "Soll"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "O K"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   2
      Top             =   4560
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid flx1 
      Height          =   4215
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   7435
      _Version        =   65541
      FocusRect       =   0
      HighLight       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flx2 
      Height          =   2055
      Left            =   5880
      TabIndex        =   1
      Top             =   4080
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3625
      _Version        =   65541
      FocusRect       =   0
      HighLight       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmd1_Click()
Dim h$

With flx1
    .SetFocus
    h$ = RTrim$(.TextMatrix(.Row, .Col))
    If (h$ = "") Then
        h$ = flx2.Text
    Else
        h$ = ""
    End If
    .TextMatrix(.Row, .Col) = h$
End With

End Sub

Private Sub cmd2_Click()

Form2.Show 1
Call UpdateFlex
flx1.SetFocus
End Sub

Private Sub flx2_gotFocus()

flx2.CellFontBold = False
End Sub

Private Sub flx2_LostFocus()

flx2.CellFontBold = True
End Sub

Private Sub Form_Load()
Dim i%, j%

st%(0) = 3
st%(1) = 4
st%(2) = 5
st%(3) = 5
st%(4) = 5
st%(5) = 5
st%(6) = 3
st%(7) = 3
st%(8) = 4
st%(9) = 5

With flx1
    .FormatString = "<8|<9|<10|<11|<12|<13|<14|<15|<16|<17"
    .Cols = 10
    For i% = 0 To 9
        .ColWidth(i%) = TextWidth("WWWWWWW")
    Next i%
    .Width = .ColWidth(0) * 10 + 120
    .Rows = 11
    .Height = .RowHeight(0) * 11 + 90
    .FixedRows = 1
    .FixedCols = 0
    
    Call UpdateFlex
    
End With

With flx2
    .Cols = 1
    .Rows = 5
    .ColWidth(0) = TextWidth("WWWWWWWWWWW")
    .Width = .ColWidth(0) + 120
    .Height = .RowHeight(0) * 5 + 90
    .FixedRows = 0
    .FixedCols = 0
    .Left = flx1.Left
    .Top = flx1.Top + flx1.Height + 150
    .TextMatrix(0, 0) = "Anna"
    .TextMatrix(1, 0) = "Bertha"
    .TextMatrix(2, 0) = "Carla"
    .TextMatrix(3, 0) = "Dora"
    .TextMatrix(4, 0) = "Emma"
End With

cmd1.Left = flx1.Left + flx1.Width - cmd1.Width
cmd2.Left = flx1.Left + flx1.Width - cmd2.Width
cmd2.Top = flx1.Top + flx1.Height + 150
cmd1.Top = cmd2.Top + cmd2.Height + 150

Me.WindowState = vbMaximized
End Sub


Sub UpdateFlex()
Dim i%, j%

With flx1
    .Redraw = False
    For i% = 0 To 9
        For j% = 0 To 9
            .Col = i%
            .Row = j% + 1
            If (j% >= st%(i%)) Then
                .CellBackColor = vbGrayText
            Else
                .CellBackColor = vbWhite
            End If
        Next j%
    Next i%
    
    .Row = 1
    .Col = 0
    .Redraw = True
End With

End Sub
