VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "Info zu WinWaWi"
   ClientHeight    =   4425
   ClientLeft      =   2370
   ClientTop       =   3210
   ClientWidth     =   5610
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   5610
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   615
      Left            =   240
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   7
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&System-Info ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4020
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3540
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4020
      TabIndex        =   1
      Top             =   3060
      Width           =   1500
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   975
      Left            =   180
      ScaleHeight     =   915
      ScaleWidth      =   5235
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1680
      Width           =   5295
   End
   Begin VB.Label Label4 
      Caption         =   "Dieses Produkt wurde lizensiert für:"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   1380
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Copyright (c)  MediaLabs GmbH"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1020
      TabIndex        =   5
      Top             =   780
      Width           =   3915
   End
   Begin VB.Label Label2 
      Caption         =   "WinWaWi - Warenwirtschaft der Zukunft"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1020
      TabIndex        =   4
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   $"About.frx":0442
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   180
      TabIndex        =   3
      Top             =   3060
      Width           =   3735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   5460
      Y1              =   2895
      Y2              =   2895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   5460
      Y1              =   2880
      Y2              =   2880
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "ABOUT.FRM"

Private Sub Command1_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Command1_Click")
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

Private Sub Command2_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Command2_Click")
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
Dim x As Long

On Error Resume Next
x = Shell("\Programme\Gemeinsame Dateien\Microsoft Shared\MSINFO\msinfo32.exe", vbNormalFocus)

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
Me.Top = frmAction.Top + (frmAction.Height - Me.Height) / 2
Me.Left = frmAction.Left + (frmAction.Width - Me.Width) / 2
Me.Caption = "Info zur " + ProgrammNamen$(0)

'Label2.Caption = Label2.Caption + " " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
Label2.Caption = frmAction.Caption + " " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)

Picture2.Print
Picture2.Print " "; para.FISTAM(0)
Picture2.Print " "; para.FISTAM(1)

Picture1.Picture = frmAction.Icon

Call DefErrPop
End Sub
