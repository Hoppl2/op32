VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Sollbesetzung"
   ClientHeight    =   4755
   ClientLeft      =   1650
   ClientTop       =   1500
   ClientWidth     =   5085
   LinkTopic       =   "Form2"
   ScaleHeight     =   4755
   ScaleWidth      =   5085
   Begin VB.CommandButton cmd1 
      Cancel          =   -1  'True
      Caption         =   "O K"
      Default         =   -1  'True
      Height          =   615
      Left            =   1680
      TabIndex        =   20
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   3960
      TabIndex        =   19
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   3960
      TabIndex        =   17
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   3960
      TabIndex        =   15
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   3960
      TabIndex        =   13
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   3960
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1200
      TabIndex        =   9
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1200
      TabIndex        =   7
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   5
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   195
      Width           =   495
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Zentriert
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   3000
      TabIndex        =   18
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Zentriert
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   3000
      TabIndex        =   16
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Zentriert
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   3000
      TabIndex        =   14
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Zentriert
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3000
      TabIndex        =   12
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Zentriert
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   10
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Zentriert
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Zentriert
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Zentriert
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Zentriert
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Zentriert
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd1_Click()
Dim i%

For i% = 0 To 9
    st%(i%) = Val(txt1(i%).Text)
Next i%

Unload Me

End Sub

Private Sub Form_Load()
Dim i%

For i% = 0 To 9
    txt1(i%).Text = st%(i%)
Next i%
End Sub

Private Sub txt1_GotFocus(Index As Integer)

With txt1(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With

End Sub
