VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLockTest2
   Caption         =   "LockTest2"
   ClientHeight    =   5715
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   7440
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   735
      Left            =   2280
      TabIndex        =   4
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdUnLock 
      Caption         =   "UnLock"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdLock 
      Caption         =   "Lock"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   3840
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid flxLockTest 
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5318
      _Version        =   65541
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
End
Attribute VB_Name = "frmLockTest2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefErrModul = "LOCKTEST2.FRM"

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
Unload Me
Call DefErrPop
End Sub

Private Sub cmdLock_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdLock_Click")
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
ww.SatzLock (1)
flxLockTest.AddItem "Lock"
Call DefErrPop
End Sub

Private Sub cmdRead_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdRead_Click")
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
ww.GetRecord (1)
flxLockTest.AddItem "Read"
Call DefErrPop
End Sub

Private Sub cmdUnLock_Click()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("cmdUnLock_Click")
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
ww.SatzUnLock (1)
flxLockTest.AddItem "UnLock"
Call DefErrPop
End Sub
