Attribute VB_Name = "modWheelHook"
Option Explicit

Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "USER32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As Rect) As Long

Declare Function GetParent Lib "USER32" (ByVal hWnd As Long) As Long

Public Const GWL_WNDPROC = (-4)
Public Const WM_MOUSEWHEEL = &H20A
Public Const CB_GETDROPPEDSTATE = &H157


Private Const DefErrModul = "WHEELHOOK.BAS"

' Check Messages
' ================================================
Private Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WindowProc")
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
Dim MouseKeys As Long
Dim Rotation As Long
Dim xpos As Long
Dim Ypos As Long
Dim fFrm As Form

Select Case Lmsg

    Case WM_MOUSEWHEEL
        MouseKeys = wParam And 65535
        Rotation = wParam / 65536
        xpos = lParam And 65535
        Ypos = lParam / 65536
        
        Set fFrm = GetForm(Lwnd)
        If fFrm Is Nothing Then
            ' it's not a form
            If Not IsOver(Lwnd, xpos, Ypos) And IsOver(GetParent(Lwnd), xpos, Ypos) Then
                ' it's not over the control and is over the form,
                ' so fire mousewheel on form (if it's not a dropped down combo)
                If SendMessage(Lwnd, CB_GETDROPPEDSTATE, 0&, 0&) <> 1 Then
                    MouseWheel GetForm(GetParent(Lwnd)), MouseKeys, Rotation, xpos, Ypos
                    Call DefErrPop: Exit Function ' Discard scroll message to control
                End If
            End If
        Else
            ' it's a form so fire mousewheel
            If IsOver(fFrm.hWnd, xpos, Ypos) Then
                MouseWheel fFrm, MouseKeys, Rotation, xpos, Ypos
            End If
        End If
End Select

WindowProc = CallWindowProc(GetProp(Lwnd, "PrevWndProc"), Lwnd, Lmsg, wParam, lParam)

Call DefErrPop
End Function

Private Function GetForm(ByVal hWnd As Long) As Form
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("GetForm")
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
  
For Each GetForm In Forms
    If (GetForm.hWnd = hWnd) Then
        Call DefErrPop: Exit Function
    End If
Next GetForm
Set GetForm = Nothing

Call DefErrPop
End Function

' Here you can add scrolling support to controls that don't normally respond.
' This Sub could always be moved to a module to make scrollwheel behaviour
' generic across forms.
' ===========================================================================
Private Sub MouseWheel(frm As Form, ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal xpos As Long, ByVal Ypos As Long)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("MouseWheel")
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
Dim ctl As Control
Dim bHandled As Boolean
Dim bOver As Boolean

If (TypeOf frm.ActiveControl Is MSFlexGrid) Then
      FlexGridScroll frm.ActiveControl, MouseKeys, Rotation, xpos, Ypos
End If

'
'  For Each ctl In Controls
'    ' Is the mouse over the control
'    On Error Resume Next
'    bOver = (ctl.Visible And IsOver(ctl.hWnd, Xpos, Ypos))
'    On Error GoTo 0
'
'    If bOver Then
'      ' If so, respond accordingly
'      bHandled = True
'      Select Case True
'
'        Case TypeOf ctl Is MSFlexGrid
'          FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
'
''        Case TypeOf ctl Is PictureBox
''          PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
''
''        Case TypeOf ctl Is ListBox, TypeOf ctl Is TextBox, TypeOf ctl Is ComboBox
''          ' These controls already handle the mousewheel themselves, so allow them to:
''          If ctl.Enabled Then ctl.SetFocus
'
'        Case Else
'          bHandled = False
'
'      End Select
'      If bHandled Then Exit Sub
'    End If
'    bOver = False
'  Next ctl
'
'  ' Scroll was not handled by any controls, so treat as a general message send to the form
'  Me.Caption = "Form Scroll " & IIf(Rotation < 0, "Down", "Up")

Call DefErrPop
End Sub

    
' Window Checks
' ================================================
Private Function IsOver(ByVal hWnd As Long, ByVal lx As Long, ByVal lY As Long) As Boolean
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("IsOver")
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
Dim rectCtl As Rect

GetWindowRect hWnd, rectCtl
With rectCtl
    IsOver = (lx >= .Left And lx <= .Right And lY >= .Top And lY <= .Bottom)
End With

Call DefErrPop
End Function

' Control Specific Behaviour
' ================================================
Private Sub FlexGridScroll(ByRef FG As MSFlexGrid, ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal xpos As Long, ByVal Ypos As Long)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("FlexGridScroll")
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
Dim NewValue As Long
Dim Lstep As Single

On Error Resume Next
With FG
    Lstep = .Height / .RowHeight(0)
    Lstep = Int(Lstep)
    '  MsgBox (FG.Name + " MouseWheel " & IIf(Rotation < 0, "Down", "Up"))
    If .Rows < Lstep Then Exit Sub
    Do While Not (.RowIsVisible(.TopRow + Lstep))
        Lstep = Lstep - 1
    Loop
    If Rotation > 0 Then
        NewValue = .TopRow - Lstep
        If NewValue < 1 Then
            NewValue = 1
        End If
    Else
        NewValue = .TopRow + Lstep
        If NewValue > .Rows - 1 Then
            NewValue = .Rows - 1
        End If
    End If
    .TopRow = NewValue
End With

Call DefErrPop
End Sub

'Public Sub PictureBoxZoom(ByRef picBox As PictureBox, ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
'  picBox.Cls
'  picBox.Print "MouseWheel " & IIf(Rotation < 0, "Down", "Up")
'  MsgBox (picBox.Name + " MouseWheel " & IIf(Rotation < 0, "Down", "Up"))
'End Sub



' Hook / UnHook
' ================================================
Public Sub WheelHook(ByVal hWnd As Long)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WheelHook")
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

If (para.Newline) And (wpara.EntwicklungsUmgebung = 0) Then
    On Error Resume Next
    SetProp hWnd, "PrevWndProc", SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
End If

Call DefErrPop
End Sub

Public Sub WheelUnHook(ByVal hWnd As Long)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("WheelUnHook")
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

If (para.Newline) And (wpara.EntwicklungsUmgebung = 0) Then
    On Error Resume Next
    SetWindowLong hWnd, GWL_WNDPROC, GetProp(hWnd, "PrevWndProc")
    RemoveProp hWnd, "PrevWndProc"
End If

Call DefErrPop
End Sub



