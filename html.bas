Attribute VB_Name = "modHTML"
Option Explicit

Type TabelleType
  element     As Variant
End Type

Type TabelleDefType
  breite      As Integer
  align       As Integer
End Type

Dim p2 As Integer
Dim PicHoehe  As Long
Dim SizeH(6)  As Integer
Dim SizeSt    As Integer
Dim SizeNp    As Integer
Dim FontSt    As String
Dim FontNp    As String
Dim HtmlText  As String
Dim frmX      As Form
Dim liste As Integer
Dim MaxPage As Integer
Dim vis As Boolean
Dim HeaderPage(16) As Integer
Dim HeaderPos(16) As Integer
Dim AnzHeader As Integer
Dim AktHeader As Integer
Dim SizeFaktor As Double

Private Const DefErrModul = "html.bas"
Sub HtmlDraw(frmX1 As Form, htm As String)

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HtmlDraw")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.source, Err.number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim i As Integer
Dim j As Integer

Dim x As Integer
Dim y As Integer
Dim tx As String
Dim tx1 As String
Dim htmltag As String
Dim body As Boolean
Dim pre As Boolean
Dim nl As Boolean
Dim inP As Boolean
Dim tabelle As Boolean
Dim tcol As Integer
Dim trow As Integer
Dim t() As TabelleType
Dim td() As TabelleDefType
Dim p As Integer

Dim Umlaute$

Umlaute$ = "äöüÄÖÜ"

Set frmX = frmX1
'frmX.picHtm1.Top = 0
'frmX.picHtm1.Left = 0
'frmX.picHtm1.Height = frmX.ScaleHeight
'frmX.picHtm1.Width = frmX.ScaleWidth

'If SizeFaktor = 0 Then
'  SizeFaktor = (frmKasse.Height / Screen.TwipsPerPixelX) / 480
'  SizeSt = Int(10 * SizeFaktor + 0.9)
'  'FontSt = "MS Sans Serif"
'  FontSt = "Arial"
'  FontNp = "Courier New"
'End If

SizeSt = Int(10 * BildFaktor! + 0.9)
'FontSt = "MS Sans Serif"
FontSt = "Arial"
FontNp = "Courier New"

SizeNp = SizeSt - 2
For i = 1 To 6
  SizeH(i) = SizeSt + (3 - i) * 2
Next i

AnzHeader = 0
AktHeader = 0

For p2 = 0 To MaxPage
  frmX.picHtm2(p2).Visible = False
Next p2

p2 = 0
vis = 0
p = 1
Call HtmlNewPage
'

Do While p <= Len(htm)
  x = InStr(p, htm, "<")
  If x > 0 Then
    y = InStr(x, htm, ">")
    If y > 0 Then
      tx = Mid(htm, p, x - p)
      htmltag = Mid(htm, x + 1, y - x - 1)
      p = y + 1
    Else
      tx = Mid(htm, p)
      p = Len(htm) + 1
    End If
  Else
    tx = Mid(htm, p)
    p = Len(htm) + 1
  End If
  
'  Debug.Print "<" + tx
  Do
    x = InStr(tx, "  ")
    If (x = 0) Then Exit Do
    tx = Left$(tx, x - 1) + Mid$(tx, x + 1)
  Loop
'  Debug.Print ">" + tx
  
  While tx <> ""
    tx1 = ""
  
    x = 1
    Do
      If InStr(" ,.;:-=!?" + Chr(13), Mid(tx, x, 1)) Then
        Exit Do
      Else
        x = x + 1
        If x = Len(tx) Then Exit Do
      End If
    Loop
    
    tx1 = Mid(tx, x + 1)
    tx = Left(tx, x)
    
    If Right(tx, 1) = Chr(13) Then
      Mid(tx, Len(tx), 1) = " "
      If pre Then nl = True
    End If
    
    Do
        x = InStr(tx, "uml;")
        If (x = 0) Then Exit Do
        y = InStr("aouAOU", Mid$(tx, x - 1, 1))
        If (y > 0) Then
            tx = Left$(tx, x - 3) + Mid$(Umlaute$, y, 1) + Mid$(tx, x + 4)
        Else
            tx = Left$(tx, x - 3) + "e" + Mid$(tx, x + 4)
        End If
    Loop
    Do
        x = InStr(tx, "&szlig;")
        If (x = 0) Then Exit Do
        tx = Left$(tx, x - 1) + "ß" + Mid$(tx, x + 7)
    Loop
    
    If tabelle Then
      t(trow, tcol).element = t(trow, tcol).element + tx
      If td(tcol).breite < frmX.picHtm2(p2).TextWidth(t(trow, tcol).element + " ") Then
        td(tcol).breite = frmX.picHtm2(p2).TextWidth(t(trow, tcol).element + " ")
      End If
      tx = ""
    Else
      If frmX.picHtm2(p2).TextWidth(tx) >= frmX.picHtm2(p2).ScaleWidth - frmX.picHtm2(p2).CurrentX Then
        Call HtmlNewLine
        tx = LTrim(tx)
      End If
      
      frmX.picHtm2(p2).Print tx;
      If nl Then
        Call HtmlNewLine
        nl = False
      End If
    End If
    tx = tx1
  Wend
  
  If htmltag <> "" Then
    Select Case UCase(htmltag)
    Case "BODY"
      body = True
    Case "/BODY"
      Call HtmlNewLine
      body = False
    Case "PRE", "TT", "CODE"
      pre = True
      Call HtmlNewLine
      frmX.picHtm2(p2).Font = FontNp
      frmX.picHtm2(p2).Font.Size = SizeNp
    Case "/PRE", "/TT", "/CODE"
      pre = False
      Call HtmlNewLine
      frmX.picHtm2(p2).Font = FontSt
      frmX.picHtm2(p2).Font.Size = SizeSt
    Case "HR"
      Call HtmlNewLine
      y = frmX.picHtm2(p2).CurrentY
      frmX.picHtm2(p2).Line (Screen.TwipsPerPixelX, y + frmX.picHtm2(p2).TextHeight("Äg") / 2)-(frmX.picHtm2(p2).ScaleWidth - Screen.TwipsPerPixelX, y + frmX.picHtm2(p2).TextHeight("Äg") / 2)
      frmX.picHtm2(p2).CurrentY = y
      Call HtmlNewLine
    Case "BR"
      Call HtmlNewLine
    Case "P"
      Call HtmlNewLine
      Call HtmlNewLine
      inP = True
    Case "/P"
      inP = False
    Case "/HL"
      Call HtmlNewLine
    Case "UL"
      liste = liste + 1
    Case "/UL"
      liste = liste - 1
    Case "OL"
      liste = liste + 1
    Case "/OL"
      liste = liste - 1
    Case "LI"
      Call HtmlNewLine
      x = frmX.picHtm2(p2).CurrentX
      frmX.picHtm2(p2).CurrentX = frmX.picHtm2(p2).CurrentX - frmX.picHtm2(p2).TextWidth("oo")
      frmX.picHtm2(p2).Print "o";
      frmX.picHtm2(p2).CurrentX = x
    Case "B"
      frmX.picHtm2(p2).Font.bold = True
    Case "/B"
      frmX.picHtm2(p2).Font.bold = False
    Case "U"
      frmX.picHtm2(p2).Font.Underline = True
    Case "/U"
      frmX.picHtm2(p2).Font.Underline = False
    Case "I"
      frmX.picHtm2(p2).Font.Italic = True
    Case "/I"
      frmX.picHtm2(p2).Font.Italic = False
    Case "STRONG"
      frmX.picHtm2(p2).Font.bold = True
    Case "/STRONG"
      frmX.picHtm2(p2).Font.bold = False
    Case "EM"
      frmX.picHtm2(p2).Font.Italic = True
    Case "/EM"
      frmX.picHtm2(p2).Font.Italic = False
    Case "/FONT"
      frmX.picHtm2(p2).Font.Size = SizeSt
    Case "TR"
      'table row
      trow = trow + 1
      tcol = 0
    Case Else
      If Len(htmltag) = 2 And UCase(Left(htmltag, 1)) = "H" Then
        i = val(Mid(htmltag, 2))
        If i < 1 Then i = 1
        If i > 6 Then i = 6
        If frmX.picHtm2(p2).CurrentY > 0 Then Call HtmlNewLine
        Call HtmlNewLine
        If AnzHeader < UBound(HeaderPage) Then
          AnzHeader = AnzHeader + 1
          HeaderPage(AnzHeader) = p2
          HeaderPos(AnzHeader) = frmX.picHtm2(p2).CurrentY
        End If
        frmX.picHtm2(p2).Font.Size = SizeH(i)
        frmX.picHtm2(p2).Font.bold = True
      ElseIf Len(htmltag) = 3 And UCase(Left(htmltag, 2)) = "/H" Then
        Call HtmlNewLine
        Call HtmlNewLine
        frmX.picHtm2(p2).Font.Size = SizeSt
        frmX.picHtm2(p2).Font.bold = False
      ElseIf UCase(Left(htmltag, 2)) = "OL" Then
        liste = liste + 1
      ElseIf UCase(Left(htmltag, 5)) = "TABLE" Then
        tabelle = True
        ReDim t(50, 1)
        ReDim td(1)
        tcol = 0
        trow = 0
      ElseIf UCase(Left(htmltag, 6)) = "/TABLE" Then
        'tabelle ausgeben
        Call HtmlNewLine
        For i = 1 To trow
          x = 0
          For j = 1 To UBound(td)
            If td(j).align = 1 Then
              frmX.picHtm2(p2).CurrentX = x + (td(j).breite - frmX.picHtm2(p2).TextWidth(t(i, j).element)) / 2
            ElseIf td(j).align = 2 Then
              frmX.picHtm2(p2).CurrentX = x + td(j).breite - frmX.picHtm2(p2).TextWidth(t(i, j).element)
            Else
              frmX.picHtm2(p2).CurrentX = x
            End If
            frmX.picHtm2(p2).Print t(i, j).element;
            x = x + td(j).breite
          Next j
          Call HtmlNewLine
        Next i
        tabelle = False
        trow = 0
        tcol = 0
      ElseIf UCase(Left(htmltag, 2)) = "TD" Then
        tcol = tcol + 1
        If UBound(td) < tcol Then
          ReDim Preserve t(50, tcol)
          ReDim Preserve td(tcol)
        End If
        If UCase(Mid(htmltag, 4, 6)) = "ALIGN=" Then
          If UCase(Mid(htmltag, 10, 6)) = "CENTER" Then
            td(tcol).align = 1
          ElseIf UCase(Mid(htmltag, 10, 5)) = "RIGHT" Then
            td(tcol).align = 2
          End If
        End If
      ElseIf UCase(Left(htmltag, 10)) = "FONT SIZE=" Then
        tx1 = Mid(htmltag, 11)
        If Left(tx1, 1) = Chr(34) Then tx1 = Mid(tx1, 2)
        If Right(tx1, 1) = Chr(34) Then tx1 = Left(tx1, Len(tx1) - 1)
        If InStr("+-", Left(tx1, 1)) > 0 Then
          frmX.picHtm2(p2).Font.Size = frmX.picHtm2(p2).Font.Size + val(tx1)
        Else
          frmX.picHtm2(p2).Font.Size = val(tx1)
        End If
      End If
    End Select
    
  End If
Loop

frmX.picHtm2(p2).Height = frmX.picHtm2(p2).CurrentY + (frmX.picHtm2(p2).Height - frmX.picHtm2(p2).ScaleHeight)

PicHoehe = 0
For i = 0 To p2
  PicHoehe = PicHoehe + frmX.picHtm2(i).Height
Next i

frmX.picHtm2(0).SetFocus

Call DefErrPop
End Sub

Sub HtmlNewLine()

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HtmlNewLine")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.source, Err.number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Static LeerZeile As Boolean

Dim nl As Boolean
Dim i As Integer
Dim t As Integer

If frmX.picHtm2(p2).CurrentX > 0 Then
  frmX.picHtm2(p2).Print
  LeerZeile = False
  nl = True
Else
  If Not LeerZeile Then
    frmX.picHtm2(p2).Print
  End If
  nl = True
  LeerZeile = True
End If

If nl And liste > 0 Then
  For i = 1 To liste
    t = t + frmX.picHtm2(p2).TextWidth("oooo")
  Next i
  frmX.picHtm2(p2).CurrentX = t
End If

If p2 = 0 And vis = 0 Then
  If frmX.picHtm2(p2).CurrentY > frmX.picHtm1.ScaleHeight Then
    frmX.picHtm2(p2).Refresh
    vis = True
  End If
End If

'If frmX.picHtm2(p2).CurrentY > 32000 Then
'  p2 = p2 + 1
'  Call HtmlNewPage
'End If

Call DefErrPop
End Sub

Sub HtmlNewPage()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("HtmlNewPage")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.source, Err.number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  
If p2 > 0 Then
  frmX.picHtm2(p2 - 1).Height = frmX.picHtm2(p2 - 1).CurrentY + (frmX.picHtm2(p2 - 1).Height - frmX.picHtm2(p2 - 1).ScaleHeight)
  If p2 > MaxPage Then
    Load frmX.picHtm2(p2)
    MaxPage = p2
  End If
End If

frmX.picHtm2(p2).Visible = False
If p2 = 0 Then
  frmX.picHtm2(p2).Top = 0
Else
  frmX.picHtm2(p2).Top = frmX.picHtm2(p2 - 1).Top + frmX.picHtm2(p2 - 1).Height
End If
frmX.picHtm2(p2).Left = 0
frmX.picHtm2(p2).Width = frmX.picHtm1.ScaleWidth
'frmX.picHtm2(p2).Height = 3276700
frmX.picHtm2(p2).Height = 100000

If p2 > 0 Then
  'letzte Einstellung übernehmen
  frmX.picHtm2(p2).Font = frmX.picHtm2(p2 - 1).Font
  frmX.picHtm2(p2).Font.Size = frmX.picHtm2(p2 - 1).Font.Size
  frmX.picHtm2(p2).Font.bold = frmX.picHtm2(p2 - 1).Font.bold
  frmX.picHtm2(p2).Font.Italic = frmX.picHtm2(p2 - 1).Font.Italic
  frmX.picHtm2(p2).Font.Underline = frmX.picHtm2(p2 - 1).Font.Underline
Else
  'Grundeinstellung
  frmX.picHtm2(p2).Font = FontSt
  frmX.picHtm2(p2).Font.Size = SizeSt
  frmX.picHtm2(p2).Font.bold = False
  frmX.picHtm2(p2).Font.Italic = False
  frmX.picHtm2(p2).Font.Underline = False
End If
frmX.picHtm2(p2).Cls
frmX.picHtm2(p2).Visible = True

Call DefErrPop
End Sub

Sub Html_ievent(ievent As String, text As String, KeyCode As Integer, Shift As Integer)

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Html_ievent")
Call DefErrMod(DefErrModul)
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
Select Case DefErrAnswer(Err.source, Err.number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim ReDraw As Boolean
Dim i As Integer
Dim yposi As Long

Select Case ievent
Case "KeyDown"
  Select Case KeyCode
  Case vbKeyAdd
    SizeSt = SizeSt + 1
    KeyCode = 0
  Case vbKeySubtract
    SizeSt = SizeSt - 1
    KeyCode = 0
  Case vbKeyHome
    KeyCode = 0
    If (Shift And vbCtrlMask) Then
      frmX.picHtm2(0).Top = 0
      ReDraw = True
    End If
  Case vbKeyEnd
    KeyCode = 0
    If (Shift And vbCtrlMask) Then
      frmX.picHtm2(0).Top = -PicHoehe + frmX.picHtm1.ScaleHeight
      ReDraw = True
    End If
  Case vbKeyUp
    KeyCode = 0
    If frmX.picHtm2(0).Top < 0 Then
      frmX.picHtm2(0).Top = frmX.picHtm2(0).Top + frmX.picHtm1.TextHeight("Äg")
    End If
    If frmX.picHtm2(0).Top > 0 Then frmX.picHtm2(0).Top = 0
    ReDraw = True
  Case vbKeyDown
    KeyCode = 0
    If frmX.picHtm2(p2).Top + frmX.picHtm2(p2).ScaleHeight > 0 Then
      frmX.picHtm2(0).Top = frmX.picHtm2(0).Top - frmX.picHtm1.TextHeight("Äg")
      ReDraw = True
    End If
  Case vbKeyPageUp
    KeyCode = 0
    If frmX.picHtm2(0).Top < 0 Then
      frmX.picHtm2(0).Top = frmX.picHtm2(0).Top + frmX.picHtm1.ScaleHeight
    End If
    If frmX.picHtm2(0).Top > 0 Then frmX.picHtm2(0).Top = 0
    ReDraw = True
  Case vbKeyPageDown
    KeyCode = 0
    If frmX.picHtm2(p2).Top + frmX.picHtm2(p2).ScaleHeight > 0 Then
      frmX.picHtm2(0).Top = frmX.picHtm2(0).Top - frmX.picHtm1.ScaleHeight
      ReDraw = True
    End If
  Case vbKeyI
    If (Shift And vbAltMask) > 0 Then
      KeyCode = 0
      If AnzHeader > 0 Then
        If AktHeader = 0 Then AktHeader = 1
        AktHeader = AktHeader + 1
        If AktHeader > AnzHeader Then AktHeader = 1
        yposi = 0
        For i = 1 To HeaderPage(AktHeader)
          yposi = yposi - frmX.picHtm2(i).Height
        Next i
        yposi = yposi - HeaderPos(AktHeader)
        frmX.picHtm2(0).Top = yposi
        ReDraw = True
      End If
    End If
  
  End Select
End Select

If ReDraw Then
  For i = 1 To p2
    frmX.picHtm2(i).Top = frmX.picHtm2(i - 1).Top + frmX.picHtm2(i - 1).Height
  Next i
End If

Call DefErrPop
End Sub

