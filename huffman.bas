Attribute VB_Name = "modHuffman"
Option Explicit

Type HuffmanNodes
    Left As Integer
    Id As Integer
    Right As Integer
End Type

Dim Nodes() As HuffmanNodes
Dim AnzNodes%
Public KompBits%(8)
Public HuffmanInit%

Private Const DefErrModul = "huffman.bas"

Function EinlesenHuffmanTree%()
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EinlesenHuffmanTree%")
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
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim HUFFTREE%, i%, s$, erg%

KompBits%(0) = &H80
KompBits%(1) = &H40
KompBits%(2) = &H20
KompBits%(3) = &H10
KompBits%(4) = &H8
KompBits%(5) = &H4
KompBits%(6) = &H2
KompBits%(7) = &H1

s$ = para.CdLw + ":\cdinfo\hufftree.dat"
'HUFFTREE% = FileOpen%(s$, "R")
'If (HUFFTREE% = False) Then
If (Dir$(s$) = "") Then
'    erg% = MsgBox("Datei " + s$ + " ist nicht auf der Plattenkopie vorhanden !", vbCritical)
    EinlesenHuffmanTree% = False
    Call DefErrPop: Exit Function
'    End
End If
HUFFTREE% = FileOpen%(s$, "R")
AnzNodes% = LOF(HUFFTREE%) / 6
ReDim Nodes(AnzNodes%)
For i% = 1 To AnzNodes%
    Get HUFFTREE%, , Nodes(i% - 1)
Next i%
Close #HUFFTREE%

EinlesenHuffmanTree% = True
Call DefErrPop
End Function

Function InBit%(Handle%)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("InBit%")
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
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Static OutBytes$
Static off%, IndPos%
If (HuffmanInit%) Then
    OutBytes$ = String$(100, 0)
    off% = 0
    IndPos% = 99
    HuffmanInit% = False
End If
If (off% = 0) Then
    IndPos% = IndPos% + 1
    If (IndPos% = 100) Then
        Get Handle%, , OutBytes$
        IndPos% = 0
    End If
End If
InBit% = Asc(Mid$(OutBytes$, IndPos% + 1, 1)) And KompBits%(off%)
off% = (off% + 1) Mod 8
Call DefErrPop
End Function

Sub Decomp(Handle%, h$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("Decomp")
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
Call DefErrAbort
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim NodePtr%, ch%

h$ = ""
Do
    NodePtr% = 0
    Do
        If (Nodes(NodePtr%).Left >= 0) Or (Nodes(NodePtr%).Right >= 0) Then
            If (InBit%(Handle%)) Then
                NodePtr% = Nodes(NodePtr%).Left
            Else
                NodePtr% = Nodes(NodePtr%).Right
            End If
        Else
            Exit Do
        End If
    Loop
    ch% = Nodes(NodePtr%).Id
    If ((ch% = 10) Or (ch% = 256)) Then
        Exit Do
    ElseIf (ch% <> 13) Then
        h$ = h$ + Chr$(ch%)
    End If
Loop
h$ = RTrim$(h$)
Call OemToChar(h$, h$)
Call DefErrPop
End Sub
