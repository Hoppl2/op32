Attribute VB_Name = "modAusdruck"
Option Explicit

Private Const DefErrModul = "AUSDRUCK.BAS"
Sub DruckZeile(ZeilenText$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("DruckZeile")
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
Dim i%, ind%, X%, iItalic%, iBold%
Dim h$, tx$

h$ = ZeilenText$

For i% = 0 To (AnzDruckSpalten% - 1)
    ind% = InStr(h$, vbTab)
    tx$ = Left$(h$, ind% - 1)
    h$ = Mid$(h$, ind% + 1)
    
    If (DruckSpalte(i%).Ausrichtung = "L") Then
        Do
            If (Len(tx$) <= 1) Then Exit Do
            If (Printer.TextWidth(tx$) <= DruckSpalte(i%).BreiteX) Then Exit Do
            tx$ = Left$(tx$, Len(tx$) - 1)
        Loop
        X% = DruckSpalte(i%).StartX
    Else
        X% = DruckSpalte(i%).StartX + DruckSpalte(i%).BreiteX - Printer.TextWidth(tx$)
    End If
    Printer.CurrentX = X%
    
    iItalic% = Printer.Font.Italic
    iBold% = Printer.Font.Bold
    If (DruckSpalte(i%).Attrib = 1) Then
        Printer.Font.Name = "Times New Roman"
        Printer.Font.Italic = True
    End If
    If (DruckSpalte(i%).Attrib = 2) Then
        Printer.Font.Name = "Symbol"
        Printer.Font.Bold = True
    End If
    
    Printer.Print tx$;
    
    If (DruckSpalte(i%).Attrib) Then
        Printer.Font.Name = "Arial"
        Printer.Font.Bold = iBold%
        Printer.Font.Italic = iItalic%
    End If
Next i%
Printer.Print " "
                
Call DefErrPop
End Sub


