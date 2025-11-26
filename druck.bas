Attribute VB_Name = "Druck"
Option Explicit

Public Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal flags As Long, ByVal name As String, ByVal Level As Long, pPrinterEnum As Long, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Declare Function PtrToStr Lib "Kernel32" Alias "lstrcpyA" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Declare Function StrLen Lib "Kernel32" Alias "lstrlenA" (ByVal Ptr As Long) As Long

Public Const PRINTER_ENUM_CONNECTIONS = &H4
Public Const PRINTER_ENUM_LOCAL = &H2
Public Const PRINTER_ENUM_NETWORK = &H40
Public Const PRINTER_ENUM_REMOTE = &H10
Public Const PRINTER_ENUM_SHARED = &H20


Public Druckernamen() As String

Sub EnumeratePrinters1(p As Long)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EnumeratePrinters1")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Dim Success As Boolean, cbRequired As Long, cbBuffer As Long
Dim Buffer() As Long, nEntries As Long
Dim i As Long, PFlags As Long, PDesc As String, PName As String
Dim PComment As String, Temp As Long

cbBuffer = 3072
ReDim Buffer((cbBuffer \ 4) - 1) As Long

Success = EnumPrinters(p, _
                      vbNullString, _
                      1, _
                      Buffer(0), _
                      cbBuffer, _
                      cbRequired, _
                      nEntries)
If Success Then
   If cbRequired > cbBuffer Then
      cbBuffer = cbRequired
      ReDim Buffer(cbBuffer \ 4) As Long
      Success = EnumPrinters(p Or _
                          PRINTER_ENUM_LOCAL, _
                          vbNullString, _
                          1, _
                          Buffer(0), _
                          cbBuffer, _
                          cbRequired, _
                          nEntries)
      If Not Success Then
         Call DebugFile("Fehler bei der Druckerauflistung.")
         Exit Sub
      End If
   End If
   For i = 0 To nEntries - 1
     PFlags = Buffer(4 * i)
     PDesc = Space$(StrLen(Buffer(i * 4 + 1)))
     Temp = PtrToStr(PDesc, Buffer(i * 4 + 1))
     PName = Space$(StrLen(Buffer(i * 4 + 2)))
     Temp = PtrToStr(PName, Buffer(i * 4 + 2))
     PComment = Space$(StrLen(Buffer(i * 4 + 2)))
     Temp = PtrToStr(PComment, Buffer(i * 4 + 2))
     ReDim Preserve Druckernamen(UBound(Druckernamen) + 1)
     Druckernamen(UBound(Druckernamen)) = PName
     
     Call DebugFile("Drucker: " & PName)
  Next i
Else
   Call DebugFile("Fehler bei der Druckerauflistung.")
End If
Call DefErrPop
End Sub


Sub EnumeratePrinters4(p As Long)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("EnumeratePrinters4")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Dim Success As Boolean, cbRequired As Long, cbBuffer As Long
Dim Buffer() As Long, nEntries As Long
Dim i As Long, PName As String, SName As String
Dim Attrib As Long, Temp As Long
cbBuffer = 3072
ReDim Buffer((cbBuffer \ 4) - 1) As Long

Success = EnumPrinters(p, _
                      vbNullString, _
                      4, _
                      Buffer(0), _
                      cbBuffer, _
                      cbRequired, _
                      nEntries)

If Success Then
   If cbRequired > cbBuffer Then
      cbBuffer = cbRequired
      ReDim Buffer(cbBuffer \ 4) As Long
      Success = EnumPrinters(p, _
                          vbNullString, _
                          4, _
                          Buffer(0), _
                          cbBuffer, _
                          cbRequired, _
                          nEntries)

      If Not Success Then
         Call DebugFile("Fehler bei der Druckerauflistung.")
         Exit Sub
      End If
   End If
   For i = 0 To nEntries - 1
     PName = Space$(StrLen(Buffer(i * 3)))
     Temp = PtrToStr(PName, Buffer(i * 3))
     SName = Space$(StrLen(Buffer(i * 3 + 1)))
     Temp = PtrToStr(SName, Buffer(i * 3 + 1))
     Attrib = Buffer(i * 3 + 2)
     ReDim Preserve Druckernamen(UBound(Druckernamen) + 1)
     Druckernamen(UBound(Druckernamen)) = PName
  
      Call DebugFile("Drucker: " & PName)
   Next i
Else
   Call DebugFile("Fehler bei der Druckerauflistung.")
End If
Call DefErrPop
End Sub



Sub main()

ReDim Druckernamen(0)

Call EnumeratePrinters1(PRINTER_ENUM_CONNECTIONS)
Call EnumeratePrinters4(PRINTER_ENUM_LOCAL)

Call SetPrinter(dNameAusINI)

End Sub


Sub SetPrinter(dName As String)

'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("SetPrinter")
On Error GoTo DefErr
GoTo DefErrEnd
DefErr:
If (Err.Number = (999 + vbObjectError)) Then End
Select Case DefErrAnswer(Err.Source, Err.Number, Err.Description, DefErrModul)
Case vbRetry
  Resume
Case vbIgnore
  Resume Next
End Select
End
'On Error GoTo 0
'Err.Raise 999 + vbObjectError, "DSK.VBP", "Fehler in DLL"
DefErrEnd:
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim hDrucker As Printer
Dim vglName As String
Dim dvglname As String
Dim s As String, i As Integer

dvglname = UCase(Trim(dName))

For i = 1 To UBound(Druckernamen)
  vglName = PrinterNameOP(Druckernamen(i))
  vglName = UCase(Trim(vglName))
  If dvglname = vglName Then
    For Each hDrucker In Printers
      If Trim(UCase(hDrucker.DeviceName)) = Trim(UCase(Druckernamen(i))) Then
        Set Printer = hDrucker
'        s = "Drucker wurde eingestellt, gedruckt wird auf: " + Printer.DeviceName
'        Call DebugFile(s)
        Exit For
      End If
    Next hDrucker
    Exit For
  Else
  End If
Next i
Call DefErrPop
End Sub


