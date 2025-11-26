VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim dHandle%, dSatzLen%, iAnzIndex%, iHandle%(1)
Dim dName$, dZugriff$, dModus$, iName$(1)

Sub OpenDatei()
Dim i%, FabsErrf%

If (RTrim$(dName$) <> "") Then
    dHandle% = clsDat.FileOpen(RTrim$(dName$), RTrim$(dZugriff$), dModus$, dSatzLen%)
End If
For i% = 0 To (iAnzIndex% - 1)
     FabsErrf% = clsDat.IndexOpen(iName$(i%), iHandle%(i%))
Next i%

End Sub


