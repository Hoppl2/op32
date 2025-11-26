Attribute VB_Name = "Module2"
Option Explicit

Private Const DefErrModul = "whotkeys.frm"

Sub ZeigeStatbild(SuchPzn$)
'DefErr!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Call DefErrFnc("ZeigeStatbild")
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
Dim TaskId&
Dim STATBILD%, erg%
Dim dPreis#
Dim s$, pzn$, ch$, SQLStr$
Dim recTaxe As Recordset

If (SuchPzn$ = "") Then Call DefErrPop: Exit Sub

FabsErrf% = ass.IndexSearch(0, SuchPzn$, FabsRecno&)
If (FabsErrf% = 0) Then
    ass.GetRecord (FabsRecno& + 1)
    
    FabsErrf% = ast.IndexSearch(0, SuchPzn$, FabsRecno&)
    If (FabsErrf% = 0) Then
        ast.GetRecord (FabsRecno& + 1)
    
        STATBILD% = FileOpen%("statb" + para.User + ".$$$", "W")
'        STATBILD% = FileOpen%("statbild.$$$", "W")
        Put STATBILD%, , ast
        Put STATBILD%, , ass
        
        s$ = String$(200, 0)
        Put STATBILD%, , s$
        
        SQLStr$ = "SELECT * FROM TAXE WHERE PZN = " + SuchPzn$
        Set recTaxe = TaxeDB.OpenRecordset(SQLStr$)
        If (recTaxe.EOF = False) Then
            dPreis# = recTaxe!EK / 100
            Call DxToMBFd(dPreis#)
            Put STATBILD%, , dPreis#
            dPreis# = recTaxe!VK / 100
            Call DxToMBFd(dPreis#)
            Put STATBILD%, , dPreis#
        End If
        
        Put STATBILD%, , Left$(s$, 2)
        
        Close #STATBILD%
        
        TaskId& = Shell("\user\dosrun.bat " + para.User + " statbild.exe", vbNormalFocus)
        Call WarteAufTaskEnde(TaskId&)
'        AppActivate Me
        
    End If
End If
        
Call DefErrPop
End Sub


