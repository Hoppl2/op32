VERSION 5.00
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   7440
   Begin MSChartLib.MSChart MSChart1 
      Height          =   5295
      Left            =   600
      OleObjectBlob   =   "chart.frx":0000
      TabIndex        =   0
      Top             =   360
      Width           =   6735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim Column As Integer
Dim row As Integer

With MSChart1
    ' Zeigt ein 3D-Diagramm mit 8 Datenspalten
    ' und 8 Datenzeilen an.
    .chartType = VtChChartType3dBar
    .ColumnCount = 3
    .RowCount = 12
    For Column = 1 To 3
        For row = 1 To 12
            .Column = Column
            .row = row
            .Data = Column
        Next row
    Next Column
    ' Das Diagramm als Hintergrund für die
    ' Legende verwenden.
'    .ShowLegend = True
'    .SelectPart VtChPartTypePlot, index1, index2, index3, index4
'    .EditCopy
'    .SelectPart VtChPartTypeLegend, index1, _
'    index2, index3, index4
'    .EditPaste
End With

End Sub
