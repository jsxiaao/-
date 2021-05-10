Attribute VB_Name = "Module1"
Option Explicit

Sub 產能標記實習檔()
Dim i, rowCnt As Integer
Dim tagetValue As Integer
tagetValue = CInt(InputBox("請輸入標記上限值(0-1000)"))
Dim rangeStr As String
rowCnt = Cells(Rows.Count, 1).End(xlUp).Row
rangeStr = "b3:b" & rowCnt
MsgBox "目前運算範圍" & rangeStr
Range(rangeStr).Interior.Color = xlNone
For i = 3 To rowCnt
   If Cells(i, "B") > tagetValue Then
   Cells(i, "B").Interior.Color = vbYellow
   End If
   
   If Cells(i, "B") < tagetValue Then
   Cells(i, "B").Interior.Color = vbBlue
   End If
   Next
   Range("a1").CurrentRegion.Borders.LineStyle = xlContinuous
End Sub
