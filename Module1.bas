Attribute VB_Name = "Module1"
Option Explicit

Sub ����аO�����()
Dim i, rowCnt As Integer
Dim tagetValue As Integer
tagetValue = CInt(InputBox("�п�J�аO�W����(0-1000)"))
Dim rangeStr As String
rowCnt = Cells(Rows.Count, 1).End(xlUp).Row
rangeStr = "b3:b" & rowCnt
MsgBox "�ثe�B��d��" & rangeStr
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
