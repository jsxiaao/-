Attribute VB_Name = "Module1"
Option Explicit

Sub firstVBA()
MsgBox ("我的第一支VBA")
End Sub
Sub assRange()
Range("A1").Value = "ExcelVBA"
Range("B1").Value = "好久沒打程式了喔"
Cells(3, 3).Value = "應該是這樣吧"
End Sub
Sub Timedisplay()
Cells(1, 1).Value = "當前時間"
Range("B1").Value = Now()
End Sub

Sub Timeclear()
Range("B1").Value = ""
End Sub
