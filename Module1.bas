Attribute VB_Name = "Module1"
Option Explicit

Sub firstVBA()
MsgBox ("�ڪ��Ĥ@��VBA")
End Sub
Sub assRange()
Range("A1").Value = "ExcelVBA"
Range("B1").Value = "�n�[�S���{���F��"
Cells(3, 3).Value = "���ӬO�o�˧a"
End Sub
Sub Timedisplay()
Cells(1, 1).Value = "��e�ɶ�"
Range("B1").Value = Now()
End Sub

Sub Timeclear()
Range("B1").Value = ""
End Sub
