Attribute VB_Name = "Module1"
Option Explicit

Sub week301()
'Cells�y�k���ϥμƦr
Cells(1, 5).Value = Cells(1, 1).Value + Cells(1, 3).Value 'Cells�Ʀr�ۥ[
Cells(2, 5).Value = Cells(1, 1).Value - Cells(1, 3).Value 'Cells�Ʀr�۴�
Cells(3, 5).Value = Cells(1, 1).Value * Cells(1, 3).Value 'Cells�Ʀr�ۭ�
Cells(4, 5).Value = Cells(1, 1).Value / Cells(1, 3).Value 'Cells�Ʀr�۰�

'Cells�y�k�ϥέ^�Ʀr
Cells(1, "E").Value = Cells(1, "A").Value + Cells(1, "C").Value 'Cells�^�Ʀr�ۥ[
Cells(2, "E").Value = Cells(1, "A").Value - Cells(1, "C").Value 'Cells�^�Ʀr�۴�
Cells(3, "E").Value = Cells(1, "A").Value * Cells(1, "C").Value 'Cells�^�Ʀr�ۭ�
Cells(4, "E").Value = Cells(1, "A").Value / Cells(1, "C").Value 'Cells�^�Ʀr�۰�

'Range�y�k
Range("E1").Value = Range("A1").Value + Range("C1").Value 'Range�^�Ʀr�ۥ[
Range("E2").Value = Range("A1").Value - Range("C1").Value 'Range�^�Ʀr�۴�
Range("E3").Value = Range("A1").Value * Range("C1").Value 'Range�^�Ʀr�ۭ�
Range("E4").Value = Range("A1").Value / Range("C1").Value 'Range�^�Ʀr�۰�
End Sub

