Attribute VB_Name = "Module1"
Option Explicit

Sub week301()
'Cells粂猭常ㄏノ计
Cells(1, 5).Value = Cells(1, 1).Value + Cells(1, 3).Value 'Cells计
Cells(2, 5).Value = Cells(1, 1).Value - Cells(1, 3).Value 'Cells计搭
Cells(3, 5).Value = Cells(1, 1).Value * Cells(1, 3).Value 'Cells计
Cells(4, 5).Value = Cells(1, 1).Value / Cells(1, 3).Value 'Cells计埃

'Cells粂猭ㄏノ璣计
Cells(1, "E").Value = Cells(1, "A").Value + Cells(1, "C").Value 'Cells璣计
Cells(2, "E").Value = Cells(1, "A").Value - Cells(1, "C").Value 'Cells璣计搭
Cells(3, "E").Value = Cells(1, "A").Value * Cells(1, "C").Value 'Cells璣计
Cells(4, "E").Value = Cells(1, "A").Value / Cells(1, "C").Value 'Cells璣计埃

'Range粂猭
Range("E1").Value = Range("A1").Value + Range("C1").Value 'Range璣计
Range("E2").Value = Range("A1").Value - Range("C1").Value 'Range璣计搭
Range("E3").Value = Range("A1").Value * Range("C1").Value 'Range璣计
Range("E4").Value = Range("A1").Value / Range("C1").Value 'Range璣计埃
End Sub

