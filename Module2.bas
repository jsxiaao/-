Attribute VB_Name = "Module2"
Option Explicit
Public Function BMI(Height, Weight) As Integer
    BMI = Weight / ((Height / 100) ^ 2)
    
End Function
