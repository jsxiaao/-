VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormW 
   Caption         =   "身家調查"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6090
   OleObjectBlob   =   "UserFormW.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "UserFormW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btn_Click()
Dim School As String
School = tbSchool.Text
Cells(2, 1).Value = School

Dim Name As String
Name = tbName.Text
Cells(2, 2).Value = Name

Dim Number As String
Number = tbNumber.Text
Cells(2, 3).Value = Number

Dim Sexual As String
Sexual = tbSex.Text
Cells(2, 4).Value = Sexual

End Sub
