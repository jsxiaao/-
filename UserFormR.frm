VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormR 
   Caption         =   "身家調查"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5415
   OleObjectBlob   =   "UserFormR.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "UserFormR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()

Dim School As String
School = Cells(2, 1).Value
schoolResult.Caption = School


Dim Name As String
Name = Cells(2, 2).Value
NameResult.Caption = Name

Dim Number As String
Number = Cells(2, 3).Value
NumberResult.Caption = Number

Dim Sexual As String
Sexual = Cells(2, 4).Value
SexResult.Caption = Sexual

End Sub

