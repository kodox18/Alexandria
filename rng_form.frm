VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} rng_form 
   Caption         =   "Диапазон"
   ClientHeight    =   6915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9975.001
   OleObjectBlob   =   "rng_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "rng_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WF As Boolean

Private Sub CommandButton5_Click()
Dim nowbook As String
nowbook = ActiveWorkbook.Name
Text_r1.Value = Workbooks(nowbook).Worksheets("log").Cells(1, 1).Value
Text_r2.Value = Workbooks(nowbook).Worksheets("log").Cells(1, 2).Value
Text_r3.Value = Workbooks(nowbook).Worksheets("log").Cells(1, 3).Value
Text_r4.Value = Workbooks(nowbook).Worksheets("log").Cells(1, 4).Value
End Sub

Private Sub UserForm_Initialize()
Dim x, file_path, getOpenfile, list_op, uld  As String
Dim cancel_button As Boolean
Dim r1, r2, r3, r4 As String
End Sub
Private Sub CommandButton2_Click()
    WF = True
    Me.Hide
End Sub

Property Get WhiteFlag() As Boolean
    WhiteFlag = WF
End Property
Private Sub CommandButton1_Click()

Me.Hide
End Sub

Private Sub CommandButton4_Click()
Label6.Caption = "Выбранные данные " & "Лист: " & Text_list.Text & "Диапазон: " & Text_r1.Text & ":" & Text_r2.Text & Text_r3.Text & ":" & Text_r4.Text
Dim nowbook As String
nowbook = ActiveWorkbook.Name
With Workbooks(nowbook).Worksheets("log")
    .Cells(1, 1).Value = Text_r1.Value
    .Cells(1, 2).Value = Text_r2.Value
    .Cells(1, 3).Value = Text_r3.Value
    .Cells(1, 4).Value = Text_r4.Value
End With
End Sub

