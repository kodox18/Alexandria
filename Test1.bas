Attribute VB_Name = "Test"
Sub test_get_file_from_kalculator()
Dim r1, r2, r3, r4, book, opbook, nowbook, rng, filename, promoname, getOpenfile, booktime, listname, cancel_button As String
Dim k1, k2, k3, k4, a1, a2, file_path, uld, list_op As String
Dim lcounter, y, r_count, x, i As Long
Dim sel_rng, z1, z2 As Range
Dim ws As Worksheet
getOpenfile = GetFilePath("Выберите путь к нужному калькулятору", , "Excel files(*.xls*)", "*.xls*")
If getOpenfile = "" Then Exit Sub
'MsgBox "Выбран файл " & getOpenfile, vbInformation
book = getOpenfile
nowbook = ThisWorkbook.Name
opbook = Dir(book)
rng = opbook
rng = CreateObject("Scripting.FileSystemObject").GetBaseName(opbook)
listname = Workbooks(nowbook).ActiveSheet.Name
GetObject (book)

rng_form.Show
    r1 = rng_form.Text_r1.Value
    r2 = rng_form.Text_r2.Value
    r3 = rng_form.Text_r3.Value
    r4 = rng_form.Text_r4.Value
    list_op = rng_form.Text_list
    
With ThisWorkbook.ActiveSheet
    y = 1
    Do Until Cells(y, 1).Value = 0
    y = y + 1
    Loop
    If Cells(y, 1).Value = 0 Then
    Cells(y, 1).Value = Format(Date, "DD:MM:YYYY") & " " & Format(Time, "HH:MM:SS")
    y = y + 1
    End If
End With
    
Unload rng_form

Set ws = Workbooks(opbook).Worksheets(list_op)
With ws
    coord = r1 & ":" & r2
    k1 = Range(coord).Rows.Count
    k2 = Range(coord).Columns.Count
    Workbooks(nowbook).Worksheets(listname).Range(Cells(y, 1), Cells(k1 + y - 1, k2)).Value = Workbooks(opbook).Worksheets(list_op).Range(coord).Value
End With

    With Range(Cells(y, 1), Cells(k1 + y - 1, k2))
        Range(Application.Cells.Find("TOTAL", LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("TOTAL", LookIn:=xlValues, lookat:=xlWhole).Offset(k1 - 1, 0)).Select
        Selection.Delete Shift:=xlShiftToLeft
        Range(Application.Cells.Find("SAVOURY", LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("SAVOURY", LookIn:=xlValues, lookat:=xlWhole).Offset(k1 - 1, 0)).Select
        Selection.Delete Shift:=xlShiftToLeft
        Range(Application.Cells.Find("DRESSINGS", LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("DRESSINGS", LookIn:=xlValues, lookat:=xlWhole).Offset(k1 - 1, 0)).Select
        Selection.Delete Shift:=xlShiftToLeft
        Range(Application.Cells.Find("SPREADS", LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("SPREADS", LookIn:=xlValues, lookat:=xlWhole).Offset(k1 - 1, 0)).Select
        Selection.Delete Shift:=xlShiftToLeft
        Range(Application.Cells.Find("IC", LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("IC", LookIn:=xlValues, lookat:=xlWhole).Offset(k1 - 1, 0)).Select
        Selection.Delete Shift:=xlShiftToLeft
        Range(Application.Cells.Find("HHC", LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("HHC", LookIn:=xlValues, lookat:=xlWhole).Offset(k1 - 1, 0)).Select
        Selection.Delete Shift:=xlShiftToLeft
        Range(Application.Cells.Find("TEA", LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("TEA", LookIn:=xlValues, lookat:=xlWhole).Offset(k1 - 1, 0)).Select
        Selection.Delete Shift:=xlShiftToLeft
    End With
    Set ws = Workbooks(nowbook).Worksheets(listname)
    With ws
        k1 = Range(Cells(y, 1), Cells(k1 + y - 1, k2 - 7)).Rows.Count
        k2 = Range(Cells(y, 1), Cells(k1 + y - 1, k2 - 7)).Columns.Count
    End With
    
        Workbooks(nowbook).Worksheets(listname).Range(Cells(y, 1), Cells(k1 + y - 1, k2)).Select
        Selection.Copy
        Workbooks(nowbook).Worksheets(listname).Range(Cells(y, k2 + 1), Cells(y + k2, k1 + k2)).Select
        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
        Workbooks(nowbook).Worksheets(listname).Range(Cells(y, 1), Cells(k1 + y - 1, k2)).Select
        Selection.Delete Shift:=xlToLeft
        
    Set ws = Workbooks(opbook).Worksheets(list_op)
With ws
    coord = r3 & ":" & r4
    k3 = Range(coord).Rows.Count
    k4 = Range(coord).Columns.Count
    Workbooks(nowbook).Worksheets(listname).Range(Cells(y, k1 + 1), Cells(y + k2 - 1, k1 + k4)).Value = Workbooks(opbook).Worksheets(list_op).Range(coord).Value
End With
    Cells(y + k2, 1).Select
    Selection = "Promo name"
    Workbooks(opbook).Close
End Sub
Function GetFilePath(Optional ByVal Title As String = "Выберите файл", _
                     Optional ByVal InitialPath As String = "c:\", _
                     Optional ByVal FilterDescription As String = "Excel", _
                     Optional ByVal FilterExtention As String = "*.xls*") As String
    On Error Resume Next
    With Application.FileDialog(msoFileDialogOpen)
        .ButtonName = "Выбрать": .Title = Title: .InitialFileName = InitialPath
        .Filters.Clear: .Filters.Add FilterDescription, FilterExtention
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        GetFilePath = .SelectedItems(1): PS = Application.PathSeparator
    End With
End Function
