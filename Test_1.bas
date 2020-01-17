Attribute VB_Name = "Test"
Sub test_get_file_from_kalculator()
Dim r1, r2, r3, r4, range_r, book, opbook, nowbook, rng, filename, row_count, column_count, promoname, getOpenfile, booktime, listname, cancel_button As String
Dim k1, k2, k3, k4, errbox, file_path, uld, list_op As String
Dim lcounter, y, r_count, x As Long
Dim sel_rng As range
Dim ws As Worksheet
getOpenfile = GetFilePath("Выберите путь к нужному калькулятору", , "Excel files(*.xls*)", "*.xls*")
If getOpenfile = "" Then Exit Sub
MsgBox "Выбран файл " & getOpenfile, vbInformation
book = getOpenfile
nowbook = ThisWorkbook.Name
booktime = Format(Date, "DDMMYYYY") & Format(Time, "HHMMSS")
opbook = Dir(book)
rng = opbook
rng = CreateObject("Scripting.FileSystemObject").GetBaseName(opbook)
listname = Left(opbook, 24)
GetObject (book)


Worksheets.Add.Name = listname
    For y = 1 To y + 1
    If Cells(y, 1).Value = Empty Then
    Cells(y, 1).Value = Format(Date, "DD:MM:YYYY") & Format(Time, "HH:MM:SS") And Cells(y, 2).Value = Label6.Caption.Value
    End If
    Next y
    
    
    
FormR_cout.Show
cancel_button = FormR_cout.TextR_count.Value
If cancel_button = 1 Then
Exit Sub
End If

r_count = rng_form.Text_r_count.Value

For lcounter = 1 To r_count
    rng_form.Show
    r1 = rng_form.Text_r1.Value
    r2 = rng_form.Text_r2.Value
    list_op = rng_form.Text_list
    Unload rng_form

    Set ws = Workbooks(opbook).Worksheets(list_op)
    With ws
    coord = r1 & ":" & r2
    Set sel_rng = range(coord)
    k1 = range(coord).Rows.Count
    k2 = range(coord).Columns.Count
    Workbooks(nowbook).Worksheets(listname).range(Cells(y, 1), Cells(k1, k2)).Value = Workbooks(opbook).Worksheets(list_op).range(coord).Value
    End With
        If range(Cells(y, 1), Cells(k1, k2)).Application.Cells.Find("TOTAL") = "TOTAL" Then
        Set r3 = range(Cells(y, 1), Cells(k1, k2)).Application.Cells.Find("TOTAL").Offset(12, 6).Delete(xlShiftToLeft)
        
        End If
Next lcounter
End Sub

