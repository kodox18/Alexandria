Attribute VB_Name = "Test"
Sub test_get_file_from_kalculator()
Dim r1, r2, r3, r4, range_r, book, opbook, nowbook, rng, filename, row_count, column_count, promoname, getOpenfile, booktime, listname, cancel_button As String
Dim k1, k2, k3, k4, z1, z2, errbox, file_path, uld, list_op As String
Dim lcounter, y, r_count, x As Long
Dim sel_rng As Range
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
'listname = Left(opbook, 24)
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
    Set sel_rng = Range(coord)
    k1 = Range(coord).Rows.Count
    k2 = Range(coord).Columns.Count
    Workbooks(nowbook).Worksheets(listname).Range(Cells(y, 1), Cells(k1 + y - 1, k2)).Value = Workbooks(opbook).Worksheets(list_op).Range(coord).Value
    End With
    Workbooks(opbook).Close
    For Each cell In Range(Cells(y, 1), Cells(k1 + y1 - 1, k2))
    z1 = Application.Cells.Find("TOTAL", LookIn:=xlValues).Offset(12, 6).Delete(xlShiftToLeft)
    Next cell
  ' If Range(Cells(y, 1), Cells(k1 + y - 1, k2)).Application.Cells.Find("TOTAL", LookIn:=xlValues) = "TOTAL" Then
  '  Range(Cells(y, 1), Cells(k1 + y - 1, k2)).Application.Cells.Find("TOTAL", Lookln:=xlValues).Offset(12, 6).Delete (xlShiftToLeft)
  '  End If
End Sub

