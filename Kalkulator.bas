Attribute VB_Name = "Kalkulator"
Sub Auto_PL_Pre_Pack()
Dim book, opbook, nowbook, rng, filename, promoname, getOpenfile, booktime, listname, cancel_button As String
Dim k1, k2, k3, k4, a1, a2, file_path, uld, list_op As String
Dim lcounter, y, r_count, x, i As Long
Dim sel_rng, z1, z2 As Range
Dim ws, wsSh As Worksheet
Static r1, r2, r3, r4 As String
getOpenfile = GetFilePath("Выберите путь к нужному калькулятору", , "Excel files(*.xls*)", "*.xls*")
If getOpenfile = "" Then Exit Sub
book = getOpenfile
nowbook = ThisWorkbook.Name
opbook = Dir(book)
rng = opbook
rng = CreateObject("Scripting.FileSystemObject").GetBaseName(opbook)
listname = Workbooks(nowbook).ActiveSheet.Name
    On Error Resume Next
    Set wsSh = Sheets("log")
    If wsSh Is Nothing Then Sheets.Add(, Sheets(Sheets.Count)).Name = "log"
    Sheets("log").Visible = 2

GetObject (book)

rng_form.Show

    If rng_form.WhiteFlag Then Exit Sub

    r1 = rng_form.Text_r1.Value
    r2 = rng_form.Text_r2.Value
    r3 = rng_form.Text_r3.Value
    r4 = rng_form.Text_r4.Value
    list_op = rng_form.Text_list
    
Workbooks(nowbook).Worksheets(listname).Activate
With ThisWorkbook.ActiveSheet
    y = 1
    Do Until Cells(y, 1).Value = 0
    y = y + 1
    Loop
    If Cells(y, 1).Value = 0 Then
    Cells(y, 1).Value = Format(Date, "DD:MM:YYYY") & " " & Format(Time, "HH:MM:SS") & " Лист: " & rng_form.Text_list.Text & " Диапазон: " & rng_form.Text_r1.Text & ":" & rng_form.Text_r2.Text & " " & rng_form.Text_r3.Text & ":" & rng_form.Text_r4.Text
    y = y + 1
    End If
End With
Unload rng_form

Set ws = Workbooks(opbook).Worksheets(list_op)
With ws
    coord = r1 & ":" & r2
    k1 = Range(coord).Rows.Count
    k2 = Range(coord).Columns.Count
    
    Range(Cells(y, 1), Cells(k1 + y - 1, k2)).Value = Workbooks(opbook).Worksheets(list_op).Range(coord).Value
End With
        Range(Application.Cells.Find("TOTAL", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("TOTAL", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(13, 0)).Select
        Selection.Delete Shift:=xlShiftToLeft
        Range(Application.Cells.Find("SAVOURY", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("SAVOURY", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(13, 0)).Select
        Selection.Delete Shift:=xlShiftToLeft
        Range(Application.Cells.Find("DRESSINGS", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("DRESSINGS", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(13, 0)).Select
        Selection.Delete Shift:=xlShiftToLeft
        Range(Application.Cells.Find("SPREADS", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("SPREADS", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(13, 0)).Select
        Selection.Delete Shift:=xlShiftToLeft
        Range(Application.Cells.Find("IC", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("IC", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(13, 0)).Select
        Selection.Delete Shift:=xlShiftToLeft
        Range(Application.Cells.Find("HHC", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("HHC", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(13, 0)).Select
        Selection.Delete Shift:=xlShiftToLeft
        Range(Application.Cells.Find("TEA", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("TEA", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(13, 0)).Select
        Selection.Delete Shift:=xlShiftToLeft
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

    Range(Cells(y + 6, 1), Cells(y + 11, 1)).Value = Range(Application.Cells.Find("STATUS", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("STATUS", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 2), Cells(y + 11, 2)).Value = Range(Application.Cells.Find("Client Group", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("Client Group", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 3), Cells(y + 11, 3)).Value = Range(Application.Cells.Find("Owner Name", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("Owner Name", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 4), Cells(y + 11, 4)).Value = Range(Application.Cells.Find("TYPE FOR SPLIT GEN", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("TYPE FOR SPLIT GEN", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 5), Cells(y + 11, 5)).Value = Range(Application.Cells.Find("TYPE for split detailed", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("TYPE for split detailed", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 6), Cells(y + 11, 6)).Value = Range(Application.Cells.Find("Client GROUP_ AUTO", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("Client GROUP_ AUTO", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 7), Cells(y + 11, 7)).Value = Range(Application.Cells.Find("Client", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("Client", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Cells(y + 6, 8).Value = "Promo name"
    Range(Cells(y + 7, 8), Cells(y + 11, 8)).Value = rng
    Range(Cells(y + 6, 9), Cells(y + 11, 9)).Value = Range(Application.Cells.Find("Period Promo", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("Period Promo", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 10), Cells(y + 11, 10)).Value = Range(Application.Cells.Find("Period budget fact", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("Period budget fact", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 11), Cells(y + 11, 11)).Value = Range(Application.Cells.Find("Promo ID", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("Promo ID", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 12), Cells(y + 11, 12)).Value = Range(Application.Cells.Find("Rub", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("Rub", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 13), Cells(y + 11, 13)).Value = Range(Application.Cells.Find("GSV", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("GSV", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 14), Cells(y + 11, 14)).Value = Range(Application.Cells.Find("Incr GSV", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("Incr GSV", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 15), Cells(y + 11, 15)).Value = Range(Application.Cells.Find("CPP on", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("CPP on", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 16), Cells(y + 11, 16)).Value = Range(Application.Cells.Find("CPP off", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("CPP off", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 17), Cells(y + 11, 17)).Value = Range(Application.Cells.Find("TCC", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("TCC", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 18), Cells(y + 11, 18)).Value = Range(Application.Cells.Find("A&V", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("A&V", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 19), Cells(y + 11, 19)).Value = Range(Application.Cells.Find("Incr Turnover", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("Incr Turnover", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 20), Cells(y + 11, 20)).Value = Range(Application.Cells.Find("SCC", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("SCC", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 21), Cells(y + 11, 21)).Value = Range(Application.Cells.Find("Inr GP", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("Inr GP", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 22), Cells(y + 11, 22)).Value = Range(Application.Cells.Find("ROI, % (w/o BMI)", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("ROI, % (w/o BMI)", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 23), Cells(y + 11, 23)).Value = Range(Application.Cells.Find("A&P", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("A&P", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 24), Cells(y + 11, 24)).Value = Range(Application.Cells.Find(" Incr PBI", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find(" Incr PBI", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 25), Cells(y + 11, 25)).Value = Range(Application.Cells.Find("ROI,%", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("ROI,%", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 26), Cells(y + 11, 26)).Value = Range(Application.Cells.Find("Comments", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("Comments", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 27), Cells(y + 11, 27)).Value = Range(Application.Cells.Find("Category Check (AUTO)", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("Category Check (AUTO)", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 28), Cells(y + 11, 28)).Value = Range(Application.Cells.Find("Total Investments", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("Total Investments", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Range(Cells(y + 6, 29), Cells(y + 11, 29)).Value = Range(Application.Cells.Find("Manager", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole), Application.Cells.Find("Manager", after:=(Cells(y, 1)), LookIn:=xlValues, lookat:=xlWhole).Offset(5, 0)).Value
    Workbooks(nowbook).Worksheets(listname).Range(Cells(y, 1), Cells(y + 5, k4 + k4 + 1)).Select
        Selection.Delete Shift:=xlUp
        
    Range(Cells(y, 1), Cells(y + 5, 7)).Interior.Color = RGB(0, 176, 240)
    Range(Cells(y, 8), Cells(y + 5, 8)).Interior.Color = RGB(146, 208, 80)
    Range(Cells(y, 9), Cells(y + 5, 11)).Interior.Color = RGB(0, 176, 240)
    Range(Cells(y, 12), Cells(y + 5, 25)).Interior.Color = RGB(146, 208, 80)
    Range(Cells(y, 26), Cells(y + 5, 29)).Interior.Color = RGB(0, 176, 240)
    Range(Cells(y, 1), Cells(y + 5, 29)).Borders.LineStyle = True
    Range(Cells(y, 1), Cells(y + 5, 29)).Columns.AutoFit
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
