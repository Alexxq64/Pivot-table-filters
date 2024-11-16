Attribute VB_Name = "Module1"
Public pt As PivotTable
Public pfMonth As PivotField
Public pfMnfcr As PivotField
Public mngrFilter As Range
Public pfMonthsQtty As Integer
Public pfMnfcrsQtty


Sub InitializePivotTable()
    ' Укажите здесь имя вашей сводной таблицы
    Set pt = ActiveSheet.PivotTables("Сводная таблица1")
    
    ' Укажите здесь имя нужного поля в сводной таблице
    Set pfMonth = pt.PivotFields("Месяц Года")
    Set pfMnfcr = pt.PivotFields("Производитель")
    ' NOTE: месяцев в фильтре может быть меньше 12!!!!
    pfMonthsQtty = pfMonth.PivotItems.Count
    pfMnfcrsQtty = pfMnfcr.PivotItems.Count

    ' Инициализируем диапазон mngrFilter для фильтрации
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Set mngrFilter = Range(Cells(6, 31), Cells(lastRow, 31))
End Sub



Sub KABFilter()
    ' Определение последней строки с данными в столбце A
    mngrFilter.AutoFilter Field:=31, Criteria1:="1"
End Sub

Sub KABUnFilter()
    ' Определение последней строки с данными в столбце A
    mngrFilter.AutoFilter Field:=31
End Sub
