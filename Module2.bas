Attribute VB_Name = "Module2"
Sub FilterByMonths(ParamArray monthNumbers() As Variant)

    Dim i As Integer
    Dim paramArrayUBound As Integer
    Dim newValue() As Boolean
    Dim shift As Integer
    Dim si As Integer

    Application.ScreenUpdating = False

    ' Устанавливаем верхнюю границу массива параметров
    paramArrayUBound = UBound(monthNumbers)
    
    ' так как сводная таблица обновляется при каждом изменении pivotItem.Visible меняем не все элементы,
    ' а только те, которые необходимо изменить
    ' иначе скорость выполнения фильтрации катастрофически падает
    ' Создаем массив newValue для хранения видимости элементов
    ReDim newValue(1 To pfMonthsQtty)

    ' Устанавливаем True для месяцев, указанных в monthNumbers
    For i = 0 To paramArrayUBound
        If monthNumbers(i) >= 1 And monthNumbers(i) <= pfMonthsQtty Then
            newValue(monthNumbers(i)) = True
        End If
    Next i

    ' Сдвиг, начиная с первого значения в monthNumbers
    shift = monthNumbers(0) - 1

    ' Проходим по элементам pivotItems, начиная с первого, у которого будет установлена visible = true
    ' чтобы избежать ситуации, когда в pivotField не будет ни одного item с visible = true
    For i = 0 To pfMonthsQtty - 1
        si = (i + shift) Mod pfMonthsQtty + 1  ' Рассчитываем текущий индекс с учетом сдвига
        ' Проверяем, нужно ли изменить видимость текущего элемента
        If pfMonth.PivotItems(si).Visible <> newValue(si) Then
            pfMonth.PivotItems(si).Visible = newValue(si)
        End If
    Next i

    ' Вызываем дополнительный фильтр, если необходимо
    KABFilter

    Application.ScreenUpdating = True

End Sub


Sub SecondQuarter()
    FilterByMonths 4, 5, 6
End Sub

Sub ThirdQuarter()
    FilterByMonths 7, 8, 9
End Sub

Sub SecondAndThirdQuarter()
    FilterByMonths 4, 5, 6, 7, 8, 9
End Sub

Sub April()
    FilterByMonths 4
End Sub

Sub May()
    FilterByMonths 5
End Sub

Sub June()
    FilterByMonths 6
End Sub

Sub July()
    FilterByMonths 7
End Sub

Sub August()
    FilterByMonths 8
End Sub

Sub September()
    FilterByMonths 9
End Sub

'
'Sub FilterByManufacturers(ParamArray manufacturerNames() As Variant)
'    Dim item As PivotItem
'    Dim i As Integer
'    Dim found As Boolean
'    Dim paramArrayUBound As Integer
'    Dim newValue() As Boolean
'
'    Application.ScreenUpdating = False
'    ' Сбрасываем фильтры перед установкой новых
'    pfMnfcr.ClearAllFilters
'
'    paramArrayUBound = UBound(manufacturerNames)
'    ' Проходим по каждому производителю в списке manufacturerNames
'    ReDim newValue(1 To pfMnfcrsQtty)
'
'
'
'    For Each item In pfMnfcr.PivotItems
'        found = False
'        ' Проверяем, присутствует ли текущий элемент в массиве manufacturerNames
'        For i = LBound(manufacturerNames) To UBound(manufacturerNames)
'            ' Сравниваем текущий элемент с переданными производителями
'            If item.Name = manufacturerNames(i) Then
'                found = True
'                Exit For
'            End If
'        Next i
'        ' Устанавливаем видимость текущего элемента
'        item.Visible = found
'    Next item
'    KABFilter
'    Application.ScreenUpdating = True
'End Sub

Sub FilterByManufacturers(ParamArray manufacturerNames() As Variant)
    Dim item As PivotItem
    Dim i As Integer
    Dim paramArrayUBound As Integer
    Dim newValue() As Boolean
    Dim itemPosition As Integer
'    Dim itemCount As Integer
    Dim j As Integer

    Application.ScreenUpdating = False

    ' Устанавливаем верхнюю границу массива параметров
    paramArrayUBound = UBound(manufacturerNames)
    
    ' Определяем количество элементов (производителей) в поле сводной таблицы
'    itemCount = pfMnfcr.PivotItems.Count
    ReDim newValue(1 To pfMnfcrsQtty)

    ' Устанавливаем True для элементов, указанных в manufacturerNames
    For Each item In pfMnfcr.PivotItems
        For i = 0 To paramArrayUBound
            If item.Name = manufacturerNames(i) Then
                ' Найти позицию элемента вручную
                For j = 1 To pfMnfcrsQtty
                    If pfMnfcr.PivotItems(j).Name = item.Name Then
                        itemPosition = j
                        Exit For
                    End If
                Next j
                
                ' Устанавливаем значение в массиве newValue
                newValue(itemPosition) = True

                ' Устанавливаем элемент видимым, если он еще не видим
                If item.Visible = False Then
                    item.Visible = True
                End If
                Exit For
            End If
        Next i
    Next item

    ' Второй проход: Устанавливаем невидимость для остальных элементов
    For i = 1 To pfMnfcrsQtty
        If newValue(i) = False Then
            ' Устанавливаем элемент невидимым, если он видим
            If pfMnfcr.PivotItems(i).Visible = True Then
                pfMnfcr.PivotItems(i).Visible = False
            End If
        End If
    Next i

    ' Вызываем дополнительный фильтр, если необходимо
    KABFilter

    Application.ScreenUpdating = True
End Sub





Sub FilterBagi()
    FilterByManufacturers "ТМ Bagi"
End Sub

Sub FilterRD()
    FilterByManufacturers "Российская дистрибьюция"
End Sub

Sub FilterHiyat()
    Call FilterByManufacturers("ООО ""Хаят Маркетинг""")
End Sub


Sub FilterDBH()
    FilterByManufacturers "ДомБытХим ООО"
End Sub

Sub FilterImpulse()
    FilterByManufacturers "Импульс ООО"
End Sub

Sub FilterAllManufacturers()
    FilterByManufacturers "Импульс ООО", "ДомБытХим ООО", "ООО ""Хаят Маркетинг""", "Российская дистрибьюция", "ТМ Bagi"
End Sub

