Attribute VB_Name = "Module2"
Sub FilterByMonths(ParamArray monthNumbers() As Variant)

    Dim i As Integer
    Dim paramArrayUBound As Integer
    Dim newValue() As Boolean
    Dim shift As Integer
    Dim si As Integer

    Application.ScreenUpdating = False

    ' ������������� ������� ������� ������� ����������
    paramArrayUBound = UBound(monthNumbers)
    
    ' ��� ��� ������� ������� ����������� ��� ������ ��������� pivotItem.Visible ������ �� ��� ��������,
    ' � ������ ��, ������� ���������� ��������
    ' ����� �������� ���������� ���������� ��������������� ������
    ' ������� ������ newValue ��� �������� ��������� ���������
    ReDim newValue(1 To pfMonthsQtty)

    ' ������������� True ��� �������, ��������� � monthNumbers
    For i = 0 To paramArrayUBound
        If monthNumbers(i) >= 1 And monthNumbers(i) <= pfMonthsQtty Then
            newValue(monthNumbers(i)) = True
        End If
    Next i

    ' �����, ������� � ������� �������� � monthNumbers
    shift = monthNumbers(0) - 1

    ' �������� �� ��������� pivotItems, ������� � �������, � �������� ����� ����������� visible = true
    ' ����� �������� ��������, ����� � pivotField �� ����� �� ������ item � visible = true
    For i = 0 To pfMonthsQtty - 1
        si = (i + shift) Mod pfMonthsQtty + 1  ' ������������ ������� ������ � ������ ������
        ' ���������, ����� �� �������� ��������� �������� ��������
        If pfMonth.PivotItems(si).Visible <> newValue(si) Then
            pfMonth.PivotItems(si).Visible = newValue(si)
        End If
    Next i

    ' �������� �������������� ������, ���� ����������
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
'    ' ���������� ������� ����� ���������� �����
'    pfMnfcr.ClearAllFilters
'
'    paramArrayUBound = UBound(manufacturerNames)
'    ' �������� �� ������� ������������� � ������ manufacturerNames
'    ReDim newValue(1 To pfMnfcrsQtty)
'
'
'
'    For Each item In pfMnfcr.PivotItems
'        found = False
'        ' ���������, ������������ �� ������� ������� � ������� manufacturerNames
'        For i = LBound(manufacturerNames) To UBound(manufacturerNames)
'            ' ���������� ������� ������� � ����������� ���������������
'            If item.Name = manufacturerNames(i) Then
'                found = True
'                Exit For
'            End If
'        Next i
'        ' ������������� ��������� �������� ��������
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

    ' ������������� ������� ������� ������� ����������
    paramArrayUBound = UBound(manufacturerNames)
    
    ' ���������� ���������� ��������� (��������������) � ���� ������� �������
'    itemCount = pfMnfcr.PivotItems.Count
    ReDim newValue(1 To pfMnfcrsQtty)

    ' ������������� True ��� ���������, ��������� � manufacturerNames
    For Each item In pfMnfcr.PivotItems
        For i = 0 To paramArrayUBound
            If item.Name = manufacturerNames(i) Then
                ' ����� ������� �������� �������
                For j = 1 To pfMnfcrsQtty
                    If pfMnfcr.PivotItems(j).Name = item.Name Then
                        itemPosition = j
                        Exit For
                    End If
                Next j
                
                ' ������������� �������� � ������� newValue
                newValue(itemPosition) = True

                ' ������������� ������� �������, ���� �� ��� �� �����
                If item.Visible = False Then
                    item.Visible = True
                End If
                Exit For
            End If
        Next i
    Next item

    ' ������ ������: ������������� ����������� ��� ��������� ���������
    For i = 1 To pfMnfcrsQtty
        If newValue(i) = False Then
            ' ������������� ������� ���������, ���� �� �����
            If pfMnfcr.PivotItems(i).Visible = True Then
                pfMnfcr.PivotItems(i).Visible = False
            End If
        End If
    Next i

    ' �������� �������������� ������, ���� ����������
    KABFilter

    Application.ScreenUpdating = True
End Sub





Sub FilterBagi()
    FilterByManufacturers "�� Bagi"
End Sub

Sub FilterRD()
    FilterByManufacturers "���������� ������������"
End Sub

Sub FilterHiyat()
    Call FilterByManufacturers("��� ""���� ���������""")
End Sub


Sub FilterDBH()
    FilterByManufacturers "��������� ���"
End Sub

Sub FilterImpulse()
    FilterByManufacturers "������� ���"
End Sub

Sub FilterAllManufacturers()
    FilterByManufacturers "������� ���", "��������� ���", "��� ""���� ���������""", "���������� ������������", "�� Bagi"
End Sub

