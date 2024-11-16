Attribute VB_Name = "Module1"
Public pt As PivotTable
Public pfMonth As PivotField
Public pfMnfcr As PivotField
Public mngrFilter As Range
Public pfMonthsQtty As Integer
Public pfMnfcrsQtty


Sub InitializePivotTable()
    ' ������� ����� ��� ����� ������� �������
    Set pt = ActiveSheet.PivotTables("������� �������1")
    
    ' ������� ����� ��� ������� ���� � ������� �������
    Set pfMonth = pt.PivotFields("����� ����")
    Set pfMnfcr = pt.PivotFields("�������������")
    ' NOTE: ������� � ������� ����� ���� ������ 12!!!!
    pfMonthsQtty = pfMonth.PivotItems.Count
    pfMnfcrsQtty = pfMnfcr.PivotItems.Count

    ' �������������� �������� mngrFilter ��� ����������
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Set mngrFilter = Range(Cells(6, 31), Cells(lastRow, 31))
End Sub



Sub KABFilter()
    ' ����������� ��������� ������ � ������� � ������� A
    mngrFilter.AutoFilter Field:=31, Criteria1:="1"
End Sub

Sub KABUnFilter()
    ' ����������� ��������� ������ � ������� � ������� A
    mngrFilter.AutoFilter Field:=31
End Sub
