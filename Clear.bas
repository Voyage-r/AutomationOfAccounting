Attribute VB_Name = "Module1"
Sub Delete() ' - �������� � ������������ ���� �����, ����� ������ ������
    Rows("2:" & Rows.Count).ClearContents
    Rows("2:" & Rows.Count).Style = "Normal"
End Sub
