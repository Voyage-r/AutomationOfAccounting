VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "����4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub Prepare() '- ���������� ���� �� ������ �������
    Application.ScreenUpdating = False ' - ���������� ������
    Application.Calculation = xlCalculationManual ' - �����������
    Application.EnableEvents = False '- ���������� �������
    ActiveSheet.DisplayPageBreaks = False ' - ���������� �������
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
End Sub
Sub FilenamesCollection_OC()

    Prepare ' - ���������� ����� ���������
    
    Dim coll As Collection, FilePath As String
    Dim Column As Integer, Name As String
    Dim Proverka As Integer
    
    FilePath = ThisWorkbook.Path & "\��" ' - �������� ���� � ������
    ' ��������� � �������� coll ������ ���� ������ � ����� .xlsx
    Set coll = FilenamesCollection(FilePath, ".xlsx") '- ������� ��� ���������� ���� � ������������ �����
    
    ' ������� ���������� �� ����
    For i = 1 To coll.Count ' ���������� ��� �������� ���������, ���������� ���� � ������
        Proverka = EnumerationDateTime(coll.Item(i)) '- �������� ����� �� �������������
        
        If (Proverka = 0) Then
            Column = WorksheetFunction.CountA(Columns(1)) + 1
            ImportData coll.Item(i), Column ' - ����������� ������
            '���� ���������� �������
            For j = 1 To 6
                '���� �������� ����������
                Cells(Column, j).Borders.LineStyle = True
                Cells(Column, j).Font.Name = "Times New Roman"
                If j = 1 Then
                    Name = CStr(Mid(coll.Item(i), InStrRev(coll.Item(i), "\") + 1)) ' ����������� ��� �����, ��� ����������
                    Cells(Column, j) = Left(Name, InStrRev(Name, ".") - 1)
                End If
                If j = 2 Then '- ������ �� ����
                    ActiveSheet.Hyperlinks.Add Anchor:=Cells(Column, j), Address:=coll.Item(i), TextToDisplay:="OK"
                End If
                If j <> 5 Then ' - ����� ����������
                    Cells(Column, j).HorizontalAlignment = xlCenter
                    Cells(Column, j).VerticalAlignment = xlCenter
                End If
                If j = 5 Then ' - ����� ����������
                    Cells(Column, j).HorizontalAlignment = xlLeft
                    Cells(Column, j).VerticalAlignment = xlTop
                    Cells(Column, j).Font.Size = 9
                End If
                Cells(Column, j).EntireColumn.AutoFit ' - �������������� ������ ������� �������
            Next
            Cells(Column, 8) = FileDateTime(coll.Item(i)) ' - ��������� ����� ����� ��� ������� �������� ��������
            Cells(Column, 9) = coll.Item(i) ' - ��������� ���� ����� ��� ������� �������� ��������
        ElseIf (Proverka > 0) Then
            Column = Proverka
            '�������������� �������
            ImportData coll.Item(i), Column ' - ����������� ������
            '�������������� ���� ��������� �����
            Cells(Column, 8) = FileDateTime(coll.Item(i))
        End If
    Next
    
    Ended ' - ��������� ����� ���������
    Dim a As Integer
    a = MsgBox("��� ���������", 0, "������������ ��������") ' - ���� � ����������� ��� ��� ������ �������
    
End Sub

Public Function EnumerationDateTime(ByVal CollectionName As String) As Integer ' - �������� ���� �� �����, ������ � ���������� ������ ����
    Dim Path As String
    Dim DateTime As Date
    
    Path = CollectionName ' - ���� �����
    DateTime = FileDateTime(CollectionName) ' - ���� ���������� �������� �����
    
    For i = 2 To WorksheetFunction.CountA(Columns(1))
        If (Path = Cells(i, 9)) And (DateTime = Cells(i, 8)) Then ' - ������������ ���� � ���� ��������� �����
            EnumerationDateTime = -1 ' - ������ ���� ��� ���������    ���� ��� ��������� �� ������ �������� -1
            Exit For
        ElseIf (Path = Cells(i, 9)) And (DateTime <> Cells(i, 8)) Then
            EnumerationDateTime = i ' - ������ ����, �� � �����������
            Exit For
        Else
            EnumerationDateTime = 0 ' - - ���� ����� ����
        End If
    Next
    
End Function

Public Function FilenamesCollection(ByVal FolderPath As String, Optional ByVal Mask As String = "") As Collection ' - ���������� ������ ������
    'Optional ByVal SearchDeep As Long = 999 - ��� ��������� �� �������
    '���������� �������� ��� �������� ���� ���� ������ � ���������
    Set FilenamesCollection = New Collection    ' ������ ������ ���������
    Set FSO = CreateObject("Scripting.FileSystemObject")    ' ������ ��������� FileSystemObject
    GetAllFileNamesUsingFSO FolderPath, Mask, FSO, FilenamesCollection ' �����
    Set FSO = Nothing: Application.StatusBar = False    ' ������� ������ ��������� Excel
End Function

Public Function GetAllFileNamesUsingFSO(ByVal FolderPath As String, ByVal Mask As String, ByRef FSO, _
                                 ByRef FileNamesColl As Collection)
    'ByVal SearchDeep As Long - �������� ��������
    On Error Resume Next: Set curfold = FSO.GetFolder(FolderPath)
    If Not curfold Is Nothing Then    ' ���� ������� �������� ������ � �����

        For Each fil In curfold.Files    ' ���������� ��� ����� � ����� FolderPath
            If fil.Name Like "[!~]*" & Mask Then FileNamesColl.Add fil.Path
        Next
        'SearchDeep = SearchDeep - 1  - ��������� ������� ������ � ���������
        'If SearchDeep Then    ' ���� ���� ������ ������
        '    For Each sfol In curfold.SubFolders    ' ' ���������� ��� �������� � ����� FolderPath
        '        GetAllFileNamesUsingFSO sfol.Path, Mask, FSO, FileNamesColl, SearchDeep
        '    Next
        'End If
        
        Set fil = Nothing: Set curfold = Nothing    ' ������� ����������
    End If
End Function
Public Function ImportData(ByVal FilePath As String, ByVal Column As Integer) ' - ���������� ������ �� ������������ ����� � ����������� � ������

    Dim sShName As String, sAddress As String
    Dim objCloseBook As Object
    Dim ArrayRange As Variant
    Dim Row As Integer
    
    Application.ScreenUpdating = False
    Set objCloseBook = GetObject(FilePath) ' - �������� �����
    ArrayRange = Array(objCloseBook.Sheets(1).Range("A39").Value, objCloseBook.Sheets(1).Range("E11").Value, objCloseBook.Sheets(1).Range("A19").Value, objCloseBook.Sheets(1).Range("A17").Value)
    objCloseBook.Close False
    
    Row = 3
    For Each element In ArrayRange
        Cells(Column, Row) = element
        Row = Row + 1
    Next
    
End Function


Public Sub Ended() ' - ��������� ����� ���������
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
End Sub
