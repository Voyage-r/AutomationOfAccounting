VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub Prepare() '- отключение всех не нужных функций
    Application.ScreenUpdating = False ' - обновление экрана
    Application.Calculation = xlCalculationManual ' - калькуляция
    Application.EnableEvents = False '- обновление ивентов
    ActiveSheet.DisplayPageBreaks = False ' - обновление дисплея
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
End Sub
Sub FilenamesCollection_OC()

    Prepare ' - отключение всего ненужного
    
    Dim coll As Collection, FilePath As String
    Dim Column As Integer, Name As String
    Dim Proverka As Integer
    
    FilePath = ThisWorkbook.Path & "\ОС" ' - добавить путь к файлам
    ' считываем в колекцию coll нужные пути файлов с типом .xlsx
    Set coll = FilenamesCollection(FilePath, ".xlsx") '- находит все подходящие пути в определенной папке
    
    ' выводим результаты на лист
    For i = 1 To coll.Count ' перебираем все элементы коллекции, содержащей пути к файлам
        Proverka = EnumerationDateTime(coll.Item(i)) '- Проверка файла на повторяемость
        
        If (Proverka = 0) Then
            Column = WorksheetFunction.CountA(Columns(1)) + 1
            ImportData coll.Item(i), Column ' - считываются данные
            'Идет заполнение таблицы
            For j = 1 To 6
                'идет красивое оформление
                Cells(Column, j).Borders.LineStyle = True
                Cells(Column, j).Font.Name = "Times New Roman"
                If j = 1 Then
                    Name = CStr(Mid(coll.Item(i), InStrRev(coll.Item(i), "\") + 1)) ' заполняется имя файла, оно уникальное
                    Cells(Column, j) = Left(Name, InStrRev(Name, ".") - 1)
                End If
                If j = 2 Then '- ссылка на файл
                    ActiveSheet.Hyperlinks.Add Anchor:=Cells(Column, j), Address:=coll.Item(i), TextToDisplay:="OK"
                End If
                If j <> 5 Then ' - снова оформление
                    Cells(Column, j).HorizontalAlignment = xlCenter
                    Cells(Column, j).VerticalAlignment = xlCenter
                End If
                If j = 5 Then ' - снова оформление
                    Cells(Column, j).HorizontalAlignment = xlLeft
                    Cells(Column, j).VerticalAlignment = xlTop
                    Cells(Column, j).Font.Size = 9
                End If
                Cells(Column, j).EntireColumn.AutoFit ' - автоматический подбор размера колонки
            Next
            Cells(Column, 8) = FileDateTime(coll.Item(i)) ' - добавляем время файла для будущих операций проверки
            Cells(Column, 9) = coll.Item(i) ' - добавляем путь файла для будущих операций проверки
        ElseIf (Proverka > 0) Then
            Column = Proverka
            'Переписывается таблица
            ImportData coll.Item(i), Column ' - считываются данные
            'Переписывается дата изменения файла
            Cells(Column, 8) = FileDateTime(coll.Item(i))
        End If
    Next
    
    Ended ' - включение всего ненужного
    Dim a As Integer
    a = MsgBox("Все обновлено", 0, "Наименование компании") ' - окно с объявлением что все прошло успешно
    
End Sub

Public Function EnumerationDateTime(ByVal CollectionName As String) As Integer ' - проверка файл на новый, старый и измененный старый файл
    Dim Path As String
    Dim DateTime As Date
    
    Path = CollectionName ' - путь файла
    DateTime = FileDateTime(CollectionName) ' - дата последнего изменеия файла
    
    For i = 2 To WorksheetFunction.CountA(Columns(1))
        If (Path = Cells(i, 9)) And (DateTime = Cells(i, 8)) Then ' - сравнивается пути и даты изменения файла
            EnumerationDateTime = -1 ' - Старый файл без изменений    если все одинаково то ставит значение -1
            Exit For
        ElseIf (Path = Cells(i, 9)) And (DateTime <> Cells(i, 8)) Then
            EnumerationDateTime = i ' - Старый файл, но с изменениями
            Exit For
        Else
            EnumerationDateTime = 0 ' - - Если новый файл
        End If
    Next
    
End Function

Public Function FilenamesCollection(ByVal FolderPath As String, Optional ByVal Mask As String = "") As Collection ' - извлечение нужных файлов
    'Optional ByVal SearchDeep As Long = 999 - доп настройка на глубину
    'необходимо добавить это свойство если надо искать в подпапках
    Set FilenamesCollection = New Collection    ' создаём пустую коллекцию
    Set FSO = CreateObject("Scripting.FileSystemObject")    ' создаём экземпляр FileSystemObject
    GetAllFileNamesUsingFSO FolderPath, Mask, FSO, FilenamesCollection ' поиск
    Set FSO = Nothing: Application.StatusBar = False    ' очистка строки состояния Excel
End Function

Public Function GetAllFileNamesUsingFSO(ByVal FolderPath As String, ByVal Mask As String, ByRef FSO, _
                                 ByRef FileNamesColl As Collection)
    'ByVal SearchDeep As Long - добавить свойство
    On Error Resume Next: Set curfold = FSO.GetFolder(FolderPath)
    If Not curfold Is Nothing Then    ' если удалось получить доступ к папке

        For Each fil In curfold.Files    ' перебираем все файлы в папке FolderPath
            If fil.Name Like "[!~]*" & Mask Then FileNamesColl.Add fil.Path
        Next
        'SearchDeep = SearchDeep - 1  - уменьшаем глубину поиска в подпапках
        'If SearchDeep Then    ' если надо искать глубже
        '    For Each sfol In curfold.SubFolders    ' ' перебираем все подпапки в папке FolderPath
        '        GetAllFileNamesUsingFSO sfol.Path, Mask, FSO, FileNamesColl, SearchDeep
        '    Next
        'End If
        
        Set fil = Nothing: Set curfold = Nothing    ' очищаем переменные
    End If
End Function
Public Function ImportData(ByVal FilePath As String, ByVal Column As Integer) ' - извлечение данных из определенных ячеек и копирование в массив

    Dim sShName As String, sAddress As String
    Dim objCloseBook As Object
    Dim ArrayRange As Variant
    Dim Row As Integer
    
    Application.ScreenUpdating = False
    Set objCloseBook = GetObject(FilePath) ' - открытие файла
    ArrayRange = Array(objCloseBook.Sheets(1).Range("A39").Value, objCloseBook.Sheets(1).Range("E11").Value, objCloseBook.Sheets(1).Range("A19").Value, objCloseBook.Sheets(1).Range("A17").Value)
    objCloseBook.Close False
    
    Row = 3
    For Each element In ArrayRange
        Cells(Column, Row) = element
        Row = Row + 1
    Next
    
End Function


Public Sub Ended() ' - включение всего ненужного
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
End Sub
