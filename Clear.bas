Attribute VB_Name = "Module1"
Sub Delete() ' - отчистка и нормальзаци€ всех €чеек, кроме первой строки
    Rows("2:" & Rows.Count).ClearContents
    Rows("2:" & Rows.Count).Style = "Normal"
End Sub
