Attribute VB_Name = "Module1"
Sub ExportAllVbaModules()
    Dim vbComp As Object
    Dim exportPath As String
    Dim ext As String
    
    ' Папка, куда сложим все файлы с кодом
    exportPath = "C:\VBA_Export\"   ' поменяй путь, если нужно

    On Error Resume Next
    MkDir exportPath
    On Error GoTo 0
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1 ' Стандартный модуль
                ext = ".bas"
            Case 2 ' Класс
                ext = ".cls"
            Case 3 ' Форма
                ext = ".frm"
            Case 100 ' ThisWorkbook/Sheet
                ext = ".cls"
            Case Else
                ext = ".txt"
        End Select
        
        vbComp.Export exportPath & vbComp.Name & ext
    Next vbComp
    
    MsgBox "Готово! Модули выгружены в: " & exportPath
End Sub

