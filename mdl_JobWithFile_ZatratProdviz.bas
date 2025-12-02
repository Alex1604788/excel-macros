Attribute VB_Name = "mdl_JobWithFile_ZatratProdviz"

Public Type mFileZatratProdviz_

    mName                   As String
    mRange                  As Range
    mNumRowStart            As Long
    mNumRowTitle            As Long
    
    mColOzonSKU             As mColumn_
    mProdvizType            As mColumn_
    mMoney                  As mColumn_
    
    mShortNameShForArc      As String
    
    mBoook                  As Workbook
    mOpenedBefore           As Boolean
    mArrIndex()             As Long
    
End Type

Public Sub mSettingsFileZatratProdviz()
    
    With mImportedFile.mFileZatratProdviz
        .mName = "Затраты на продвижение"
        .mOpenedBefore = False
        
        .mShortNameShForArc = "ЗатрНаПродв"
        
        .mColOzonSKU.mCaption = f_str_FindParam("A0037")
        .mProdvizType.mCaption = f_str_FindParam("A0038")
        .mMoney.mCaption = f_str_FindParam("A0039")
    End With
    
End Sub

Public Function mOpenFileZatratProdviz() As Boolean
    
    mOpenFileZatratProdviz = False
    
    Dim mArrTemp1()             As String
    Dim mPath                   As String
    Dim mFindedSheet            As mFindedSheetInBook_
    Dim mColNamesString         As String
    
    mArrTemp1 = mFileDialog(msoFileDialogFilePicker, "Укажите файл " & mImportedFile.mFileZatratProdviz.mName, _
        "Выбрать", False, ThisWorkbook.Path, "Файлы Excel,*.xlsx*;*.xlsm*;*.xlsb*;*.csv*")
    
    If -1 = Not mArrTemp1 Then Exit Function
    mPath = mArrTemp1(0)
    If Len(mPath) = 0 Then Exit Function
    
    mColNamesString = mColNamesString & mImportedFile.mFileZatratProdviz.mColOzonSKU.mCaption & vbNewLine
    mColNamesString = mColNamesString & mImportedFile.mFileZatratProdviz.mProdvizType.mCaption & vbNewLine
    mColNamesString = mColNamesString & mImportedFile.mFileZatratProdviz.mMoney.mCaption
    
    If mBookIsOpen(mPath) = True Then mImportedFile.mFileZatratProdviz.mOpenedBefore = True
    Set mImportedFile.mFileZatratProdviz.mBoook = f_OpenBook_book(mPath): If mImportedFile.mFileZatratProdviz.mBoook Is Nothing Then Exit Function
    
    mFindedSheet = mFindSheetInBookByColNames(mImportedFile.mFileZatratProdviz.mBoook, mColNamesString)
    If mFindedSheet.mNumRowTitle = 0 Then
        MsgBox "В файле " & mImportedFile.mFileZatratProdviz.mBoook.Name & vbCrLf & _
            "Не удалось найти необходимого листа с данными" & vbCrLf & _
            "Проверьте наличие необходимых колонок в нем" & vbCrLf & _
            "Работа программы завершена", vbCritical
        If mImportedFile.mFileZatratProdviz.mOpenedBefore = False Then mImportedFile.mFileZatratProdviz.mBoook.Close False
        Exit Function
    End If
    
    mImportedFile.mFileZatratProdviz.mNumRowTitle = mFindedSheet.mNumRowTitle
    mImportedFile.mFileZatratProdviz.mNumRowStart = mImportedFile.mFileZatratProdviz.mNumRowTitle + 1
    
    Set mImportedFile.mFileZatratProdviz.mRange = mImportedFile.mFileZatratProdviz.mBoook.Sheets(mFindedSheet.mSheetName).UsedRange
    If mImportedFile.mFileZatratProdviz.mRange Is Nothing Then Exit Function
    
    With mImportedFile.mFileZatratProdviz
        
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mColOzonSKU
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mProdvizType
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mMoney
        
    End With
    
    mOpenFileZatratProdviz = True
    
End Function

Public Function mFindInRangeZatratProdviz(ByRef mIndexedSKU As Long, ByRef mSKU As String) As Double
    
    Dim mRes            As Double
    
    With mImportedFile.mFileZatratProdviz
        For NumRow = .mNumRowStart To .mRange.Rows.Count
            If .mArrIndex(0, NumRow) = mIndexedSKU Then
            If f_str_ValToStr(.mRange(NumRow, .mColOzonSKU.mNum)) = mSKU Then
                mRes = mRes + f_dbl_ValToDbl(.mRange(NumRow, .mMoney.mNum))
            End If
            End If
        Next
    End With
    
    mFindInRangeZatratProdviz = mRes
    
End Function


