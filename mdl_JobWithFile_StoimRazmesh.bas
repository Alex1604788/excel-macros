Attribute VB_Name = "mdl_JobWithFile_StoimRazmesh"

Public Type mFileStoimRazmesh_

    mName                   As String
    mRange                  As Range
    mNumRowStart            As Long
    mNumRowTitle            As Long
    
    mArt                    As mColumn_
    mColOzonSKU             As mColumn_
    mStoimRazm              As mColumn_
    mColOstatokTekyshOZON   As mColumn_
    mColDate                As mColumn_
    
    mShortNameShForArc      As String
    
    mBoook                  As Workbook
    mOpenedBefore           As Boolean
    mArrIndex()             As Long
    
End Type

Public Sub mSettingsFileStoimRazmesh()
    
    With mImportedFile.mFileStoimRazmesh
        .mName = "Стоимость размещения по товарам"
        .mOpenedBefore = False
        
        .mShortNameShForArc = "СтоимРазм"
        
        .mArt.mCaption = f_str_FindParam("A0034")
        .mColOzonSKU.mCaption = f_str_FindParam("A0035")
        .mStoimRazm.mCaption = f_str_FindParam("A0036")
        .mColOstatokTekyshOZON.mCaption = f_str_FindParam("A0052")
        .mColDate.mCaption = f_str_FindParam("A0071")
        
    End With
    
End Sub

Public Function mOpenFileStoimRazmesh() As Boolean
    
    mOpenFileStoimRazmesh = False
    
    Dim mArrTemp1()             As String
    Dim mPath                   As String
    Dim mFindedSheet            As mFindedSheetInBook_
    Dim mColNamesString         As String
    
    mArrTemp1 = mFileDialog(msoFileDialogFilePicker, "Укажите файл " & mImportedFile.mFileStoimRazmesh.mName, _
        "Выбрать", False, ThisWorkbook.Path, "Файлы Excel,*.xlsx*;*.xlsm*;*.xlsb*;*.csv*")
    
    If -1 = Not mArrTemp1 Then Exit Function
    mPath = mArrTemp1(0)
    If Len(mPath) = 0 Then Exit Function
    
    mColNamesString = mColNamesString & mImportedFile.mFileStoimRazmesh.mArt.mCaption & vbNewLine
    mColNamesString = mColNamesString & mImportedFile.mFileStoimRazmesh.mColOzonSKU.mCaption & vbNewLine
    mColNamesString = mColNamesString & mImportedFile.mFileStoimRazmesh.mColOstatokTekyshOZON.mCaption & vbNewLine
    mColNamesString = mColNamesString & mImportedFile.mFileStoimRazmesh.mColDate.mCaption & vbNewLine
    mColNamesString = mColNamesString & mImportedFile.mFileStoimRazmesh.mStoimRazm.mCaption
    
    If mBookIsOpen(mPath) = True Then mImportedFile.mFileStoimRazmesh.mOpenedBefore = True
    Set mImportedFile.mFileStoimRazmesh.mBoook = f_OpenBook_book(mPath): If mImportedFile.mFileStoimRazmesh.mBoook Is Nothing Then Exit Function
    
    mFindedSheet = mFindSheetInBookByColNames(mImportedFile.mFileStoimRazmesh.mBoook, mColNamesString)
    If mFindedSheet.mNumRowTitle = 0 Then
        MsgBox "В файле " & mImportedFile.mFileStoimRazmesh.mBoook.Name & vbCrLf & _
            "Не удалось найти необходимого листа с данными" & vbCrLf & _
            "Проверьте наличие необходимых колонок в нем" & vbCrLf & _
            "Работа программы завершена", vbCritical
        If mImportedFile.mFileStoimRazmesh.mOpenedBefore = False Then mImportedFile.mFileStoimRazmesh.mBoook.Close False
        Exit Function
    End If
    
    mImportedFile.mFileStoimRazmesh.mNumRowTitle = mFindedSheet.mNumRowTitle
    mImportedFile.mFileStoimRazmesh.mNumRowStart = mImportedFile.mFileStoimRazmesh.mNumRowTitle + 1
    
    Set mImportedFile.mFileStoimRazmesh.mRange = mImportedFile.mFileStoimRazmesh.mBoook.Sheets(mFindedSheet.mSheetName).UsedRange
    If mImportedFile.mFileStoimRazmesh.mRange Is Nothing Then Exit Function
    
    With mImportedFile.mFileStoimRazmesh
        
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mArt
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mColOzonSKU
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mStoimRazm
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mColOstatokTekyshOZON
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mColDate
        
    End With
    
    mOpenFileStoimRazmesh = True
    
End Function

Public Function mFindInRangeStoimRazmesh(ByRef mIndexedArt As Long, ByRef mArt As String, _
ByVal mNumColFile As Long) As Double
    
    Dim mRes            As Double
    
    With mImportedFile.mFileStoimRazmesh
        For NumRow = .mNumRowStart To .mRange.Rows.Count
            If .mArrIndex(0, NumRow) = mIndexedArt Then
            If f_str_ValToStr(.mRange(NumRow, .mArt.mNum)) = mArt Then
                mRes = mRes + f_dbl_ValToDbl(.mRange(NumRow, mNumColFile))
            End If
            End If
        Next
    End With
    
    mFindInRangeStoimRazmesh = mRes
    
End Function

Public Function mFindInRangeStoimRazmeshForLastDate(ByRef mIndexedArt As Long, ByRef mArt As String, _
ByVal mNumColFile As Long, ByVal mLastDate As Date) As Double
    
    Dim mRes            As Double
    
    With mImportedFile.mFileStoimRazmesh
        For NumRow = .mNumRowStart To .mRange.Rows.Count
            If .mArrIndex(0, NumRow) = mIndexedArt Then
            If f_str_ValToStr(.mRange(NumRow, .mArt.mNum)) = mArt Then
            If CDate(.mRange(NumRow, .mColDate.mNum)) = mLastDate Then
                mRes = mRes + f_dbl_ValToDbl(.mRange(NumRow, mNumColFile))
            End If
            End If
            End If
        Next
    End With
    
    mFindInRangeStoimRazmeshForLastDate = mRes
    
End Function

Public Function mFindInRangeStoimRazmeshLastOfDate() As Date
    
    Dim mRes            As Date
    
    With mImportedFile.mFileStoimRazmesh
        For NumRow = .mNumRowStart To .mRange.Rows.Count
            If CDate(.mRange(NumRow, .mColDate.mNum)) > mRes Then
                mRes = CDate(.mRange(NumRow, .mColDate.mNum))
            End If
        Next
    End With
    
    mFindInRangeStoimRazmeshLastOfDate = mRes
    
End Function
