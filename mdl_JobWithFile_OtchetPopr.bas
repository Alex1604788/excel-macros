Attribute VB_Name = "mdl_JobWithFile_OtchetPopr"


Public Type mFileOtchetPopr_

    mName                   As String
    mRange                  As Range
    mNumRowStart            As Long
    mNumRowTitle            As Long
    
    mNachislType            As mColumn_
    mSKU                    As mColumn_
    mArt                    As mColumn_
    mCount                  As mColumn_
    mZaProdDoVichitKomm     As mColumn_
    mFinalSumm              As mColumn_
    
    mShortNameShForArc      As String
    
    mBoook                  As Workbook
    mOpenedBefore           As Boolean
    mArrIndex()             As Long
    
End Type

Public Sub mSettingsFileOtchetPopr()
    
    With mImportedFile.mFileOtchetPopr
        .mName = "Отчет по товарам"
        .mOpenedBefore = False
        
        .mShortNameShForArc = "ОтчПоТовар"
        
        .mNachislType.mCaption = f_str_FindParam("A0041")
        .mSKU.mCaption = f_str_FindParam("A0042")
        .mArt.mCaption = f_str_FindParam("A0043")
        .mCount.mCaption = f_str_FindParam("A0044")
        .mZaProdDoVichitKomm.mCaption = f_str_FindParam("A0045")
        .mFinalSumm.mCaption = f_str_FindParam("A0046")
    End With
    
End Sub

Public Function mOpenFileOtchetPopr() As Boolean
    
    mOpenFileOtchetPopr = False
    
    Dim mArrTemp1()             As String
    Dim mPath                   As String
    Dim mFindedSheet            As mFindedSheetInBook_
    Dim mColNamesString         As String
    
    mArrTemp1 = mFileDialog(msoFileDialogFilePicker, "Укажите файл " & mImportedFile.mFileOtchetPopr.mName, _
        "Выбрать", False, ThisWorkbook.Path, "Файлы Excel,*.xlsx*;*.xlsm*;*.xlsb*;*.csv*")
    
    If -1 = Not mArrTemp1 Then Exit Function
    mPath = mArrTemp1(0)
    If Len(mPath) = 0 Then Exit Function
    
    mColNamesString = mColNamesString & mImportedFile.mFileOtchetPopr.mNachislType.mCaption & vbNewLine
    mColNamesString = mColNamesString & mImportedFile.mFileOtchetPopr.mSKU.mCaption & vbNewLine
    mColNamesString = mColNamesString & mImportedFile.mFileOtchetPopr.mArt.mCaption & vbNewLine
    mColNamesString = mColNamesString & mImportedFile.mFileOtchetPopr.mCount.mCaption & vbNewLine
    mColNamesString = mColNamesString & mImportedFile.mFileOtchetPopr.mZaProdDoVichitKomm.mCaption & vbNewLine
    mColNamesString = mColNamesString & mImportedFile.mFileOtchetPopr.mFinalSumm.mCaption
    
    If mBookIsOpen(mPath) = True Then mImportedFile.mFileOtchetPopr.mOpenedBefore = True
    Set mImportedFile.mFileOtchetPopr.mBoook = f_OpenBook_book(mPath): If mImportedFile.mFileOtchetPopr.mBoook Is Nothing Then Exit Function
    
    mFindedSheet = mFindSheetInBookByColNames(mImportedFile.mFileOtchetPopr.mBoook, mColNamesString)
    If mFindedSheet.mNumRowTitle = 0 Then
        MsgBox "В файле " & mImportedFile.mFileOtchetPopr.mBoook.Name & vbCrLf & _
            "Не удалось найти необходимого листа с данными" & vbCrLf & _
            "Проверьте наличие необходимых колонок в нем" & vbCrLf & _
            "Работа программы завершена", vbCritical
        If mImportedFile.mFileOtchetPopr.mOpenedBefore = False Then mImportedFile.mFileOtchetPopr.mBoook.Close False
        Exit Function
    End If
    
    mImportedFile.mFileOtchetPopr.mNumRowTitle = mFindedSheet.mNumRowTitle
    mImportedFile.mFileOtchetPopr.mNumRowStart = mImportedFile.mFileOtchetPopr.mNumRowTitle + 1
    
    Set mImportedFile.mFileOtchetPopr.mRange = mImportedFile.mFileOtchetPopr.mBoook.Sheets(mFindedSheet.mSheetName).UsedRange
    If mImportedFile.mFileOtchetPopr.mRange Is Nothing Then Exit Function
    
    With mImportedFile.mFileOtchetPopr
        
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mNachislType
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mSKU
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mArt
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mCount
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mZaProdDoVichitKomm
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mFinalSumm
        
    End With
    
    mOpenFileOtchetPopr = True
    
End Function

Public Function mFindInRangeOtchetPopr(ByRef mIndexedArt As Long, ByRef mArt As String, _
ByRef mArrTypes() As String, ByVal mNumColResult As Long, ByRef mIndexedType() As Long) As Double
    
    Dim mRes            As Double
    Dim mIndArr         As Long
    Dim mIndType        As Long
    
    With mImportedFile.mFileOtchetPopr
        For NumRow = .mNumRowStart To .mRange.Rows.Count
            If .mArrIndex(0, NumRow) = mIndexedArt Then
            If f_str_ValToStr(.mRange(NumRow, .mArt.mNum)) = mArt Then
                
                For mIndType = 0 To UBound(mIndexedType)
                    If .mArrIndex(1, NumRow) = mIndexedType(mIndType) Then
                    If f_str_ValToStr(.mRange(NumRow, .mNachislType.mNum)) = Split(mArrTypes(mIndType), "@")(0) Then
                        
                        mRes = mRes + (f_dbl_ValToDbl(.mRange(NumRow, mNumColResult)) * Split(mArrTypes(mIndType), "@")(1))
                        
                    End If
                    End If
                Next
                
            End If
            End If
        Next
    End With
    
    mFindInRangeOtchetPopr = mRes
    
End Function
