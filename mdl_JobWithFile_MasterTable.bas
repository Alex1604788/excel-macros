Attribute VB_Name = "mdl_JobWithFile_MasterTable"

Public Type mFileMasterTable_

    mName                   As String
    mRange                  As Range
    mNumRowStart            As Long
    mNumRowTitle            As Long
    
    mColOzonSKU             As mColumn_
    mColArt                 As mColumn_
    mColPost                As mColumn_
    mColCaption             As mColumn_
    mPoductType             As mColumn_
    mColEI                  As mColumn_
    mColCategory            As mColumn_
    mColCostZakup           As mColumn_
    mColOstSkladTek         As mColumn_
    mColOstAndTVP           As mColumn_
        
    mShortNameShForArc      As String
    
    mBoook                  As Workbook
    mOpenedBefore           As Boolean
    mArrIndex()             As Long
    
End Type

Public Sub mSettingsFileMasterTable()
    
    With mImportedFile.mFileMasterTable
        
        .mName = "Мастер таблица"
        .mOpenedBefore = False
        
        .mShortNameShForArc = "МастерТабл"
        
        .mColOzonSKU.mCaption = f_str_FindParam("A0021")
        .mColArt.mCaption = f_str_FindParam("A0022")
        .mColPost.mCaption = f_str_FindParam("A0023")
        .mColCaption.mCaption = f_str_FindParam("A0024")
        .mPoductType.mCaption = f_str_FindParam("A0025")
        .mColEI.mCaption = f_str_FindParam("A0026")
        .mColCategory.mCaption = f_str_FindParam("A0027")
        .mColCostZakup.mCaption = f_str_FindParam("A0028")
        .mColOstSkladTek.mCaption = f_str_FindParam("A0080")
        .mColOstAndTVP.mCaption = f_str_FindParam("A0081")
        
    End With
    
End Sub

Public Function mOpenFileMasterTable() As Boolean
    
    mOpenFileMasterTable = False
    
    Dim mArrTemp1()             As String
    Dim mPath                   As String
    Dim mFindedSheet            As mFindedSheetInBook_
    Dim mColNamesString         As String
    
    mArrTemp1 = mFileDialog(msoFileDialogFilePicker, "Укажите файл " & mImportedFile.mFileMasterTable.mName, _
        "Выбрать", False, ThisWorkbook.Path, "Файлы Excel,*.xlsx*;*.xlsm*;*.xlsb*;*.csv*")
    
    If -1 = Not mArrTemp1 Then Exit Function
    mPath = mArrTemp1(0)
    If Len(mPath) = 0 Then Exit Function
    
    mColNamesString = mColNamesString & mImportedFile.mFileMasterTable.mColOzonSKU.mCaption & vbNewLine
    mColNamesString = mColNamesString & mImportedFile.mFileMasterTable.mColArt.mCaption & vbNewLine
    mColNamesString = mColNamesString & mImportedFile.mFileMasterTable.mColPost.mCaption & vbNewLine
    mColNamesString = mColNamesString & mImportedFile.mFileMasterTable.mColCaption.mCaption & vbNewLine
    mColNamesString = mColNamesString & mImportedFile.mFileMasterTable.mPoductType.mCaption & vbNewLine
    mColNamesString = mColNamesString & mImportedFile.mFileMasterTable.mColEI.mCaption & vbNewLine
    mColNamesString = mColNamesString & mImportedFile.mFileMasterTable.mColCategory.mCaption & vbNewLine
    mColNamesString = mColNamesString & mImportedFile.mFileMasterTable.mColOstSkladTek.mCaption & vbNewLine
    mColNamesString = mColNamesString & mImportedFile.mFileMasterTable.mColOstAndTVP.mCaption & vbNewLine
    mColNamesString = mColNamesString & mImportedFile.mFileMasterTable.mColCostZakup.mCaption
    
    If mBookIsOpen(mPath) = True Then mImportedFile.mFileMasterTable.mOpenedBefore = True
    Set mImportedFile.mFileMasterTable.mBoook = f_OpenBook_book(mPath): If mImportedFile.mFileMasterTable.mBoook Is Nothing Then Exit Function
    
    mFindedSheet = mFindSheetInBookByColNames(mImportedFile.mFileMasterTable.mBoook, mColNamesString)
    If mFindedSheet.mNumRowTitle = 0 Then
        MsgBox "В файле " & mImportedFile.mFileMasterTable.mBoook.Name & vbCrLf & _
            "Не удалось найти необходимого листа с данными" & vbCrLf & _
            "Проверьте наличие необходимых колонок в нем" & vbCrLf & _
            "Работа программы завершена", vbCritical
        If mImportedFile.mFileMasterTable.mOpenedBefore = False Then mImportedFile.mFileMasterTable.mBoook.Close False
        Exit Function
    End If
    
    mImportedFile.mFileMasterTable.mNumRowTitle = mFindedSheet.mNumRowTitle
    mImportedFile.mFileMasterTable.mNumRowStart = mImportedFile.mFileMasterTable.mNumRowTitle + 2
    
    Set mImportedFile.mFileMasterTable.mRange = mImportedFile.mFileMasterTable.mBoook.Sheets(mFindedSheet.mSheetName).UsedRange
    If mImportedFile.mFileMasterTable.mRange Is Nothing Then Exit Function
    
    With mImportedFile.mFileMasterTable
        
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mColOzonSKU
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mColArt
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mColPost
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mColCaption
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mPoductType
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mColEI
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mColCategory
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mColCostZakup
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mColOstSkladTek
        mCheckCol .mName, .mRange, .mNumRowTitle, .mRange.Parent.Name, True, False, .mColOstAndTVP
        
    End With
    
    mOpenFileMasterTable = True
    
End Function

Public Function mFindRowInMasterTableByProdName(ByVal mProdName As String) As Long
    
    Dim mIndexedProdName            As Long
    
    mIndexedProdName = CreateIndex(mProdName)
    
    With mImportedFile.mFileMasterTable
        For NumRow = .mNumRowStart To .mRange.Rows.Count
            If .mArrIndex(2, NumRow) = mIndexedProdName Then
            If .mRange(NumRow, .mColCaption.mNum) = mProdName Then
                mFindRowInMasterTableByProdName = NumRow
                Exit Function
            End If
            End If
        Next
    End With
    
End Function


