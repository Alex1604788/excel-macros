Attribute VB_Name = "mdl_FindTableInBook"
Public Function mFindSheetInBookByColNames(ByRef mBook As Workbook, _
ByVal mColNamesString As String, Optional ByVal mSheetName As String, _
Optional ByVal mPersentGoodJob As Long) As mFindedSheetInBook_
        
    Dim mArrColNames()          As String
    Dim mNumCol                 As Long
    Dim mRange                  As Range
    
    Dim mFindedColsCount        As Long
    Dim mArrPersentsForSh()     As Double
    Dim mArrNumRowsForSh()      As Long
    
    Dim mIndexBestSh            As Long
    'Dim mNumRowTitle            As Long
    
    Dim mTempStr1               As String
    
    mArrColNames = Split(mColNamesString, vbNewLine)
    mFindedColsCount = 0
    If mPersentGoodJob = 0 Then mPersentGoodJob = 65
    
    For mIndColName = 0 To UBound(mArrColNames)
        mArrColNames(mIndColName) = UCase(mArrColNames(mIndColName))
    Next
    
    ' find inf for all sh
    For Each mSh In mBook.Sheets
    
        
        If Len(mSheetName) > 0 Then
            If mSh.Name <> mSheetName Then
                GoTo Line_NextSheet
            End If
        End If
        
        Set mRange = mSh.UsedRange
        mFindedColsCount = 0
        
        If -1 = Not mArrPersentsForSh Then
            ReDim mArrPersentsForSh(0)
            ReDim mArrNumRowsForSh(0)
        Else
            ReDim Preserve mArrPersentsForSh(UBound(mArrPersentsForSh) + 1)
            ReDim Preserve mArrNumRowsForSh(UBound(mArrNumRowsForSh) + 1)
        End If
        
        mTempStr1 = mFindRowTitle(mRange, mArrColNames, mPersentGoodJob)
        'mNumRowTitle = Split(mTempStr1, vbNewLine)(0) * 1
        mArrNumRowsForSh(UBound(mArrNumRowsForSh)) = Split(mTempStr1, vbNewLine)(0) * 1
        mArrPersentsForSh(UBound(mArrPersentsForSh)) = Split(mTempStr1, vbNewLine)(1) * 1
        
        'If mNumRowTitle = 0 Then
        '    mArrPersentsForSh(UBound(mArrPersentsForSh)) = 0
        'Else
        '    For mNumCol = 1 To mRange.Columns.Count
        '        For mIndColName = 0 To UBound(mArrColNames)
        '            If UCase(mRange(mNumRowTitle, mNumCol)) = mArrColNames(mIndColName) Then
        '                mFindedColsCount = mFindedColsCount + 1
        '                Exit For
        '            End If
        '        Next
        '    Next
        '    mArrPersentsForSh(UBound(mArrPersentsForSh)) = (mFindedColsCount * 100) / (UBound(mArrColNames) + 1)
        'End If
        
Line_NextSheet:
        
    Next
    
    ' find best sh
    
    mIndexBestSh = 0
    
    If -1 = Not mArrPersentsForSh Then
        mFindSheetInBookByColNames.mNumRowTitle = 0
        GoTo Line_Exit
    End If
    
    For mIndArr = 0 To UBound(mArrPersentsForSh)
        If mArrPersentsForSh(mIndArr) > mArrPersentsForSh(mIndexBestSh) Then
            mIndexBestSh = mIndArr
        End If
    Next
    
    If mArrPersentsForSh(mIndexBestSh) >= mPersentGoodJob Then
        mFindSheetInBookByColNames.mNumRowTitle = mArrNumRowsForSh(mIndexBestSh)
        If Len(mSheetName) > 0 Then
            mFindSheetInBookByColNames.mSheetName = mSheetName
        Else
            mFindSheetInBookByColNames.mSheetName = mBook.Sheets(mIndexBestSh + 1).Name
        End If
    End If
    
Line_Exit:
    
    Set mRange = Nothing
    Erase mArrColNames
    Erase mArrPersentsForSh
    Erase mArrNumRowsForSh
    
End Function

Private Function mFindRowTitle(ByRef mRange As Range, ByRef mArrColNames() As String, _
ByVal mPersentGoodJob As Long) As String
    
    Dim mMaxRowsToSearch            As Long
    Dim mArrPersentsForRows()       As Double
    Dim mNumCol                     As Long
    Dim mIndColName                 As Long
    Dim mNumRow                     As Long
    Dim mFindedColsCount            As Long
    Dim mIndBestRow                 As Long
    Dim mIndFirstTitleRow           As Long
    
    mMaxRowsToSearch = 30
    ReDim mArrPersentsForRows(mMaxRowsToSearch - 1)
    mIndFirstTitleRow = -1
    
    For mNumRow = 1 To mMaxRowsToSearch
        
        mFindedColsCount = 0
        
        For mNumCol = 1 To mRange.Columns.Count
            For mIndColName = 0 To UBound(mArrColNames)
                If UCase(f_str_ValToStr(mRange(mNumRow, mNumCol))) = mArrColNames(mIndColName) Then
                    mFindedColsCount = mFindedColsCount + 1
                    'Exit For
                End If
            Next
        Next
        
        mArrPersentsForRows(mNumRow - 1) = (mFindedColsCount * 100) / (UBound(mArrColNames) + 1)
        
    Next
    
    mIndBestRow = 0
    For mIndArr = 0 To UBound(mArrPersentsForRows)
        If mArrPersentsForRows(mIndArr) > mArrPersentsForRows(mIndBestRow) Then
            mIndBestRow = mIndArr
        End If
        If mIndFirstTitleRow = -1 Then
            If mArrPersentsForRows(mIndArr) > 0 Then
                mIndFirstTitleRow = mIndArr
            End If
        End If
    Next
    
    If mArrPersentsForRows(mIndBestRow) >= mPersentGoodJob Then
        mFindRowTitle = mIndFirstTitleRow + 1 & vbNewLine & mArrPersentsForRows(mIndBestRow)
    Else
        mFindRowTitle = "0" & vbNewLine & "0"
    End If
    
    Erase mArrPersentsForRows
    
End Function

