Attribute VB_Name = "mdl_JobWithFile__General"

Public Type mImportedFile_
    
    mFileMasterTable            As mFileMasterTable_
    mFileStoimRazmesh           As mFileStoimRazmesh_
    mFileZatratProdviz          As mFileZatratProdviz_
    mFileOtchetPopr             As mFileOtchetPopr_
    
End Type

Public mImportedFile        As mImportedFile_

Public Function mSettingsImportedFile()
    
    mSettingsFileMasterTable
    mSettingsFileStoimRazmesh
    mSettingsFileZatratProdviz
    mSettingsFileOtchetPopr
    
End Function

Public Function mNameOrNumCol(ByRef mCol As mColumn_, ByVal mParamText As String)
    
    Dim mTempStr            As String
    
    If InStr(mParamText, "@@@") > 0 Then
        
        mTempStr = Split(mParamText, "@@@")(1)
        If IsNumeric(mTempStr) = True Then
            mCol.mWidthToSet = mTempStr * 1
        End If
        
        mParamText = Split(mParamText, "@@@")(0)
        
    End If
    
    If mParamText Like "[#]*[#]" Then
        mTempStr = Split(mParamText, "#")(1)
        If f_bool_ItsNumber(mTempStr) = True Then
            mCol.mNum = mTempStr * 1
            mCol.mCaption = ""
        Else
            mCol.mNum = 0
            mCol.mCaption = mParamText
        End If
    Else
        mCol.mNum = 0
        mCol.mCaption = mParamText
    End If
        
End Function

Public Function mCloseImportedFile(ByRef mBook As Workbook, ByVal mOpenedBefore As Boolean)
    
    If mOpenedBefore = False Then mBook.Close False
    Set mBook = Nothing
    
End Function

Public Function mJobForFindNumCol(ByRef mImpFileRange As Range, ByVal mImpFileNumRowTitle As Long, _
ByVal mImpFileName As String, ByVal mNeedCheckNextRow As Boolean, _
ByVal mReversToFndCol As Boolean, ByRef mColumn As mColumn_)
    
    With mColumn
        If .mNum = 0 Then
        .mNum = mFindNumColInRange(mImpFileRange, mImpFileNumRowTitle, .mCaption, mNeedCheckNextRow, mReversToFndCol)
        If .mNum = -1 Then mColInFileNotFnd mImpFileName, .mCaption
    End If
    End With
    
End Function

Public Sub mResetVarsImportedFile(ByRef mRange As Range, ByVal mImpFileName As String)
    
    Set mRange = Nothing
    
    Select Case mImpFileName
        
        Case mImportedFile.mFileMasterTable.mName
            Erase mImportedFile.mFileMasterTable.mArrIndex
            mImportedFile.mFileMasterTable.mName = ""
            
        Case mImportedFile.mFileStoimRazmesh.mName
            Erase mImportedFile.mFileStoimRazmesh.mArrIndex
            mImportedFile.mFileStoimRazmesh.mName = ""
            
        Case mImportedFile.mFileZatratProdviz.mName
            Erase mImportedFile.mFileZatratProdviz.mArrIndex
            mImportedFile.mFileZatratProdviz.mName = ""
        
        Case mImportedFile.mFileOtchetPopr.mName
            Erase mImportedFile.mFileOtchetPopr.mArrIndex
            mImportedFile.mFileOtchetPopr.mName = ""
            
    End Select
    
End Sub

Public Function mErrToShErrImportedFile(ByVal mShName As String, _
ByVal mRow As Long, ByVal mErrCaption As String, _
ByRef mRangeDataErr As Range, ByVal mDataErrRow As Long, _
ByRef mArrErrDataColNum As Variant) As Long
    
    With ThisWorkbook.Sheets(mShName)
        
        If mRow = 0 Then
            .Cells.NumberFormat = "@"
            .Range("A1") = mNow
            .Range("A2") = "Это лист с перечнем ошибок произошедших во время обработки"
            .Range("A3") = mErrCaption
            .Range("A1:A3").Font.Bold = True
            .Range("A1:A3").Font.Color = vbRed
            mRow = 4
        End If
        
        mRow = mRow + 1
        For mIndCol = 1 To UBound(mArrErrDataColNum) + 1
            .Cells(mRow, mIndCol) = mRangeDataErr(mDataErrRow, mArrErrDataColNum(mIndCol - 1))
            .Cells(mRow, mIndCol) = mRangeDataErr(mDataErrRow, mArrErrDataColNum(mIndCol - 1))
        Next
    End With
    
    mErrToShErrImportedFile = mRow
    
End Function
