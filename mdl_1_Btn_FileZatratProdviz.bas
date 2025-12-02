Attribute VB_Name = "mdl_1_Btn_FileZatratProdviz"

Public Sub mLoadZatratProdviz()
    
    mFlagJob = True
    SetActivityExcel False
    Settings
    
    '----------------------------------------------------------------------------------------
    
    Dim mTempStr_1                  As String
    
    mTempStr_1 = Application.Caller
    
    If mOpenFileZatratProdviz = False Then
        StopMySub True, True
    End If
    
    ThisWorkbook.Activate
    
    mCreateArrIndexAll mSheetI.mColArt.mNum, True, False, False, True, False
    If mFillShI(mTempStr_1) = False Then
        MsgBox "Нет данных для добавления", vbCritical
        If mImportedFile.mFileZatratProdviz.mOpenedBefore = False Then
            mImportedFile.mFileZatratProdviz.mBoook.Close False
        End If
        StopMySub True, True
    End If
    
    If mTempStr_1 = "Кнопка 2" Then
        mTempStr_1 = Replace(Replace(Replace(mNow, ".", ""), " ", ""), "-", "") & " " & _
            mImportedFile.mFileZatratProdviz.mShortNameShForArc & " 1"
    Else
        mTempStr_1 = Replace(Replace(Replace(mNow, ".", ""), " ", ""), "-", "") & " " & _
            mImportedFile.mFileOtchetPopr.mShortNameShForArc & " 2"
    End If
    
    '----------------------------------------------------------------------------------------
    
    mImportedFile.mFileZatratProdviz.mBoook.Sheets(mImportedFile.mFileZatratProdviz.mRange.Parent.Name).Copy _
        After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = mTempStr_1
    
    If mImportedFile.mFileZatratProdviz.mOpenedBefore = False Then
        mImportedFile.mFileZatratProdviz.mBoook.Close False
    End If
    
    '----------------------------------------------------------------------------------------
    
    ThisWorkbook.Sheets(mSheetI.mName).Select
    SetActivityExcel True
    mSetColWidthSheetI
    
    '----------------------------------------------------------------------------------------
    
Line_Exit:
    ResetVars
    StopMySub False, True
        
End Sub

Private Function mFillShI(ByVal mCaller As String) As Boolean
    
    Dim mRowShI                     As Long
    
    For mRowShI = mSheetI.mNumRowStart To mSheetI.mRange.Rows.Count
        If mSheetI.mRange(mRowShI, 1).Interior.Color <> mColors.mShPostGroup1 Then
        If mSheetI.mRange(mRowShI, 1).Interior.Color <> mColors.mShPostGroup2 Then
            
            If mCaller = "Кнопка 2" Then
                mSheetI.mRange(mRowShI, mSheetI.mColProdvizCost1.mNum) = _
                    f_dbl_ValToDbl(mFindInRangeZatratProdviz(CreateIndex(mSheetI.mRange(mRowShI, mSheetI.mColOzonSKU.mNum)), _
                    mSheetI.mRange(mRowShI, mSheetI.mColOzonSKU.mNum)))
            Else
                mSheetI.mRange(mRowShI, mSheetI.mColProdvizCost2.mNum) = _
                    f_dbl_ValToDbl(mFindInRangeZatratProdviz(CreateIndex(mSheetI.mRange(mRowShI, mSheetI.mColOzonSKU.mNum)), _
                    mSheetI.mRange(mRowShI, mSheetI.mColOzonSKU.mNum)))
            End If
            
        End If
        End If
    Next
    
    mFillShI = True
    
End Function



