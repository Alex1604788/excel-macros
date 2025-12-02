Attribute VB_Name = "mdl_1_Btn_FileStoimRazmesh"

Public Sub mLoadStoimRazmesh()
    
    mFlagJob = True
    SetActivityExcel False
    Settings
    
    '----------------------------------------------------------------------------------------
    
    Dim mTempStr_1                  As String
    
    mTempStr_1 = Application.Caller
    
    If mOpenFileStoimRazmesh = False Then
        StopMySub True, True
    End If
    
    ThisWorkbook.Activate
    
    mCreateArrIndexAll mSheetI.mColArt.mNum, True, False, True, False, False
    If mFillShI(mTempStr_1) = False Then
        MsgBox "Нет данных для добавления", vbCritical
        If mImportedFile.mFileStoimRazmesh.mOpenedBefore = False Then
            mImportedFile.mFileStoimRazmesh.mBoook.Close False
        End If
        StopMySub True, True
    End If
    
    If mTempStr_1 = "Кнопка 3" Then
        mTempStr_1 = Replace(Replace(Replace(mNow, ".", ""), " ", ""), "-", "") & " " & _
            mImportedFile.mFileStoimRazmesh.mShortNameShForArc & " 1"
    Else
        mTempStr_1 = Replace(Replace(Replace(mNow, ".", ""), " ", ""), "-", "") & " " & _
            mImportedFile.mFileStoimRazmesh.mShortNameShForArc & " 2"
    End If
    
    '----------------------------------------------------------------------------------------
    
    mImportedFile.mFileStoimRazmesh.mBoook.Sheets(mImportedFile.mFileStoimRazmesh.mRange.Parent.Name).Copy _
        After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = mTempStr_1
    
    If mImportedFile.mFileStoimRazmesh.mOpenedBefore = False Then
        mImportedFile.mFileStoimRazmesh.mBoook.Close False
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
    Dim mLastDate                   As Date
    
    If mCaller = "Кнопка 7" Then
        mLastDate = mFindInRangeStoimRazmeshLastOfDate
    End If
    
    For mRowShI = mSheetI.mNumRowStart To mSheetI.mRange.Rows.Count
        If mSheetI.mRange(mRowShI, 1).Interior.Color <> mColors.mShPostGroup1 Then
        If mSheetI.mRange(mRowShI, 1).Interior.Color <> mColors.mShPostGroup2 Then
            
            If mCaller = "Кнопка 3" Then ' 3 кнопка это первый период, у второго периода 7 кнопка
                
                mSheetI.mRange(mRowShI, mSheetI.mColStoimRazmes1.mNum) = _
                    mFindInRangeStoimRazmesh(mSheetI.mArrIndex(mRowShI), _
                    mSheetI.mRange(mRowShI, mSheetI.mColArt.mNum), _
                    mImportedFile.mFileStoimRazmesh.mStoimRazm.mNum)
                
            Else
                
                mSheetI.mRange(mRowShI, mSheetI.mColStoimRazmes2.mNum) = _
                    mFindInRangeStoimRazmesh(mSheetI.mArrIndex(mRowShI), _
                    mSheetI.mRange(mRowShI, mSheetI.mColArt.mNum), _
                    mImportedFile.mFileStoimRazmesh.mStoimRazm.mNum)
                
                mSheetI.mRange(mRowShI, mSheetI.mColOstatokTekyshOZON.mNum) = _
                    mFindInRangeStoimRazmeshForLastDate(mSheetI.mArrIndex(mRowShI), _
                    mSheetI.mRange(mRowShI, mSheetI.mColArt.mNum), _
                    mImportedFile.mFileStoimRazmesh.mColOstatokTekyshOZON.mNum, _
                    mLastDate)
                
            End If
                                
        End If
        End If
    Next
    
    mFillShI = True
    
End Function

