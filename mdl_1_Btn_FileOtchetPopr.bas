Attribute VB_Name = "mdl_1_Btn_FileOtchetPopr"

Public Sub mLoadOtchetPopr()
    
    mFlagJob = True
    SetActivityExcel False
    Settings
    
    '----------------------------------------------------------------------------------------
    
    Dim mTempStr_1                  As String
    
    mTempStr_1 = Application.Caller
    
    If mOpenFileOtchetPopr = False Then
        StopMySub True, True
    End If
    
    ThisWorkbook.Activate
    
    mCreateArrIndexAll mSheetI.mColArt.mNum, True, False, False, False, True
    If mFillShI(mTempStr_1) = False Then
        MsgBox "Нет данных для добавления", vbCritical
        If mImportedFile.mFileOtchetPopr.mOpenedBefore = False Then
            mImportedFile.mFileOtchetPopr.mBoook.Close False
        End If
        StopMySub True, True
    End If
    
    If mTempStr_1 = "Кнопка 4" Then
        mTempStr_1 = Replace(Replace(Replace(mNow, ".", ""), " ", ""), "-", "") & " " & _
            mImportedFile.mFileOtchetPopr.mShortNameShForArc & " 1"
    Else
        mTempStr_1 = Replace(Replace(Replace(mNow, ".", ""), " ", ""), "-", "") & " " & _
            mImportedFile.mFileOtchetPopr.mShortNameShForArc & " 2"
    End If
    
    '----------------------------------------------------------------------------------------
    
    mImportedFile.mFileOtchetPopr.mBoook.Sheets(mImportedFile.mFileOtchetPopr.mRange.Parent.Name).Copy _
        After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = mTempStr_1
    
    If mImportedFile.mFileOtchetPopr.mOpenedBefore = False Then
        mImportedFile.mFileOtchetPopr.mBoook.Close False
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
    
    Dim mArt                            As String
    Dim mIndexedArt                     As Long
    Dim mRowShI                         As Long
    
    Dim mArrTypesForEkvaring()          As String
    'Dim mArrTypesForProdInCeniOzon()    As String
    Dim mArrTypesForNachislenie()       As String
    Dim mArrTypesForCount()             As String
    
    Dim mArrIndexTypesForEkvaring()          As Long
    'Dim mArrIndexTypesForProdInCeniOzon()    As Long
    Dim mArrIndexTypesForNachislenie()       As Long
    Dim mArrIndexTypesForCount()             As Long
    
    mArrTypesForEkvaring = Split(mSheetI.mTipNachislForEkvarung, "#")
    'mArrTypesForProdInCeniOzon = Split("Доставка покупателю@1#Доставка и обработка возврата, отмены, невыкупа@1#Доставка покупателю — отмена начисления@1", "#")
    mArrTypesForNachislenie = Split(mSheetI.mTipNachislForNachislen, "#")
    mArrTypesForCount = Split(mSheetI.mTipNachislForCount, "#")
    
    ReDim mArrIndexTypesForEkvaring(UBound(mArrTypesForEkvaring))
    For mIndArr = 0 To UBound(mArrIndexTypesForEkvaring)
        mArrIndexTypesForEkvaring(mIndArr) = CreateIndex(Split(mArrTypesForEkvaring(mIndArr), "@")(0))
    Next
    
    'ReDim mArrIndexTypesForProdInCeniOzon(UBound(mArrTypesForProdInCeniOzon))
    'For mIndArr = 0 To UBound(mArrIndexTypesForProdInCeniOzon)
    '    mArrIndexTypesForProdInCeniOzon(mIndArr) = CreateIndex(Split(mArrTypesForProdInCeniOzon(mIndArr), "@")(0))
    'Next
    
    ReDim mArrIndexTypesForNachislenie(UBound(mArrTypesForNachislenie))
    For mIndArr = 0 To UBound(mArrIndexTypesForNachislenie)
        mArrIndexTypesForNachislenie(mIndArr) = CreateIndex(Split(mArrTypesForNachislenie(mIndArr), "@")(0))
    Next
    
    ReDim mArrIndexTypesForCount(UBound(mArrTypesForCount))
    For mIndArr = 0 To UBound(mArrIndexTypesForCount)
        mArrIndexTypesForCount(mIndArr) = CreateIndex(Split(mArrTypesForCount(mIndArr), "@")(0))
    Next
    
    For mRowShI = mSheetI.mNumRowStart To mSheetI.mRange.Rows.Count
        If mSheetI.mRange(mRowShI, 1).Interior.Color <> mColors.mShPostGroup1 Then
        If mSheetI.mRange(mRowShI, 1).Interior.Color <> mColors.mShPostGroup2 Then
            
            mTxtBar = "Гружу отчет по товарам, строка " & mRowShI & " из " & mSheetI.mRange.Rows.Count
            Application.StatusBar = mTxtBar
            DoEvents
            
            mArt = mSheetI.mRange(mRowShI, mSheetI.mColArt.mNum)
            mIndexedArt = CreateIndex(mArt)
                        
            'mSheetI.mRange(mRowShI, mSheetI.mNacenProcentItog.mNum) = _
            '    f_dbl_ValToDbl(mSheetI.mRange(mRowShI, mSheetI.mNacenProcentItog.mNum)) + _
            '    mFindInRangeOtchetPopr(mIndexedArt, mArt, mArrTypesForProdInCeniOzon, _
            '    mImportedFile.mFileOtchetPopr.mZaProdDoVichitKomm.mNum, mArrIndexTypesForProdInCeniOzon)
            
            If mCaller = "Кнопка 4" Then
                
                mSheetI.mRange(mRowShI, mSheetI.mColSellsMoneyPer1.mNum) = _
                    mFindInRangeOtchetPopr(mIndexedArt, mArt, mArrTypesForNachislenie, _
                    mImportedFile.mFileOtchetPopr.mFinalSumm.mNum, mArrIndexTypesForNachislenie)
                
                mSheetI.mRange(mRowShI, mSheetI.mColSellsCountPer1.mNum) = _
                    mFindInRangeOtchetPopr(mIndexedArt, mArt, mArrTypesForCount, _
                    mImportedFile.mFileOtchetPopr.mCount.mNum, mArrIndexTypesForCount)
                
                mSheetI.mRange(mRowShI, mSheetI.mColEkvaring1.mNum) = _
                    mFindInRangeOtchetPopr(mIndexedArt, mArt, mArrTypesForEkvaring, _
                    mImportedFile.mFileOtchetPopr.mFinalSumm.mNum, mArrIndexTypesForEkvaring)
                
            Else
                
                mSheetI.mRange(mRowShI, mSheetI.mColSellsMoneyPer2.mNum) = _
                    mFindInRangeOtchetPopr(mIndexedArt, mArt, mArrTypesForNachislenie, _
                    mImportedFile.mFileOtchetPopr.mFinalSumm.mNum, mArrIndexTypesForNachislenie)
                
                mSheetI.mRange(mRowShI, mSheetI.mColSellsCountPer2.mNum) = _
                    mFindInRangeOtchetPopr(mIndexedArt, mArt, mArrTypesForCount, _
                    mImportedFile.mFileOtchetPopr.mCount.mNum, mArrIndexTypesForCount)
                
                mSheetI.mRange(mRowShI, mSheetI.mColEkvaring2.mNum) = _
                    mFindInRangeOtchetPopr(mIndexedArt, mArt, mArrTypesForEkvaring, _
                    mImportedFile.mFileOtchetPopr.mFinalSumm.mNum, mArrIndexTypesForEkvaring)
                    
            End If
                        
        End If
        End If
    Next
    
    Application.StatusBar = False
    
    mFillShI = True
    
End Function





