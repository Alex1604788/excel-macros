Attribute VB_Name = "mdl_1_Btn_FileMasterTable"

Public Sub mLoadMasterTable()
    
    mFlagJob = True
    SetActivityExcel False
    Settings
    
    '----------------------------------------------------------------------------------------
    
    Dim mTempStr_1                  As String
    
    If mOpenFileMasterTable = False Then
        StopMySub True, True
    End If
    
    mCheckFileMasterTable
    
    ThisWorkbook.Activate
    
    ClearSheetIam
    mCreateArrIndexAll 0, False, True, False, False, False
    If mFillShI = False Then
        MsgBox "Нет данных для добавления", vbCritical
        If mImportedFile.mFileMasterTable.mOpenedBefore = False Then
            mImportedFile.mFileMasterTable.mBoook.Close False
        End If
        StopMySub True, True
    End If
        
    mSetFormulsToShIForNoGroup
    mSummForGroups
    mSetFormatsAndGrd
    
    '----------------------------------------------------------------------------------------
    
    ' sheet copy
    mTempStr_1 = Replace(Replace(Replace(mNow, ".", ""), " ", ""), "-", "")
    mImportedFile.mFileMasterTable.mBoook.Sheets(mImportedFile.mFileMasterTable.mRange.Parent.Name).Copy _
        After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = mTempStr_1 & " " & _
        mImportedFile.mFileMasterTable.mShortNameShForArc
    
    If mImportedFile.mFileMasterTable.mOpenedBefore = False Then
        mImportedFile.mFileMasterTable.mBoook.Close False
    End If
    
    '----------------------------------------------------------------------------------------
    
    ThisWorkbook.Sheets(mSheetI.mName).Select
    SetActivityExcel True
    mSetColWidthSheetI
    ActiveSheet.Outline.ShowLevels RowLevels:=1
    
    '----------------------------------------------------------------------------------------
    
Line_Exit:
    ResetVars
    StopMySub False, True
        
End Sub

Private Function mCheckFileMasterTable()
    
    Dim mRow            As Long
    
    With mImportedFile.mFileMasterTable
        For mRow = .mNumRowStart To .mRange.Rows.Count
            If .mRange(mRow, .mColPost.mNum) = "" Then .mRange(mRow, .mColPost.mNum) = "не указан"
            If .mRange(mRow, .mPoductType.mNum) = "" Then .mRange(mRow, .mPoductType.mNum) = "не указан"
        Next
    End With
    
End Function

Private Function mFillShI() As Boolean
    
    Dim mRowShI                     As Long
    Dim mArrPost()                  As Variant
    Dim mIndArrPost                 As Long
    Dim mBarText                    As String
    
    mRowShI = mSheetI.mNumRowStart
    mArrPost = mFindAllPost
    
    If -1 = Not mArrPost Then
        mFillShI = False
        GoTo Line_Exit
    End If
        
    For mIndArrPost = 0 To UBound(mArrPost)
        
        mBarText = "Обработка поставщика " & mIndArrPost + 1 & " из " & UBound(mArrPost) + 1
        Application.StatusBar = mBarText
        DoEvents
        
        If mArrPost(mIndArrPost) <> "1" And _
        UCase(mArrPost(mIndArrPost)) <> mImportedFile.mFileMasterTable.mColPost.mCaption Then
            mRowShI = mAddDataPost(mArrPost(mIndArrPost), mRowShI, mBarText)
            mRowShI = mRowShI - 1
        End If
        
    Next
    Application.StatusBar = False
    
    mFillShI = True
    
Line_Exit:
        Erase mArrPost
    
End Function

Private Function mFindAllPost() As Variant()
    
    Dim mArr1()             As Variant
    Dim mArr2()             As Variant
    
    With mImportedFile.mFileMasterTable
        mArr1 = Range(.mRange(1, .mColPost.mNum), .mRange(.mRange.Rows.Count, .mColPost.mNum))
        mArr2 = f_SortArray_arrvar(mArr1)
        Erase mArr1
        mArr1 = f_SelectUniqueInArr_arrvar(mArr2)
    End With
    
    mFindAllPost = mArr1
        
    Erase mArr1
    Erase mArr2
        
End Function

Private Function mAddDataPost(ByVal mPostName As String, ByVal mRowShI As Long, ByVal mBarText As String) As Long
    
    Dim mFirstRowPost               As Long
    Dim mIndexedPost                As Long
    Dim mArrProductGroup()          As Variant
    Dim mIndProdGroup               As Long
    Dim mIndexedGroup               As Long
    Dim mTxtBar                     As String
    
    mFirstRowPost = mRowShI
    mIndexedPost = CreateIndex(mPostName)
    
    mSheetI.mRange(mFirstRowPost, mSheetI.mColOzonSKU.mNum) = mPostName
    mSheetI.mRange(mFirstRowPost, mSheetI.mColOzonSKU.mNum).EntireRow.Interior.Color = mColors.mShPostGroup1
    mArrProductGroup = mFindProductGroup(mIndexedPost, mPostName)
    mRowShI = mRowShI + 1
    
    For mIndProdGroup = 0 To UBound(mArrProductGroup)
        
        mTxtBar = mBarText & ", перебор группы " & _
            mIndProdGroup + 1 & " из " & UBound(mArrProductGroup) + 1
        Application.StatusBar = mTxtBar
        DoEvents
        
        mIndexedGroup = CreateIndex(mArrProductGroup(mIndProdGroup))
        mRowShI = mAddDataForGroup(mIndexedPost, mPostName, mIndexedGroup, mArrProductGroup(mIndProdGroup), mRowShI, mBarText)
        mRowShI = mRowShI - 1
        
    Next
    
    mSheetI.mRange.Rows(mRowShI - 1 & ":" & mFirstRowPost + 1).Group
    
    mRowShI = mRowShI + 1
    mAddDataPost = mRowShI
    
End Function

Private Function mFindProductGroup(ByVal mIndexedPost As Long, ByVal mPostName As String) As Variant()
    
    Dim mArr1()             As Variant
    Dim mArr2()             As Variant
    Dim mIndArr1            As Long
    
    Dim mNotNull            As Long
    Dim mArr3()             As Variant
    Dim mTempLong           As Long
    
    With mImportedFile.mFileMasterTable
        
        mArr1 = Range(.mRange(1, .mPoductType.mNum), .mRange(.mRange.Rows.Count, .mPoductType.mNum))
        
        For mIndArr1 = 1 To UBound(mArr1, 1)
            'If .mArrIndex(0, mIndArr1) <> mIndexedPost Then
            If .mRange(mIndArr1, .mColPost.mNum) <> mPostName Then
                mArr1(mIndArr1, 1) = ""
            End If
            'End If
        Next
        
        ' ---------------------------------
        ' fnd not null
        For mIndArr1 = 1 To UBound(mArr1, 1)
            If Len(mArr1(mIndArr1, 1)) > 0 Then mNotNull = mNotNull + 1
        Next
        ReDim mArr3(mNotNull, 1)
        mTempLong = 0
        For mIndArr1 = 1 To UBound(mArr1, 1)
            If Len(mArr1(mIndArr1, 1)) > 0 Then
                mArr3(mTempLong, 1) = mArr1(mIndArr1, 1)
                mTempLong = mTempLong + 1
            End If
        Next
        ' ---------------------------------
        
        mArr2 = f_SortArray_arrvar(mArr3)
        Erase mArr1
        mArr1 = f_SelectUniqueInArr_arrvar(mArr2)
                
    End With
    
    mFindProductGroup = mArr1
    
    Erase mArr1
    Erase mArr2
    Erase mArr3
    
End Function

Private Function mAddDataForGroup(ByVal mIndexedPost As Long, ByVal mPostName As String, _
ByVal mIndexedGroup As Long, ByVal mGroupName As String, ByVal mRowShI As Long, _
ByVal mBarText As String) As Long
        
    Dim mFirstRowGroup              As Long
    Dim mArrProducts()              As Variant
    Dim mIndProd                    As Long
    Dim mRowIFileMaterTable         As Long
    Dim mTxtBar                     As String
    
    mFirstRowGroup = mRowShI
    
    mSheetI.mRange(mFirstRowGroup, mSheetI.mColOzonSKU.mNum) = mGroupName
    mSheetI.mRange(mFirstRowGroup, mSheetI.mColOzonSKU.mNum).EntireRow.Interior.Color = mColors.mShPostGroup2
    mRowShI = mRowShI + 1
    
    mArrProducts = mFindProducts(mIndexedPost, mPostName, mIndexedGroup, mGroupName)
    
    If Not (-1 = Not mArrProducts) Then
    For mIndProd = 0 To UBound(mArrProducts)
        'If UBound(mArrProducts) > 0 Then
        
            mTxtBar = mBarText & ", заполняю товары " & _
                mIndProd + 1 & " из " & UBound(mArrProducts) + 1
            Application.StatusBar = mTxtBar
            DoEvents
            
            mSheetI.mRange(mRowShI, mSheetI.mColCaption.mNum) = mArrProducts(mIndProd)
            
            mRowIFileMaterTable = mFindRowInMasterTableByProdName(mArrProducts(mIndProd))
            'If mRowIFileMaterTable > 0 Then
                With mImportedFile.mFileMasterTable
                    
                    mSheetI.mRange(mRowShI, mSheetI.mColArt.mNum) = .mRange(mRowIFileMaterTable, .mColArt.mNum)
                    mSheetI.mRange(mRowShI, mSheetI.mColEI.mNum) = .mRange(mRowIFileMaterTable, .mColEI.mNum)
                    mSheetI.mRange(mRowShI, mSheetI.mColCostZakup.mNum) = f_dbl_ValToDbl(.mRange(mRowIFileMaterTable, .mColCostZakup.mNum))
                    mSheetI.mRange(mRowShI, mSheetI.mColOzonSKU.mNum) = .mRange(mRowIFileMaterTable, .mColOzonSKU.mNum)
                    mSheetI.mRange(mRowShI, mSheetI.mColCategory.mNum) = .mRange(mRowIFileMaterTable, .mColCategory.mNum)
                    
                    mSheetI.mRange(mRowShI, mSheetI.mColOstatokIzMT.mNum) = f_dbl_ValToDbl(.mRange(mRowIFileMaterTable, .mColOstSkladTek.mNum)) + _
                        f_dbl_ValToDbl(.mRange(mRowIFileMaterTable, .mColOstAndTVP.mNum))
                    
                End With
                mRowShI = mRowShI + 1
            'End If
        
        'End If
    Next
    End If
    
    mSheetI.mRange.Rows(mRowShI - 1 & ":" & mFirstRowGroup + 1).Group
    
    mRowShI = mRowShI + 1
    mAddDataForGroup = mRowShI
        
End Function

Private Function mFindProducts(ByVal mIndexedPost As Long, ByVal mPostName As String, _
ByVal mIndexedGroup As Long, ByVal mGroupName As String) As Variant()
    
    Dim mArr1()             As Variant
    Dim mArr2()             As Variant
    Dim mIndArr1            As Long
    
    Dim mNotNull            As Long
    Dim mArr3()             As Variant
    Dim mTempLong           As Long
    
    With mImportedFile.mFileMasterTable
        
        mArr1 = Range(.mRange(1, .mColCaption.mNum), .mRange(.mRange.Rows.Count, .mColCaption.mNum))
        
        For mIndArr1 = 1 To UBound(mArr1, 1)
            
            'If .mArrIndex(0, mIndArr1) <> mIndexedPost Then
            If .mRange(mIndArr1, .mColPost.mNum) <> mPostName Then
                mArr1(mIndArr1, 1) = ""
                'GoTo Line_NextPos
            End If
            'End If
            
            'If .mArrIndex(1, mIndArr1) <> mIndexedGroup Then
            If .mRange(mIndArr1, .mPoductType.mNum) <> mGroupName Then
                mArr1(mIndArr1, 1) = ""
            End If
            'End If
            
Line_NextPos:
        Next
        
        ' ---------------------------------
        ' fnd not null
        For mIndArr1 = 1 To UBound(mArr1, 1)
            If Len(mArr1(mIndArr1, 1)) > 0 Then
                mNotNull = mNotNull + 1
            End If
        Next
        ReDim mArr3(mNotNull, 1)
        mTempLong = 0
        For mIndArr1 = 1 To UBound(mArr1, 1)
            If Len(mArr1(mIndArr1, 1)) > 0 Then
                mArr3(mTempLong, 1) = mArr1(mIndArr1, 1)
                mTempLong = mTempLong + 1
            End If
        Next
        ' ---------------------------------
        
        mArr2 = f_SortArray_arrvar(mArr3)
        Erase mArr1
        mArr1 = f_SelectUniqueInArr_arrvar(mArr2)
        
    End With
    
    mFindProducts = mArr1
    
    Erase mArr1
    Erase mArr2
    Erase mArr3
    
End Function

Private Function mSetFormulsToShIForNoGroup()
    
    Dim mNumRow                     As Long
    Dim mTempStr_1                  As String
    
    mTempStr_1 = Split(ThisWorkbook.Sheets(mSheetI.mName).UsedRange.Address, ":")(1)
    Set mSheetI.mRange = Range(ThisWorkbook.Sheets(mSheetI.mName).Range("A1"), _
        ThisWorkbook.Sheets(mSheetI.mName).Range(mTempStr_1))
    
    For mNumRow = mSheetI.mNumRowStart To mSheetI.mRange.Rows.Count
        If mSheetI.mRange(mNumRow, 1).Interior.Color <> mColors.mShPostGroup1 Then
        If mSheetI.mRange(mNumRow, 1).Interior.Color <> mColors.mShPostGroup2 Then
            
            ' --------------------------------------------------------------------------
            
            mSheetI.mRange(mNumRow, mSheetI.mColSellsTotalMoney.mNum).Formula = "=" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColSellsMoneyPer1.mNum).Address, "$", "") & "+" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColSellsMoneyPer2.mNum).Address, "$", "")
            
            mSheetI.mRange(mNumRow, mSheetI.mColSellsTotalCount.mNum).Formula = "=" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColSellsCountPer1.mNum).Address, "$", "") & "+" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColSellsCountPer2.mNum).Address, "$", "")
            
            mSheetI.mRange(mNumRow, mSheetI.mCostProdazi.mNum).Formula = "=IFERROR(" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColSellsTotalMoney.mNum).Address, "$", "") & "/" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColSellsTotalCount.mNum).Address, "$", "") & _
                ","""")"
                
            ' --------------------------------------------------------------------------
            
            mSheetI.mRange(mNumRow, mSheetI.mColTotalMarzaRub1.mNum).Formula = "=" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColValovka1.mNum).Address, "$", "") & "-" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColStoimRazmes1.mNum).Address, "$", "") & "-" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColProdvizCost1.mNum).Address, "$", "")
            
            mSheetI.mRange(mNumRow, mSheetI.mColTotalMarzaRub2.mNum).Formula = "=" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColValovka2.mNum).Address, "$", "") & "-" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColStoimRazmes2.mNum).Address, "$", "") & "-" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColProdvizCost2.mNum).Address, "$", "")
                
            mSheetI.mRange(mNumRow, mSheetI.mColTotalMarzaRubItog.mNum).Formula = "=" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColTotalMarzaRub1.mNum).Address, "$", "") & "+" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColTotalMarzaRub2.mNum).Address, "$", "")
                
            ' --------------------------------------------------------------------------
            
            mSheetI.mRange(mNumRow, mSheetI.mColValovka1.mNum).Formula = "=" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColSellsMoneyPer1.mNum).Address, "$", "") & "-(" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColCostZakup.mNum).Address, "$", "") & "*" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColSellsCountPer1.mNum).Address, "$", "") & ")+" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColEkvaring1.mNum).Address, "$", "")
            
            mSheetI.mRange(mNumRow, mSheetI.mColValovka2.mNum).Formula = "=" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColSellsMoneyPer2.mNum).Address, "$", "") & "-(" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColCostZakup.mNum).Address, "$", "") & "*" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColSellsCountPer2.mNum).Address, "$", "") & ")+" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColEkvaring2.mNum).Address, "$", "")
            
            ' --------------------------------------------------------------------------
            
            mSheetI.mRange(mNumRow, mSheetI.mColNazenka1.mNum).Formula = "=IFERROR(" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColSellsMoneyPer1.mNum).Address, "$", "") & "/(" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColSellsMoneyPer1.mNum).Address, "$", "") & "-" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColValovka1.mNum).Address, "$", "") & ")-1" & _
                ","""")"
            
            mSheetI.mRange(mNumRow, mSheetI.mColNazenka2.mNum).Formula = "=IFERROR(" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColSellsMoneyPer2.mNum).Address, "$", "") & "/(" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColSellsMoneyPer2.mNum).Address, "$", "") & "-" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColValovka2.mNum).Address, "$", "") & ")-1" & _
                ","""")"
                
            ' --------------------------------------------------------------------------
            
            mSheetI.mRange(mNumRow, mSheetI.mNacenProcentItog.mNum).Formula = "=IFERROR(" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColSellsTotalMoney.mNum).Address, "$", "") & "/(" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColSellsTotalMoney.mNum).Address, "$", "") & "-" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColTotalMarzaRubItog.mNum).Address, "$", "") & ")-1" & _
                ","""")"
            
            ' --------------------------------------------------------------------------
            
            mSheetI.mRange(mNumRow, mSheetI.mColTotalMarzaPers1.mNum).Formula = "=IFERROR(" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColSellsMoneyPer1.mNum).Address, "$", "") & "/(" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColSellsMoneyPer1.mNum).Address, "$", "") & "-" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColTotalMarzaRub1.mNum).Address, "$", "") & ")-1" & _
                ","""")"
                
            mSheetI.mRange(mNumRow, mSheetI.mColTotalMarzaPers2.mNum).Formula = "=IFERROR(" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColSellsMoneyPer2.mNum).Address, "$", "") & "/(" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColSellsMoneyPer2.mNum).Address, "$", "") & "-" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColTotalMarzaRub2.mNum).Address, "$", "") & ")-1" & _
                ","""")"
                
            mSheetI.mRange(mNumRow, mSheetI.mColTotalMarzaPersItog.mNum).Formula = "=IFERROR(" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColSellsTotalMoney.mNum).Address, "$", "") & "/(" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColSellsTotalMoney.mNum).Address, "$", "") & "-" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColTotalMarzaRubItog.mNum).Address, "$", "") & ")-1" & _
                ","""")"
                                
            ' --------------------------------------------------------------------------
            
            mSheetI.mRange(mNumRow, mSheetI.mColOstatokTekyshRub.mNum).Formula = "=IFERROR(" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mColOstatokTekyshOZON.mNum).Address, "$", "") & "*" & _
                Replace(mSheetI.mRange(mNumRow, mSheetI.mCostProdazi.mNum).Address, "$", "") & _
                ","""")"
            
            ' --------------------------------------------------------------------------
            
        End If
        End If
    Next
        
End Function

Private Function mSummForGroups()
    
    Dim mNumRow1            As Long
    Dim mNumRow2            As Long
    Dim mFormula            As String
    Dim mArrRows()          As Long
    Dim mColName            As String
    
    For mNumRow1 = mSheetI.mNumRowStart To mSheetI.mRange.Rows.Count
        If mSheetI.mRange(mNumRow1, 1).Interior.Color = mColors.mShPostGroup1 Then
            
            If -1 = Not mArrRows Then
                ReDim mArrRows(0)
            Else
                ReDim Preserve mArrRows(UBound(mArrRows) + 1)
            End If
            mArrRows(UBound(mArrRows)) = mNumRow1
            
            mSheetI.mRange(mNumRow1, mSheetI.mColNazenka1.mNum).Formula = "=IFERROR(" & _
                Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer1.mNum).Address, "$", "") & "/(" & _
                Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer1.mNum).Address, "$", "") & "-" & _
                Replace(mSheetI.mRange(mNumRow1, mSheetI.mColValovka1.mNum).Address, "$", "") & ")-1" & _
                ","""")"
            
            mSheetI.mRange(mNumRow1, mSheetI.mColNazenka2.mNum).Formula = "=IFERROR(" & _
                Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer2.mNum).Address, "$", "") & "/(" & _
                Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer2.mNum).Address, "$", "") & "-" & _
                Replace(mSheetI.mRange(mNumRow1, mSheetI.mColValovka2.mNum).Address, "$", "") & ")-1" & _
                ","""")"
                
            mSheetI.mRange(mNumRow1, mSheetI.mNacenProcentItog.mNum).Formula = "=IFERROR(" & _
                Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsTotalMoney.mNum).Address, "$", "") & "/(" & _
                Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsTotalMoney.mNum).Address, "$", "") & "-" & _
                Replace(mSheetI.mRange(mNumRow1, mSheetI.mColTotalMarzaRubItog.mNum).Address, "$", "") & ")-1" & _
                ","""")"
            
            mSheetI.mRange(mNumRow1, mSheetI.mColTotalMarzaPers1.mNum).Formula = "=IFERROR(" & _
                Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer1.mNum).Address, "$", "") & "/(" & _
                Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer1.mNum).Address, "$", "") & "-" & _
                Replace(mSheetI.mRange(mNumRow1, mSheetI.mColTotalMarzaRub1.mNum).Address, "$", "") & ")-1" & _
                ","""")"
                
            mSheetI.mRange(mNumRow1, mSheetI.mColTotalMarzaPers2.mNum).Formula = "=IFERROR(" & _
                Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer2.mNum).Address, "$", "") & "/(" & _
                Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer2.mNum).Address, "$", "") & "-" & _
                Replace(mSheetI.mRange(mNumRow1, mSheetI.mColTotalMarzaRub2.mNum).Address, "$", "") & ")-1" & _
                ","""")"
                
            mSheetI.mRange(mNumRow1, mSheetI.mColTotalMarzaPersItog.mNum).Formula = "=IFERROR(" & _
                Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsTotalMoney.mNum).Address, "$", "") & "/(" & _
                Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsTotalMoney.mNum).Address, "$", "") & "-" & _
                Replace(mSheetI.mRange(mNumRow1, mSheetI.mColTotalMarzaRubItog.mNum).Address, "$", "") & ")-1" & _
                ","""")"
                
        End If
    Next
    
    mFrmlGen False, mArrRows, mSheetI.mNumRowStart - 1, mSheetI.mColSellsMoneyPer1.mNum
    mFrmlGen False, mArrRows, mSheetI.mNumRowStart - 1, mSheetI.mColSellsMoneyPer2.mNum
    mFrmlGen False, mArrRows, mSheetI.mNumRowStart - 1, mSheetI.mColSellsTotalMoney.mNum
    
    mFrmlGen False, mArrRows, mSheetI.mNumRowStart - 1, mSheetI.mColEkvaring1.mNum
    mFrmlGen False, mArrRows, mSheetI.mNumRowStart - 1, mSheetI.mColEkvaring2.mNum
    mFrmlGen False, mArrRows, mSheetI.mNumRowStart - 1, mSheetI.mColValovka1.mNum
    mFrmlGen False, mArrRows, mSheetI.mNumRowStart - 1, mSheetI.mColValovka2.mNum
    mFrmlGen False, mArrRows, mSheetI.mNumRowStart - 1, mSheetI.mColProdvizCost1.mNum
    mFrmlGen False, mArrRows, mSheetI.mNumRowStart - 1, mSheetI.mColProdvizCost2.mNum
    mFrmlGen False, mArrRows, mSheetI.mNumRowStart - 1, mSheetI.mColStoimRazmes1.mNum
    mFrmlGen False, mArrRows, mSheetI.mNumRowStart - 1, mSheetI.mColStoimRazmes2.mNum
    
    mFrmlGen False, mArrRows, mSheetI.mNumRowStart - 1, mSheetI.mColTotalMarzaRub1.mNum
    mFrmlGen False, mArrRows, mSheetI.mNumRowStart - 1, mSheetI.mColTotalMarzaRub2.mNum
    mFrmlGen False, mArrRows, mSheetI.mNumRowStart - 1, mSheetI.mColTotalMarzaRubItog.mNum
    
    mFrmlGen False, mArrRows, mSheetI.mNumRowStart - 1, mSheetI.mColOstatokTekyshRub.mNum
    
    
    mSheetI.mRange(mSheetI.mNumRowStart - 1, mSheetI.mColNazenka1.mNum).Formula = "=IFERROR(" & _
        Replace(mSheetI.mRange(mSheetI.mNumRowStart - 1, mSheetI.mColSellsMoneyPer1.mNum).Address, "$", "") & "/(" & _
        Replace(mSheetI.mRange(mSheetI.mNumRowStart - 1, mSheetI.mColSellsMoneyPer1.mNum).Address, "$", "") & "-" & _
        Replace(mSheetI.mRange(mSheetI.mNumRowStart - 1, mSheetI.mColValovka1.mNum).Address, "$", "") & ")-1" & _
        ","""")"
    
    mSheetI.mRange(mSheetI.mNumRowStart - 1, mSheetI.mColNazenka2.mNum).Formula = "=IFERROR(" & _
        Replace(mSheetI.mRange(mSheetI.mNumRowStart - 1, mSheetI.mColSellsMoneyPer2.mNum).Address, "$", "") & "/(" & _
        Replace(mSheetI.mRange(mSheetI.mNumRowStart - 1, mSheetI.mColSellsMoneyPer2.mNum).Address, "$", "") & "-" & _
        Replace(mSheetI.mRange(mSheetI.mNumRowStart - 1, mSheetI.mColValovka2.mNum).Address, "$", "") & ")-1" & _
        ","""")"
    
    mSheetI.mRange(mSheetI.mNumRowStart - 1, mSheetI.mColTotalMarzaPers1.mNum).Formula = "=IFERROR(" & _
        Replace(mSheetI.mRange(mSheetI.mNumRowStart - 1, mSheetI.mColSellsMoneyPer1.mNum).Address, "$", "") & "/(" & _
        Replace(mSheetI.mRange(mSheetI.mNumRowStart - 1, mSheetI.mColSellsMoneyPer1.mNum).Address, "$", "") & "-" & _
        Replace(mSheetI.mRange(mSheetI.mNumRowStart - 1, mSheetI.mColTotalMarzaRub1.mNum).Address, "$", "") & ")-1" & _
        ","""")"
        
    mSheetI.mRange(mSheetI.mNumRowStart - 1, mSheetI.mColTotalMarzaPers2.mNum).Formula = "=IFERROR(" & _
        Replace(mSheetI.mRange(mSheetI.mNumRowStart - 1, mSheetI.mColSellsMoneyPer2.mNum).Address, "$", "") & "/(" & _
        Replace(mSheetI.mRange(mSheetI.mNumRowStart - 1, mSheetI.mColSellsMoneyPer2.mNum).Address, "$", "") & "-" & _
        Replace(mSheetI.mRange(mSheetI.mNumRowStart - 1, mSheetI.mColTotalMarzaRub2.mNum).Address, "$", "") & ")-1" & _
        ","""")"
        
    mSheetI.mRange(mSheetI.mNumRowStart - 1, mSheetI.mColTotalMarzaPersItog.mNum).Formula = "=IFERROR(" & _
        Replace(mSheetI.mRange(mSheetI.mNumRowStart - 1, mSheetI.mColSellsTotalMoney.mNum).Address, "$", "") & "/(" & _
        Replace(mSheetI.mRange(mSheetI.mNumRowStart - 1, mSheetI.mColSellsTotalMoney.mNum).Address, "$", "") & "-" & _
        Replace(mSheetI.mRange(mSheetI.mNumRowStart - 1, mSheetI.mColTotalMarzaRubItog.mNum).Address, "$", "") & ")-1" & _
        ","""")"
    
    ' ------------------------------------------------------------------------------------------------------
    
    Erase mArrRows
    For mNumRow1 = mSheetI.mNumRowStart To mSheetI.mRange.Rows.Count
        If mSheetI.mRange(mNumRow1, 1).Interior.Color = mColors.mShPostGroup1 Then
                        
            Erase mArrRows
            For mNumRow2 = mNumRow1 + 1 To mSheetI.mRange.Rows.Count
                If mSheetI.mRange(mNumRow2, 1).Interior.Color = mColors.mShPostGroup1 Then Exit For
                If mSheetI.mRange(mNumRow2, 1).Interior.Color = mColors.mShPostGroup2 Then
                    
                    If -1 = Not mArrRows Then
                        ReDim mArrRows(0)
                    Else
                        ReDim Preserve mArrRows(UBound(mArrRows) + 1)
                    End If
                    mArrRows(UBound(mArrRows)) = mNumRow2
                    
                    mSheetI.mRange(mNumRow1, mSheetI.mColNazenka1.mNum).Formula = "=IFERROR(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer1.mNum).Address, "$", "") & "/(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer1.mNum).Address, "$", "") & "-" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColValovka1.mNum).Address, "$", "") & ")-1" & _
                        ","""")"
                    
                    mSheetI.mRange(mNumRow1, mSheetI.mColNazenka2.mNum).Formula = "=IFERROR(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer2.mNum).Address, "$", "") & "/(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer2.mNum).Address, "$", "") & "-" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColValovka2.mNum).Address, "$", "") & ")-1" & _
                        ","""")"
                        
                    mSheetI.mRange(mNumRow1, mSheetI.mNacenProcentItog.mNum).Formula = "=IFERROR(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsTotalMoney.mNum).Address, "$", "") & "/(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsTotalMoney.mNum).Address, "$", "") & "-" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColTotalMarzaRubItog.mNum).Address, "$", "") & ")-1" & _
                        ","""")"
                    
                    mSheetI.mRange(mNumRow1, mSheetI.mColTotalMarzaPers1.mNum).Formula = "=IFERROR(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer1.mNum).Address, "$", "") & "/(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer1.mNum).Address, "$", "") & "-" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColTotalMarzaRub1.mNum).Address, "$", "") & ")-1" & _
                        ","""")"
                        
                    mSheetI.mRange(mNumRow1, mSheetI.mColTotalMarzaPers2.mNum).Formula = "=IFERROR(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer2.mNum).Address, "$", "") & "/(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer2.mNum).Address, "$", "") & "-" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColTotalMarzaRub2.mNum).Address, "$", "") & ")-1" & _
                        ","""")"
                        
                    mSheetI.mRange(mNumRow1, mSheetI.mColTotalMarzaPersItog.mNum).Formula = "=IFERROR(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsTotalMoney.mNum).Address, "$", "") & "/(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsTotalMoney.mNum).Address, "$", "") & "-" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColTotalMarzaRubItog.mNum).Address, "$", "") & ")-1" & _
                        ","""")"
                
                End If
            Next
            If Not (-1 = Not mArrRows) Then
                
                mFrmlGen False, mArrRows, mNumRow1, mSheetI.mColSellsMoneyPer1.mNum
                mFrmlGen False, mArrRows, mNumRow1, mSheetI.mColSellsMoneyPer2.mNum
                mFrmlGen False, mArrRows, mNumRow1, mSheetI.mColSellsTotalMoney.mNum
                
                mFrmlGen False, mArrRows, mNumRow1, mSheetI.mColEkvaring1.mNum
                mFrmlGen False, mArrRows, mNumRow1, mSheetI.mColEkvaring2.mNum
                mFrmlGen False, mArrRows, mNumRow1, mSheetI.mColValovka1.mNum
                mFrmlGen False, mArrRows, mNumRow1, mSheetI.mColValovka2.mNum
                mFrmlGen False, mArrRows, mNumRow1, mSheetI.mColProdvizCost1.mNum
                mFrmlGen False, mArrRows, mNumRow1, mSheetI.mColProdvizCost2.mNum
                mFrmlGen False, mArrRows, mNumRow1, mSheetI.mColStoimRazmes1.mNum
                mFrmlGen False, mArrRows, mNumRow1, mSheetI.mColStoimRazmes2.mNum
                
                mFrmlGen False, mArrRows, mNumRow1, mSheetI.mColTotalMarzaRub1.mNum
                mFrmlGen False, mArrRows, mNumRow1, mSheetI.mColTotalMarzaRub2.mNum
                mFrmlGen False, mArrRows, mNumRow1, mSheetI.mColTotalMarzaRubItog.mNum
                
                mFrmlGen False, mArrRows, mNumRow1, mSheetI.mColOstatokTekyshRub.mNum
                
            End If
            
        End If
    Next
    
    ' ------------------------------------------------------------------------------------------------------
    
    Erase mArrRows
    For mNumRow1 = mSheetI.mNumRowStart To mSheetI.mRange.Rows.Count
        If mSheetI.mRange(mNumRow1, 1).Interior.Color = mColors.mShPostGroup2 Then
                        
            Erase mArrRows
            For mNumRow2 = mNumRow1 + 1 To mSheetI.mRange.Rows.Count
                If mSheetI.mRange(mNumRow2, 1).Interior.Color = mColors.mShPostGroup1 Then Exit For
                If mSheetI.mRange(mNumRow2, 1).Interior.Color = mColors.mShPostGroup2 Then Exit For
                'If mSheetI.mRange(mNumRow2, 1).Interior.Color = mColors.mShPostGroup2 Then
                    If -1 = Not mArrRows Then
                        ReDim mArrRows(0)
                    Else
                        ReDim Preserve mArrRows(UBound(mArrRows) + 1)
                    End If
                    mArrRows(UBound(mArrRows)) = mNumRow2
                    
                    mSheetI.mRange(mNumRow1, mSheetI.mColNazenka1.mNum).Formula = "=IFERROR(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer1.mNum).Address, "$", "") & "/(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer1.mNum).Address, "$", "") & "-" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColValovka1.mNum).Address, "$", "") & ")-1" & _
                        ","""")"
                    
                    mSheetI.mRange(mNumRow1, mSheetI.mColNazenka2.mNum).Formula = "=IFERROR(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer2.mNum).Address, "$", "") & "/(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer2.mNum).Address, "$", "") & "-" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColValovka2.mNum).Address, "$", "") & ")-1" & _
                        ","""")"
                        
                    mSheetI.mRange(mNumRow1, mSheetI.mNacenProcentItog.mNum).Formula = "=IFERROR(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsTotalMoney.mNum).Address, "$", "") & "/(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsTotalMoney.mNum).Address, "$", "") & "-" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColTotalMarzaRubItog.mNum).Address, "$", "") & ")-1" & _
                        ","""")"
                    
                    mSheetI.mRange(mNumRow1, mSheetI.mColTotalMarzaPers1.mNum).Formula = "=IFERROR(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer1.mNum).Address, "$", "") & "/(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer1.mNum).Address, "$", "") & "-" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColTotalMarzaRub1.mNum).Address, "$", "") & ")-1" & _
                        ","""")"
                        
                    mSheetI.mRange(mNumRow1, mSheetI.mColTotalMarzaPers2.mNum).Formula = "=IFERROR(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer2.mNum).Address, "$", "") & "/(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsMoneyPer2.mNum).Address, "$", "") & "-" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColTotalMarzaRub2.mNum).Address, "$", "") & ")-1" & _
                        ","""")"
                        
                    mSheetI.mRange(mNumRow1, mSheetI.mColTotalMarzaPersItog.mNum).Formula = "=IFERROR(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsTotalMoney.mNum).Address, "$", "") & "/(" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColSellsTotalMoney.mNum).Address, "$", "") & "-" & _
                        Replace(mSheetI.mRange(mNumRow1, mSheetI.mColTotalMarzaRubItog.mNum).Address, "$", "") & ")-1" & _
                        ","""")"
                
                'End If
            Next
            If Not (-1 = Not mArrRows) Then
                
                mFrmlGen True, mArrRows, mNumRow1, mSheetI.mColSellsMoneyPer1.mNum
                mFrmlGen True, mArrRows, mNumRow1, mSheetI.mColSellsMoneyPer2.mNum
                mFrmlGen True, mArrRows, mNumRow1, mSheetI.mColSellsTotalMoney.mNum
                
                mFrmlGen True, mArrRows, mNumRow1, mSheetI.mColEkvaring1.mNum
                mFrmlGen True, mArrRows, mNumRow1, mSheetI.mColEkvaring2.mNum
                mFrmlGen True, mArrRows, mNumRow1, mSheetI.mColValovka1.mNum
                mFrmlGen True, mArrRows, mNumRow1, mSheetI.mColValovka2.mNum
                mFrmlGen True, mArrRows, mNumRow1, mSheetI.mColProdvizCost1.mNum
                mFrmlGen True, mArrRows, mNumRow1, mSheetI.mColProdvizCost2.mNum
                mFrmlGen True, mArrRows, mNumRow1, mSheetI.mColStoimRazmes1.mNum
                mFrmlGen True, mArrRows, mNumRow1, mSheetI.mColStoimRazmes2.mNum
                
                mFrmlGen True, mArrRows, mNumRow1, mSheetI.mColTotalMarzaRub1.mNum
                mFrmlGen True, mArrRows, mNumRow1, mSheetI.mColTotalMarzaRub2.mNum
                mFrmlGen True, mArrRows, mNumRow1, mSheetI.mColTotalMarzaRubItog.mNum
                
                mFrmlGen False, mArrRows, mNumRow1, mSheetI.mColOstatokTekyshRub.mNum
                
            End If
            
        End If
    Next
    
End Function

Private Function mFrmlGen(ByVal mLastEndRow As Boolean, ByRef mArrRows() As Long, _
ByVal mRowToFrml As Long, ByVal mNumCol As Long)
    
    Dim mIndArr         As Long
    Dim mFormula        As String
    
    If mLastEndRow = True Then
        mFormula = Replace(mSheetI.mRange(mArrRows(0), mNumCol).Address, "$", "") & ":" & _
            Replace(mSheetI.mRange(mArrRows(UBound(mArrRows)), mNumCol).Address, "$", "")
    Else
        For mIndArr = 0 To UBound(mArrRows)
            mFormula = mFormula & "," & _
                Replace(mSheetI.mRange(mArrRows(mIndArr), mNumCol).Address, "$", "")
        Next
        mFormula = Mid(mFormula, 2)
    End If
        
    mFormula = "=SUM(" & mFormula & ")"
    mSheetI.mRange(mRowToFrml, mNumCol).Formula = mFormula
    
End Function

Private Function mSetFormatsAndGrd()
    
    Dim mTempStr_1                  As String
    
    mTempStr_1 = Split(ThisWorkbook.Sheets(mSheetI.mName).UsedRange.Address, ":")(1)
    Set mSheetI.mRange = Range(ThisWorkbook.Sheets(mSheetI.mName).Cells(mSheetI.mNumRowStart, 1), _
        ThisWorkbook.Sheets(mSheetI.mName).Range(mTempStr_1))
    
    mSetGrd mSheetI.mRange, 1, 1
    mSetGrd mSheetI.mRange, 1, 2
    
    mSetGrd Range(ThisWorkbook.Sheets(mSheetI.mName).Cells(mSheetI.mNumRowTitle, 1), _
        ThisWorkbook.Sheets(mSheetI.mName).Cells(mSheetI.mNumRowStart - 1, mSheetI.mRange.Columns.Count)), 2, 1
    
    mSheetI.mRange.Columns(mSheetI.mColCostZakup.mNum).NumberFormat = "0.00"
    mSheetI.mRange.Columns(mSheetI.mColNazenka1.mNum).NumberFormat = "0.0%"
    mSheetI.mRange.Columns(mSheetI.mColProdvizCost1.mNum).NumberFormat = "0.00"
    mSheetI.mRange.Columns(mSheetI.mColSellsMoneyPer1.mNum).NumberFormat = "0.00"
    mSheetI.mRange.Columns(mSheetI.mColSellsMoneyPer2.mNum).NumberFormat = "0.00"
    mSheetI.mRange.Columns(mSheetI.mColSellsTotalMoney.mNum).NumberFormat = "0.00"
    mSheetI.mRange.Columns(mSheetI.mColStoimRazmes1.mNum).NumberFormat = "0.00"
    mSheetI.mRange.Columns(mSheetI.mColTotalMarzaRub1.mNum).NumberFormat = "0.00"
    mSheetI.mRange.Columns(mSheetI.mColValovka1.mNum).NumberFormat = "0.00"
    mSheetI.mRange.Columns(mSheetI.mCostProdazi.mNum).NumberFormat = "0.00"
    mSheetI.mRange.Columns(mSheetI.mColEkvaring1.mNum).NumberFormat = "0.00"
    mSheetI.mRange.Columns(mSheetI.mNacenProcentItog.mNum).NumberFormat = "0.0%"
    mSheetI.mRange.Columns(mSheetI.mColNazenka2.mNum).NumberFormat = "0.0%"
    mSheetI.mRange.Columns(mSheetI.mColValovka2.mNum).NumberFormat = "0.00"
    mSheetI.mRange.Columns(mSheetI.mColProdvizCost2.mNum).NumberFormat = "0.00"
    mSheetI.mRange.Columns(mSheetI.mColStoimRazmes2.mNum).NumberFormat = "0.00"
    mSheetI.mRange.Columns(mSheetI.mColTotalMarzaRub2.mNum).NumberFormat = "0.00"
    mSheetI.mRange.Columns(mSheetI.mColTotalMarzaRubItog.mNum).NumberFormat = "0.00"
    mSheetI.mRange.Columns(mSheetI.mColEkvaring2.mNum).NumberFormat = "0.00"
    mSheetI.mRange.Columns(mSheetI.mColTotalMarzaPers1.mNum).NumberFormat = "0.0%"
    mSheetI.mRange.Columns(mSheetI.mColTotalMarzaPers2.mNum).NumberFormat = "0.0%"
    mSheetI.mRange.Columns(mSheetI.mColTotalMarzaPersItog.mNum).NumberFormat = "0.0%"
    mSheetI.mRange.Columns(mSheetI.mColOstatokTekyshOZON.mNum).NumberFormat = "0"
    mSheetI.mRange.Columns(mSheetI.mColOstatokIzMT.mNum).NumberFormat = "0"
    mSheetI.mRange.Columns(mSheetI.mColOstatokTekyshRub.mNum).NumberFormat = "0.00"
    
    mSheetI.mRange.Columns(mSheetI.mColSellsCountPer1.mNum).NumberFormat = "0"
    mSheetI.mRange.Columns(mSheetI.mColSellsCountPer2.mNum).NumberFormat = "0"
    mSheetI.mRange.Columns(mSheetI.mColSellsTotalCount.mNum).NumberFormat = "0"
    
    
End Function


