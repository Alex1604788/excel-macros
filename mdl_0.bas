Attribute VB_Name = "mdl_0"

Public Type mFindedSheetInBook_
    mSheetName              As String
    mNumRowTitle            As Long
End Type

Public Type mColumn_
    mNum                As Double
    mCaption            As String
    mWidthToSet         As Double
End Type

Public Type mShI_
    
    mName                   As String
    mRange                  As Range
    mNumRowTitle            As Long
    mNumRowStart            As Long
    
    mColOzonSKU             As mColumn_
    mColArt                 As mColumn_
    mColCaption             As mColumn_
    mColEI                  As mColumn_
    mColCategory            As mColumn_
    mColSellsMoneyPer1      As mColumn_
    mColSellsMoneyPer2      As mColumn_
    mColSellsCountPer1      As mColumn_
    mColSellsCountPer2      As mColumn_
    mColSellsTotalMoney     As mColumn_
    mColSellsTotalCount     As mColumn_
    mColNazenka1            As mColumn_
    mColValovka1            As mColumn_
    mColProdvizCost1        As mColumn_
    mColStoimRazmes1        As mColumn_
    mColTotalMarzaRub1      As mColumn_
    mColNazenka2            As mColumn_
    mColValovka2            As mColumn_
    mColProdvizCost2        As mColumn_
    mColStoimRazmes2        As mColumn_
    mColTotalMarzaRub2      As mColumn_
    mColTotalMarzaRubItog   As mColumn_
    mColEkvaring2           As mColumn_
    mColTotalMarzaPers1     As mColumn_
    mColTotalMarzaPers2     As mColumn_
    mColTotalMarzaPersItog  As mColumn_
    mColCostZakup           As mColumn_
    mCostProdazi            As mColumn_
    mColEkvaring1           As mColumn_
    mNacenProcentItog       As mColumn_
    mColOstatokTekyshOZON   As mColumn_
    mColOstatokTekyshRub    As mColumn_
    mColOstatokIzMT         As mColumn_
    
    mTipNachislForEkvarung     As String
    mTipNachislForNachislen As String
    mTipNachislForCount     As String
    
    
    mArrIndex()             As Long
    
End Type

Public Type mColors_
    
    mRed                    As Long
    mLightBlue              As Long
    mGreen                  As Long
    mOrange                 As Long
    mGray                   As Long
    mViolet                 As Long
    mYellow                 As Long
    mLightGreen             As Long
    mLightRed               As Long
    mDarkGreen              As Long
    
    mShPostGroup1           As Long
    mShPostGroup2           As Long
    
End Type

Public mSheetI              As mShI_
Public mColors              As mColors_
Public mSettingsGood        As Boolean
Public mFlagJob             As Boolean

Public Sub Settings()
    
    Dim mTempStr_1          As String
    Dim mFlagJobBefore      As Boolean
    
    With mSheetI
        
        .mName = "Продажи"
        mTempStr_1 = Split(ThisWorkbook.Sheets(.mName).UsedRange.Address, ":")(1)
        Set .mRange = Range(ThisWorkbook.Sheets(.mName).Range("$A$1"), _
            ThisWorkbook.Sheets(.mName).Range(mTempStr_1))
        
        .mNumRowTitle = f_str_FindParam("A0029")
        .mNumRowStart = f_str_FindParam("A0002")
                
        mNameOrNumCol .mColOzonSKU, f_str_FindParam("A0003")
        mNameOrNumCol .mColArt, f_str_FindParam("A0004")
        mNameOrNumCol .mColCaption, f_str_FindParam("A0005")
        mNameOrNumCol .mColEI, f_str_FindParam("A0006")
        mNameOrNumCol .mColCategory, f_str_FindParam("A0007")
        mNameOrNumCol .mColSellsMoneyPer1, f_str_FindParam("A0008")
        mNameOrNumCol .mColSellsMoneyPer2, f_str_FindParam("A0009")
        mNameOrNumCol .mColSellsCountPer1, f_str_FindParam("A0010")
        mNameOrNumCol .mColSellsCountPer2, f_str_FindParam("A0011")
        mNameOrNumCol .mColSellsTotalMoney, f_str_FindParam("A0012")
        mNameOrNumCol .mColSellsTotalCount, f_str_FindParam("A0013")
        mNameOrNumCol .mColNazenka1, f_str_FindParam("A0014")
        mNameOrNumCol .mColValovka1, f_str_FindParam("A0015")
        mNameOrNumCol .mColProdvizCost1, f_str_FindParam("A0016")
        mNameOrNumCol .mColStoimRazmes1, f_str_FindParam("A0017")
        mNameOrNumCol .mColTotalMarzaRub1, f_str_FindParam("A0018")
        mNameOrNumCol .mColCostZakup, f_str_FindParam("A0019")
        mNameOrNumCol .mCostProdazi, f_str_FindParam("A0030")
        mNameOrNumCol .mColEkvaring1, f_str_FindParam("A0031")
        mNameOrNumCol .mNacenProcentItog, f_str_FindParam("A0040")
        mNameOrNumCol .mColNazenka2, f_str_FindParam("A0064")
        mNameOrNumCol .mColValovka2, f_str_FindParam("A0060")
        mNameOrNumCol .mColProdvizCost2, f_str_FindParam("A0061")
        mNameOrNumCol .mColStoimRazmes2, f_str_FindParam("A0062")
        mNameOrNumCol .mColTotalMarzaRub2, f_str_FindParam("A0065")
        mNameOrNumCol .mColTotalMarzaRubItog, f_str_FindParam("A0066")
        mNameOrNumCol .mColEkvaring2, f_str_FindParam("A0063")
        mNameOrNumCol .mColTotalMarzaPers1, f_str_FindParam("A0067")
        mNameOrNumCol .mColTotalMarzaPers2, f_str_FindParam("A0068")
        mNameOrNumCol .mColTotalMarzaPersItog, f_str_FindParam("A0069")
        mNameOrNumCol .mColOstatokTekyshOZON, f_str_FindParam("A0050")
        mNameOrNumCol .mColOstatokTekyshRub, f_str_FindParam("A0059")
        mNameOrNumCol .mColOstatokIzMT, f_str_FindParam("A0070")
        
        .mTipNachislForEkvarung = f_str_FindParam("A0047")
        .mTipNachislForNachislen = f_str_FindParam("A0048")
        .mTipNachislForCount = f_str_FindParam("A0049")
        
        '----------------------------------------------------------------------------------------
        
        mFlagJobBefore = mFlagJob
        mFlagJob = True
        
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColOzonSKU
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColArt
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColCaption
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColEI
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColCategory
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColSellsMoneyPer1
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColSellsMoneyPer2
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColSellsCountPer1
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColSellsCountPer2
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColSellsTotalMoney
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColSellsTotalCount
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColNazenka1
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColValovka1
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColProdvizCost1
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColStoimRazmes1
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColTotalMarzaRub1
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColCostZakup
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mCostProdazi
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColEkvaring1
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mNacenProcentItog
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColNazenka2
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColValovka2
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColProdvizCost2
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColStoimRazmes2
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColTotalMarzaRub2
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColTotalMarzaRubItog
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColEkvaring2
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColTotalMarzaPers1
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColTotalMarzaPers2
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColTotalMarzaPersItog
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColOstatokTekyshOZON
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColOstatokTekyshRub
        mCheckCol ThisWorkbook.Name, .mRange, .mNumRowTitle, .mName, True, False, .mColOstatokIzMT
        
        mFlagJob = mFlagJobBefore
        
    End With
       
    With mColors
        .mRed = RGB(255, 150, 130)
        .mLightBlue = 15773696
        .mGreen = 5296274
        .mOrange = 49407
        .mGray = 12566463
        .mViolet = 15420361
        .mYellow = vbYellow
        .mLightGreen = 6750156
        .mLightRed = 10079487
        .mDarkGreen = 5287936
        .mShPostGroup1 = f_str_FindParam("A0032").Interior.Color
        .mShPostGroup2 = f_str_FindParam("A0033").Interior.Color
    End With
    
    mSettingsImportedFile
    
    mSettingsGood = True
    
End Sub

Public Sub ResetVars()
    
    Set mSheetI.mRange = Nothing
    
    Erase mSheetI.mArrIndex
    
    With mImportedFile
        mResetVarsImportedFile .mFileMasterTable.mRange, .mFileMasterTable.mName
        mResetVarsImportedFile .mFileStoimRazmesh.mRange, .mFileStoimRazmesh.mName
        mResetVarsImportedFile .mFileZatratProdviz.mRange, .mFileZatratProdviz.mName
    End With
    
    mSettingsGood = False
    
End Sub

Public Sub ClearSheetIam()
    
    On Error Resume Next
    
    If mSettingsGood = False Then Settings
    
    Dim mNumRowEnd          As Long
    Dim mTempStr_1          As String
    Dim mTempLong1          As Long
    Dim mTempLong2          As Long
    
    mNumRowEnd = mSheetI.mRange.Rows.Count
    mTempLong2 = 0
    
    ThisWorkbook.Sheets(mSheetI.mName).Select
    mTempLong2 = ActiveWindow.SplitRow
    ActiveWindow.FreezePanes = False
    
    For mInd = 1 To 9
        ThisWorkbook.Sheets(mSheetI.mName).Rows.Ungroup
    Next
    
    If mNumRowEnd >= mSheetI.mNumRowStart Then
        ThisWorkbook.Sheets(mSheetI.mName).Cells.EntireRow.Hidden = False
        With ThisWorkbook.Sheets(mSheetI.mName).Range("A" & mSheetI.mNumRowStart & ":A" & mNumRowEnd + 1)
            .EntireRow.Delete Shift:=xlUp
        End With
    End If

    ActiveWindow.SplitRow = mTempLong2
    ActiveWindow.FreezePanes = True
    
    Err.Clear
    On Error GoTo 0
    
End Sub

Public Sub StopMySub(ByVal mWithoutMsg As Boolean, ByVal mNeedEnd As Boolean)
    
    ResetVars
    mFlagJob = False
    SetActivityExcel True
    If mWithoutMsg = False Then MsgBox "готово", vbInformation
    If mNeedEnd = True Then End
    
End Sub
