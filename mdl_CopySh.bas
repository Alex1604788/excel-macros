Attribute VB_Name = "mdl_CopySh"

Public Sub mCopyShProd()
    
    mFlagJob = True
    SetActivityExcel False
    Settings
    
    '----------------------------------------------------------------------------------------
    
    Dim mNameShTemp             As String
    Dim mNameShNew              As String
    Dim mRange                  As Range
    Dim mNumRow                 As Long
    Dim mPostNow                As String
    Dim mNomenklTypeNow         As String
    
    '----------------------------------------------------------------------------------------
    
    mNameShTemp = Replace(Replace(Replace(mNow, " ", ""), ".", ""), "-", "")
    mNameShNew = "_" & mSheetI.mName
    mDeleteSheet ThisWorkbook, mNameShNew
    ThisWorkbook.Sheets(mSheetI.mName).Copy _
        After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = mNameShTemp
        
    '----------------------------------------------------------------------------------------
    
    For Each mShape In ThisWorkbook.Sheets(mNameShTemp).Shapes
        mShape.Delete
    Next
    ThisWorkbook.Sheets(mNameShTemp).Columns("D:D").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ThisWorkbook.Sheets(mNameShTemp).Columns("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    mUngroupMe mNameShTemp
    Set mRange = ThisWorkbook.Sheets(mNameShTemp).UsedRange
    
    '----------------------------------------------------------------------------------------
    
    mRange(5, 3) = "Основной поставщик"
    Range(mRange(5, 3), mRange(6, 3)).Merge
    
    mRange(5, 5) = "Вид номенклатуры"
    Range(mRange(5, 5), mRange(6, 5)).Merge
    
    For mNumRow = mSheetI.mNumRowStart To mRange.Rows.Count
        
        If mRange(mNumRow, 1).Interior.Color = mColors.mShPostGroup1 Then
            mPostNow = mRange(mNumRow, 1)
        End If
        If mRange(mNumRow, 1).Interior.Color = mColors.mShPostGroup2 Then
            mNomenklTypeNow = mRange(mNumRow, 1)
        End If
        
        mRange(mNumRow, 3) = mPostNow
        mRange(mNumRow, 5) = mNomenklTypeNow
        
    Next
    
    For mNumRow = mRange.Rows.Count To mSheetI.mNumRowStart Step -1
        If mRange(mNumRow, 1).Interior.Color = mColors.mShPostGroup1 Or _
        mRange(mNumRow, 1).Interior.Color = mColors.mShPostGroup2 Then
            mRange.Rows(mNumRow).EntireRow.Interior.Color = xlNone
            mRange.Rows(mNumRow).Delete Shift:=xlUp
        End If
    Next
    
    '----------------------------------------------------------------------------------------
    
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = 0
    
    mRange.Rows(7).Delete Shift:=xlUp
    mRange.Rows("1:4").Delete Shift:=xlUp
    
    ActiveWindow.SplitRow = 2
    ActiveWindow.FreezePanes = True
    
    '----------------------------------------------------------------------------------------
    
    Set mRange = Nothing
    ThisWorkbook.Sheets(mNameShTemp).Name = mNameShNew
    ThisWorkbook.Sheets(mSheetI.mName).Select
    
    '----------------------------------------------------------------------------------------
    
Line_Exit:
    ResetVars
    StopMySub False, True
    
End Sub

Private Sub mUngroupMe(ByVal mShName As String)
    
    On Error Resume Next
    
    For mInd = 1 To 9
        ThisWorkbook.Sheets(mShName).Rows.Ungroup
    Next
    
    ThisWorkbook.Sheets(mShName).Cells.EntireRow.Hidden = False
        
    Err.Clear
    On Error GoTo 0
    
End Sub
