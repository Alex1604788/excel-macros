Attribute VB_Name = "mdl_Index"

Public Function CreateIndex(ByVal mData As String) As Long
    
    mData = Trim(mData)
    
    If Len(mData) = 0 Then
        CreateIndex = 0
    Else
        CreateIndex = ((Len(mData) * Asc(Mid(mData, 1, 1))) * Asc(Mid(mData, Len(mData), 1))) + Asc(Mid(mData, 1, 1)) + Asc(Mid(mData, Len(mData), 1))
        'For mPos = 1 To Len(mData)
        '    CreateIndex = CreateIndex + Asc(Mid(mData, mPos, 1))
        'Next
    End If
    
End Function

Public Sub mCreateArrIndexAll(ByVal mSheetINumColForIndex As Long, _
ByVal mNeedIndex_ThisFileShProdazi As Boolean, _
ByVal mNeedIndex_FileMasterTable As Boolean, _
ByVal mNeedIndex_FileStoimRazmesh As Boolean, _
ByVal mNeedIndex_FileZatratProdviz As Boolean, _
ByVal mNeedIndex_FileOtchetPopr As Boolean)

    Dim NumRow          As Long
    Dim mArrayCols()    As Variant
    
    If mNeedIndex_ThisFileShProdazi = True Then
        With mSheetI
            ReDim .mArrIndex(.mRange.Rows.Count)
            For NumRow = .mNumRowStart To .mRange.Rows.Count
                .mArrIndex(NumRow) = CreateIndex(.mRange(NumRow, mSheetINumColForIndex))
            Next
        End With
    End If
    
    With mImportedFile
        
        If mNeedIndex_FileMasterTable = True Then
            Erase mArrayCols
            mArrayCols = Array(.mFileMasterTable.mColPost.mNum, .mFileMasterTable.mPoductType.mNum, _
                .mFileMasterTable.mColCaption.mNum, .mFileMasterTable.mColArt.mNum)
            mCreateArrIndexImportedFile .mFileMasterTable.mName, mArrayCols
        End If
        
        If mNeedIndex_FileStoimRazmesh = True Then
            Erase mArrayCols
            mArrayCols = Array(.mFileStoimRazmesh.mArt.mNum)
            mCreateArrIndexImportedFile .mFileStoimRazmesh.mName, mArrayCols
        End If
        
        If mNeedIndex_FileZatratProdviz = True Then
            Erase mArrayCols
            mArrayCols = Array(.mFileZatratProdviz.mColOzonSKU.mNum)
            mCreateArrIndexImportedFile .mFileZatratProdviz.mName, mArrayCols
        End If
        
        If mNeedIndex_FileOtchetPopr = True Then
            Erase mArrayCols
            mArrayCols = Array(.mFileOtchetPopr.mArt.mNum, .mFileOtchetPopr.mNachislType.mNum)
            mCreateArrIndexImportedFile .mFileOtchetPopr.mName, mArrayCols
        End If
        
    End With
    
    Erase mArrayCols
    
End Sub

Public Sub mCreateArrIndexImportedFile(ByVal mImpFileName As String, _
ByRef mArrNumCols() As Variant)
    
    Dim NumRow              As Long
    Dim mCol                As Long
    
    Select Case mImpFileName
        
        Case mImportedFile.mFileMasterTable.mName
            With mImportedFile.mFileMasterTable
                ReDim .mArrIndex(UBound(mArrNumCols) + 1, .mRange.Rows.Count)
                For NumRow = .mNumRowStart To .mRange.Rows.Count
                    For mCol = 0 To UBound(mArrNumCols)
                        .mArrIndex(mCol, NumRow) = CreateIndex(.mRange(NumRow, mArrNumCols(mCol)))
                    Next
                Next
            End With
        
        Case mImportedFile.mFileStoimRazmesh.mName
            With mImportedFile.mFileStoimRazmesh
                ReDim .mArrIndex(UBound(mArrNumCols) + 1, .mRange.Rows.Count)
                For NumRow = .mNumRowStart To .mRange.Rows.Count
                    For mCol = 0 To UBound(mArrNumCols)
                        .mArrIndex(mCol, NumRow) = CreateIndex(.mRange(NumRow, mArrNumCols(mCol)))
                    Next
                Next
            End With
        
        Case mImportedFile.mFileZatratProdviz.mName
            With mImportedFile.mFileZatratProdviz
                ReDim .mArrIndex(UBound(mArrNumCols) + 1, .mRange.Rows.Count)
                For NumRow = .mNumRowStart To .mRange.Rows.Count
                    For mCol = 0 To UBound(mArrNumCols)
                        .mArrIndex(mCol, NumRow) = CreateIndex(.mRange(NumRow, mArrNumCols(mCol)))
                    Next
                Next
            End With
        
        Case mImportedFile.mFileOtchetPopr.mName
            With mImportedFile.mFileOtchetPopr
                ReDim .mArrIndex(UBound(mArrNumCols) + 1, .mRange.Rows.Count)
                For NumRow = .mNumRowStart To .mRange.Rows.Count
                    For mCol = 0 To UBound(mArrNumCols)
                        .mArrIndex(mCol, NumRow) = CreateIndex(.mRange(NumRow, mArrNumCols(mCol)))
                    Next
                Next
            End With
            
    End Select
    
End Sub
