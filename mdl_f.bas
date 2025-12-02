Attribute VB_Name = "mdl_f"

Public Type mButtonInfo_
    
    mName           As String
    mCaption        As String
    mLine           As Long
    mPosition       As Long
    
End Type

Public Sub EnableActivityExcel()
    SetActivityExcel True
End Sub

Public Function SetActivityExcel(ByVal mValue As Boolean, Optional ByVal mOptionCalcOFF As Boolean)
    
    Application.EnableEvents = mValue
    Application.ScreenUpdating = mValue
    If mOptionCalcOFF = False Then
        Application.Calculation = IIf(mValue, xlAutomatic, xlManual)
    End If
    Application.DisplayAlerts = mValue
    
End Function

Public Function f_bool_ItsNumber(ByRef mData As Variant) As Boolean
    
    If Len(mData) = 0 Then GoTo Label_Err
    f_bool_ItsNumber = True
    
    On Error GoTo Label_Err
    
    Dim mTemp         As Double
    mTemp = mData * 1.1
    
    Exit Function
    
Label_Err:
    Err.Clear
    f_bool_ItsNumber = False
    
End Function

Public Function f_dbl_ValToDbl(ByRef mData As Variant) As Double
    
    If f_bool_ItsNumber(mData) = False Then
        f_dbl_ValToDbl = 0
        Exit Function
    End If
    
    If Len(mData) = 0 Or mData = "0" Then
        f_dbl_ValToDbl = 0
    Else
        f_dbl_ValToDbl = mData * 1
    End If
    
End Function

Public Function ClearFilterOnSheet(ByVal mSheetName As String)
    
    If ThisWorkbook.Sheets(mSheetName).AutoFilterMode = True Then
        If ThisWorkbook.Sheets(mSheetName).FilterMode = True Then
            ThisWorkbook.Sheets(mSheetName).AutoFilter.ShowAllData
        End If
    End If
    
End Function

Public Function mNow() As String

    Dim a, b As Double
    Dim c, d, e, f, g, h, i As String
    a = Now * 1
    b = Timer * 1
    c = Year(a)
    If Month(a) < 10 Then d = "0" & Month(a) Else d = Month(a)
    If Day(a) < 10 Then e = "0" & Day(a) Else e = Day(a)
    If Hour(a) < 10 Then f = "0" & Hour(a) Else f = Hour(a)
    If Minute(a) < 10 Then g = "0" & Minute(a) Else g = Minute(a)
    If Second(a) < 10 Then h = "0" & Second(a) Else h = Second(a)
    i = Int((b - Int(b)) * 1000)
    If i < 10 Then i = "0" & i
    If i < 100 Then i = "0" & i
    mNow = c & "." & d & "." & e & " " & f & "-" & g & "-" & h & "-" & i
    
End Function

Public Function mBackup(ByVal mTime As String) As Boolean
    
    On Error Resume Next
    
    Dim mFileName           As String
    Dim mPath               As String
    Dim mMaxLen             As Long
    Dim mFullPath           As String
    
    mFileName = FixEncode(ThisWorkbook.Name)
    mPath = ThisWorkbook.Path & Application.PathSeparator & "backup file " & mFileName
    mMaxLen = 218
    
    mFullPath = mPath & Application.PathSeparator & mTime & " - " & mFileName
    If Len(mFullPath) > mMaxLen Then
        mFullPath = mPath & Application.PathSeparator & mTime & " - " & "backup file." & _
            Split(mFileName, ".")(UBound(Split(mFileName, ".")))
        If Len(mFullPath) > mMaxLen Then
            mFullPath = ThisWorkbook.Path & Application.PathSeparator & mTime & " - " & "backup file " & mFileName
            If Len(mFullPath) > mMaxLen Then
                mFullPath = ThisWorkbook.Path & Application.PathSeparator & mTime & " - " & "backup file." & _
                    Split(mFileName, ".")(UBound(Split(mFileName, ".")))
                If Len(mFullPath) > mMaxLen Then
                    MsgBox "Не удалось сохранить backup file из-за превышения длины пути файла в " & _
                        mMaxLen & " симв", vbExclamation
                    Exit Function
                End If
            End If
        Else
            If Dir(mPath, vbDirectory) = "" Then MkDir (mPath)
        End If
    Else
        If Dir(mPath, vbDirectory) = "" Then MkDir (mPath)
    End If
        
    ThisWorkbook.SaveCopyAs Filename:=mFullPath
    
    If Err.Number <> 0 Then
        MsgBox Err.Number & vbCrLf & Err.Description
        Err.Clear
    End If
    
    On Error GoTo 0
    
End Function

Public Function FixEncode(ByVal Text As String) As String

    ' костыль к кодировкам
    
    Dim mSimb()           As String
    
    mSimb = Split("\./.:.*.?."".<.>.|", ".")
    
Line_Rep1:
    
    For indT = 1 To Len(Text)
        If AscW(Mid(Text, indT, 1)) = 1080 Then
            If indT + 1 <= Len(Text) Then
                If AscW(Mid(Text, indT + 1, 1)) = 774 Then
                    Text = Mid(Text, 1, indT - 1) & ChrW(1081) & Mid(Text, indT + 2)
                    GoTo Line_Rep1
                End If
            End If
        End If
    Next
    
    For indT = 1 To Len(Text)
        For indS = LBound(mSimb) To UBound(mSimb)
            If Mid(Text, indT, 1) = mSimb(indS) Then
                Text = Mid(Text, 1, indT - 1) & " " & Mid(Text, indT + 1)
            Else
                If Mid(Text, indT, 1) Like "[!A-Z]" And _
                Mid(Text, indT, 1) Like "[!a-z]" And _
                Mid(Text, indT, 1) Like "[!А-Я]" And _
                Mid(Text, indT, 1) Like "[!а-я]" And _
                Mid(Text, indT, 1) Like "[!0-9]" And _
                Mid(Text, indT, 1) Like "[!-`~^@$%(){}&_+=№#;:.,]" Then
                    Text = Mid(Text, 1, indT - 1) & " " & Mid(Text, indT + 1)
                End If
            End If
        Next
    Next
    
    FixEncode = Text
    
End Function

Public Sub ColorActiveCell()
    MsgBox ActiveCell.Interior.Color
End Sub

Public Function mRndLong(ByVal lMin As Long, ByVal lMax As Long) As Long
    Randomize
    mRndLong = Abs(Int(Rnd(1) * (lMax + 1 - lMin) + lMin))
End Function

Public Function mSetGrd(ByRef mRange As Range, ByVal mLineType As Long, ByVal mEdgeOrInside As Long)
    
    Select Case mLineType
        Case -1
            If mEdgeOrInside = 1 Then
                mRange.Borders(xlEdgeLeft).LineStyle = xlNone
                mRange.Borders(xlEdgeTop).LineStyle = xlNone
                mRange.Borders(xlEdgeBottom).LineStyle = xlNone
                mRange.Borders(xlEdgeRight).LineStyle = xlNone
            End If
            If mEdgeOrInside = 2 Then
                mRange.Borders(xlInsideVertical).LineStyle = xlNone
                mRange.Borders(xlInsideHorizontal).LineStyle = xlNone
            End If
        Case 0
            If mEdgeOrInside = 1 Then
                mRange.Borders(xlEdgeLeft).Weight = xlHairline
                mRange.Borders(xlEdgeTop).Weight = xlHairline
                mRange.Borders(xlEdgeBottom).Weight = xlHairline
                mRange.Borders(xlEdgeRight).Weight = xlHairline
            End If
            If mEdgeOrInside = 2 Then
                mRange.Borders(xlInsideVertical).Weight = xlHairline
                mRange.Borders(xlInsideHorizontal).Weight = xlHairline
            End If
        Case 1
            If mEdgeOrInside = 1 Then
                mRange.Borders(xlEdgeLeft).Weight = xlThin
                mRange.Borders(xlEdgeTop).Weight = xlThin
                mRange.Borders(xlEdgeBottom).Weight = xlThin
                mRange.Borders(xlEdgeRight).Weight = xlThin
            End If
            If mEdgeOrInside = 2 Then
                mRange.Borders(xlInsideVertical).Weight = xlThin
                mRange.Borders(xlInsideHorizontal).Weight = xlThin
            End If
        Case 2
            If mEdgeOrInside = 1 Then
                mRange.Borders(xlEdgeLeft).Weight = xlMedium
                mRange.Borders(xlEdgeTop).Weight = xlMedium
                mRange.Borders(xlEdgeBottom).Weight = xlMedium
                mRange.Borders(xlEdgeRight).Weight = xlMedium
            End If
            If mEdgeOrInside = 2 Then
                mRange.Borders(xlInsideVertical).Weight = xlMedium
                mRange.Borders(xlInsideHorizontal).Weight = xlMedium
            End If
    End Select
    
End Function

Public Function mNewSheet(ByRef mBook As Workbook, Optional ByVal mShName As String) As String
    
    mBook.Sheets.Add After:=mBook.Sheets(mBook.Sheets.Count)
    
    If mShName <> "" Then
        If mShName = "#" Then
            Dim mGen As String
            mGen = mCreateRndStr(2)
            Do While mSheetExist(mBook, mGen) <> False
                mGen = mCreateRndStr(2)
            Loop
            mBook.Sheets(mBook.Sheets.Count).Name = mGen
        Else
            mBook.Sheets(mBook.Sheets.Count).Name = mShName
        End If
    Else
        mBook.Sheets(mBook.Sheets.Count).Name = "temp " & mNow
    End If

    mNewSheet = mBook.Sheets(mBook.Sheets.Count).Name
    
End Function

Public Function mSheetExist(ByRef mBook As Workbook, ByVal mShName As String) As Boolean
    mSheetExist = False
    For Each mSh In mBook.Sheets
        If mSh.Name = mShName Then
            mSheetExist = True
            Exit Function
        End If
    Next
End Function

Public Function mDeleteSheet(ByRef mBook As Workbook, ByVal mShName As String)
    
    'Application.EnableEvents = False
    'Application.DisplayAlerts = False
    
    If mSheetExist(mBook, mShName) = True Then mBook.Sheets(mShName).Delete
    
    'Do While mSheetExist(mBook, mShName) = True
    '    mBook.Sheets(mShName).Delete
    '    Application.Wait (Now + TimeValue("0:00:01"))
    'Loop
    
    'Application.EnableEvents = True
    'Application.DisplayAlerts = True
    
End Function

Public Function mCreateRndStr(ByVal mLenStr As Long) As String
    
    Dim mCount As Long
    Dim mNum As Double
    
    Do While mCount < mLenStr
        mNum = Rnd(1)
        If Round(mNum * 100, 0) < 91 And Round(mNum * 100, 0) > 64 Then
            mCreateRndStr = mCreateRndStr & Chr(Round(mNum * 100, 0))
            mCount = mCount + 1
        End If
    Loop
    
End Function

Public Function mBookIsOpen(ByVal mPath As String) As Boolean
    
    Dim ind                                 As LongPtr
    mBookIsOpen = False
    For ind = 1 To Workbooks.Count
        If Workbooks.Item(ind).Name = Split(mPath, Application.PathSeparator)(UBound(Split(mPath, Application.PathSeparator))) Then
            mBookIsOpen = True
            Exit Function
        End If
    Next
    
End Function

Public Function f_OpenBook_book(ByVal str_Path As String) As Workbook
    
    On Error Resume Next
    
    Dim lng_Ind         As Long
    For lng_Ind = 1 To Workbooks.Count
        If Workbooks.Item(lng_Ind).Name = _
        Split(str_Path, Application.PathSeparator)(UBound(Split(str_Path, Application.PathSeparator))) Then
            Set f_OpenBook_book = Workbooks.Item(lng_Ind)
            Exit Function
        End If
    Next
    
    If UCase(Mid(str_Path, InStrRev(str_Path, "."))) = ".CSV" Then
        
        'Set f_OpenBook_book = Workbooks.Open(str_Path, Local:=True)
        
        Set f_OpenBook_book = Workbooks.Add
        f_OpenBook_book.Sheets(1).Cells.NumberFormat = "@"
        
        With f_OpenBook_book.Sheets(1).QueryTables.Add( _
        Connection:="TEXT;" & str_Path, _
        Destination:=f_OpenBook_book.Sheets(1).Range("$A$1"))
            
            .Name = "csv_file"
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = 65001
            .TextFileStartRow = 1
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = True
            .TextFileCommaDelimiter = False
            .TextFileSpaceDelimiter = False
            .TextFileColumnDataTypes = Array(2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, _
            2, 2)
            .TextFileTrailingMinusNumbers = True
            .Refresh
            
        End With
        
    Else
        Set f_OpenBook_book = Workbooks.Open(str_Path, UpdateLinks:=False)
    End If
    
End Function

Public Function mFilesInDir(ByVal mStartDir As String, ByVal mNeedFolders As Boolean, _
Optional ByVal mStartSizeArrPath As Long) As String()
    
    On Error GoTo Line_Err
    
    ' -----------------------------------------------------------------------------------------------
    
    Dim mFilePathNow            As String
    Dim mFileNum                As Long
    Dim mFilesFolders           As Long
    
    Dim mArrFinal()             As String
    Dim mLastIndexArr           As Long
    
    mFilePathNow = ""
    mFileNum = 0
    mFilesFolders = IIf(mNeedFolders = True, 16, 32)
    
    ReDim mArrFinal(mStartSizeArrPath - 1)
    mLastIndexArr = mStartSizeArrPath - 1
    
    If mStartSizeArrPath = 0 Then mStartSizeArrPath = 100
    
    ' -----------------------------------------------------------------------------------------------
    
    mFilePathNow = Dir("")
    mFilePathNow = Dir(Replace(mStartDir & "\*", "\\", "\"), mFilesFolders)
    
    Do While mFilePathNow <> ""
        If mFilePathNow = "." Or mFilePathNow = ".." Then GoTo Label_NextFile
        
        If mLastIndexArr < mFileNum + 1 Then
            mLastIndexArr = mLastIndexArr + mStartSizeArrPath
            ReDim Preserve mArrFinal(mLastIndexArr)
        End If
        
        mArrFinal(mFileNum + 1) = mStartDir & "\" & mFilePathNow
        mFileNum = mFileNum + 1
        
Label_NextFile:
        mFilePathNow = Dir
    Loop
    
    mArrFinal(0) = mFileNum
    mFilesInDir = mArrFinal
    Erase mArrFinal
    
    ' -----------------------------------------------------------------------------------------------
        
    Exit Function
    
Line_Err:
    Err.Number = 0
    Err.Clear
    mArrFinal(0) = "-1"
    mFilesInDir = mArrFinal
    Erase mArrFinal
    
End Function

Public Function mFileDialog(ByVal mTypeDialog As Office.MsoFileDialogType, _
ByVal mTitle As String, ByVal mButtonName As String, ByVal mMultiSelect As Boolean, _
ByVal mInitDirPathOrFullPathFile As String, ByVal mFilters As String) As String()
    
    Dim mFilterName             As String
    Dim mFilterExt              As String
    Dim mArrResult()            As String
    
    mFilterName = Split(mFilters, ",")(0)
    mFilterExt = Split(mFilters, ",")(1)
    
    With Application.FileDialog(mTypeDialog)
        
        .Title = mTitle
        .ButtonName = mButtonName
        .AllowMultiSelect = mMultiSelect
        
        .InitialFileName = mInitDirPathOrFullPathFile
        
        If mTypeDialog <> msoFileDialogSaveAs And mTypeDialog <> msoFileDialogFolderPicker Then
            .Filters.Clear
            .Filters.Add mFilterName, mFilterExt
            .FilterIndex = 1
        End If
        
        .InitialView = msoFileDialogViewDetails
        .Show
        
        For Each mItem In .SelectedItems
            If -1 = Not mArrResult Then
                ReDim mArrResult(0)
            Else
                ReDim Preserve mArrResult(UBound(mArrResult) + 1)
            End If
            mArrResult(UBound(mArrResult)) = mItem
        Next
        
        mFileDialog = mArrResult
        
    End With
    
    Erase mArrResult
    
End Function

Public Function f_str_ValToStr(ByRef mData As Variant) As String
    On Error GoTo Line_Err
    f_str_ValToStr = mData
    Exit Function
Line_Err:
    f_str_ValToStr = ""
    Err.Number = 0
    Err.Clear
    On Error GoTo 0
End Function

Public Function f_str_FindParam(ByVal str_ParamName As String) As Range
    
    Dim mRange                          As Range
    Dim mLng                            As Long
    
    Set mRange = ThisWorkbook.Sheets("#1").UsedRange
    
    For mLng = 2 To mRange.Rows.Count
        If mRange(mLng, 2) = str_ParamName Then
            Set f_str_FindParam = mRange(mLng, 4)
            Exit For
        End If
    Next
    
    Set mRange = Nothing
    
End Function

Public Function mFindNumColInRange(ByRef mRange As Range, ByRef mNumRowTitle As Long, _
ByRef mColName As String, ByVal mNeedCheckNextRow As Boolean, ByVal mReversToFndCol As Boolean) As Long
    
    mFindNumColInRange = -1
    
    Dim mNumCol             As Long
    Dim mStep               As Long
    Dim mColStart           As Long
    Dim mColFinish          As Long
    
    mColName = UCase(mColName)
    
    If mReversToFndCol = True Then
        mStep = -1
        mColStart = mRange.Columns.Count
        mColFinish = 1
    Else
        mStep = 1
        mColStart = 1
        mColFinish = mRange.Columns.Count
    End If
    
    For mNumCol = mColStart To mColFinish Step mStep
        If UCase(mRange(mNumRowTitle, mNumCol)) = mColName Then
            mFindNumColInRange = mNumCol
            Exit For
        End If
        If mNeedCheckNextRow = True Then
            If UCase(mRange(mNumRowTitle + 1, mNumCol)) = mColName Then
                mFindNumColInRange = mNumCol
                Exit For
            End If
        End If
    Next
    
End Function

Public Function mCheckCol(ByVal mFileName As String, ByRef mRange As Range, _
ByVal mNumRowTitle As Long, ByVal mShName As String, _
ByVal mNeedCheckNextRow As Boolean, ByVal mReversToFndCol As Boolean, _
ByRef mCol As mColumn_)
    
    With mCol
    If .mNum = 0 Then .mNum = mFindNumColInRange(mRange, mNumRowTitle, .mCaption, mNeedCheckNextRow, mReversToFndCol)
    If .mNum = -1 Then mColInFileNotFnd mFileName, mShName, .mCaption
    End With
    
End Function

Public Function mColInFileNotFnd(ByVal mFileName As String, ByVal mShName As String, ByVal mColName As String)
    
    MsgBox "В файле " & mFileName & vbCrLf & _
        "На листе " & mShName & vbCrLf & _
        "Не найдена колонка " & mColName & vbCrLf & _
        "Работа программы остановлена", vbCritical
    StopMySub True, True
    
End Function

Public Sub mArrangeButtonsSheetI()
    
    Dim mBtn()          As mButtonInfo_
    Dim mButtonW        As Double
    Dim mButtonH        As Double
    Dim mMargnW          As Double
    Dim mMargnH          As Double
    Dim mStartLeft      As Double
    Dim mStartTop       As Double
    
    ReDim mBtn(7)
    mButtonW = 130 '120
    mButtonH = 40
    mMargnW = 2 '3
    mMargnH = 0
    mStartLeft = mMargnW  '50
    mStartTop = -15
    
    Settings
    
    'For Each msh In ThisWorkbook.Sheets(mSheetI.mName).Shapes
    '    Debug.Print msh.Name
    'Next
    'End
    
    With mBtn(0):   .mLine = 1: .mPosition = 1:  .mName = "Button 1":    .mCaption = "1": End With
    With mBtn(1):   .mLine = 1: .mPosition = 2:  .mName = "Button 2":    .mCaption = "2": End With
    With mBtn(2):   .mLine = 1: .mPosition = 3:  .mName = "Button 6":    .mCaption = "3": End With
    With mBtn(3):   .mLine = 1: .mPosition = 4:  .mName = "Button 3":    .mCaption = "4": End With
    With mBtn(4):   .mLine = 1: .mPosition = 5:  .mName = "Button 7":    .mCaption = "5": End With
    
    With mBtn(5):   .mLine = 1: .mPosition = 6:  .mName = "Button 4":    .mCaption = "5": End With
    With mBtn(6):   .mLine = 1: .mPosition = 7:  .mName = "Button 5":    .mCaption = "5": End With
    With mBtn(7):   .mLine = 1: .mPosition = 8:  .mName = "Button 8":    .mCaption = "5": End With
    
    'With mBtn(5):   .mLine = 2: .mPosition = 4:  .mName = "Rounded Rectangle 1":    .mCaption = "Прайс закуп": End With
    'With mBtn(6):   .mLine = 2: .mPosition = 5:  .mName = "Button 6":    .mCaption = "Категория товара": End With
    
    For mIndArr = 0 To UBound(mBtn)
        With ThisWorkbook.Sheets(mSheetI.mName).Shapes(mBtn(mIndArr).mName)
            
            Select Case mBtn(mIndArr).mLine
                
                Case 1
                    .Width = mButtonW
                    .Height = mButtonH
                
                Case 2
                    .Width = 80
                    .Height = 30
                    
            End Select
            
            ThisWorkbook.Sheets(mSheetI.mName).Shapes.Range(Array(mBtn(mIndArr).mName)).Select
            Selection.Font.Name = "Arial" 'Arial ' Calibri
            Selection.Font.FontStyle = "обычный"
            Selection.Font.Size = 10
            
            ' for Rounded Rectangle N
            'Selection.ShapeRange.Adjustments.Item(1) = 0.25
            
        End With
    Next
    
    For mNumLine = 1 To 4
    For mNumPosition = 1 To 10
        For mIndArr = 0 To UBound(mBtn)
            If mBtn(mIndArr).mLine = mNumLine Then
            If mBtn(mIndArr).mPosition = mNumPosition Then
                With ThisWorkbook.Sheets(mSheetI.mName).Shapes(mBtn(mIndArr).mName)
                    
                    Select Case mNumLine
                        
                        Case 1
                            '.Left = mStartLeft + ((mNumPosition - 1) * (mButtonW + mMargnW))
                            If mNumPosition = 1 Then
                                .Left = mStartLeft
                            Else
                                .Left = ThisWorkbook.Sheets(mSheetI.mName).Shapes(mBtn(mIndArr - 1).mName).Left + _
                                    ThisWorkbook.Sheets(mSheetI.mName).Shapes(mBtn(mIndArr - 1).mName).Width + _
                                    mMargnW
                            End If
                            
                            .Top = mStartTop + (mNumLine * (mButtonH + mMargnH))
                    
                        Case 2
                            '.Left = mStartLeft + ((mNumPosition - 1) * (mButtonW + mMargnW))
                            
                            
                            .Top = mStartTop + (mNumLine * (mButtonH + mMargnH))
                            
                    End Select
                    
                End With
            End If
            End If
        Next
    Next
    Next
        
End Sub

Public Sub mInterfaceExcel(ByVal mVal As Boolean)
    
    'Dim iCommandBar As CommandBar
    
    With Application
        
        '.Caption = IIf(Value = True, Empty, ThisWorkbook.Name)
        .DisplayStatusBar = mVal
        .DisplayFormulaBar = mVal
        
        'Debug.Print .DisplayStatusBar
                
        For Each iCommandBar In .CommandBars
            iCommandBar.Enabled = mVal
            'Debug.Print iCommandBar.Name
        Next
        
        For Each mSheet In ThisWorkbook.Sheets
            mSheet.Activate
            With .ActiveWindow
            
                '.Caption = IIf(Value = True, .Parent.Name, "")
                .DisplayHeadings = mVal
                .DisplayGridlines = mVal
                .DisplayHorizontalScrollBar = mVal
                .DisplayVerticalScrollBar = mVal
                .DisplayWorkbookTabs = mVal
                
            End With
        Next
        
        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", " & mVal & ")"
        
    End With
    
End Sub

Public Function f_bool_CopyFile(ByVal str_FullPathFile_1 As String, _
ByVal str_FullPathFile_2 As String) As Boolean
    
    f_bool_CopyFile = False
    
    Dim obj_FSO         As Object
    
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    
    If Dir(str_FullPathFile_1) <> "" Then
        obj_FSO.CopyFile str_FullPathFile_1, str_FullPathFile_2
        If Dir(str_FullPathFile_2) <> "" Then
            f_bool_CopyFile = True
        End If
    End If
    
End Function

Public Function mFindExcelFilesInArr(ByRef mArrFiles() As String) As String()
    
    Dim mArrFinal()                 As String
    Dim mIndArrAllFiles                    As Long
    Dim mGoodFilesCount             As Long
    Dim mTempArr1()                 As String
    Dim mArrGoodExt()               As String
    Dim mIndArrExt                  As Long
    
    ReDim mArrFinal(0)
    mArrFinal(0) = "-1"
    mGoodFilesCount = 0
    mArrGoodExt = Split("XLS,XLSX,XLSM,XLSB", ",")
    
    For mIndArrAllFiles = 1 To UBound(mArrFiles)
        If InStr(mArrFiles(mIndArrAllFiles), ".") > 0 Then
            Erase mTempArr1
            mTempArr1 = Split(mArrFiles(mIndArrAllFiles), ".")
            For mIndArrExt = 0 To UBound(mArrGoodExt)
                If mArrGoodExt(mIndArrExt) = UCase(mTempArr1(UBound(mTempArr1))) Then
                    mGoodFilesCount = mGoodFilesCount + 1
                    mArrFinal(0) = mGoodFilesCount
                    ReDim Preserve mArrFinal(UBound(mArrFinal) + 1)
                    mArrFinal(UBound(mArrFinal)) = mArrFiles(mIndArrAllFiles)
                    Exit For
                End If
            Next
        End If
    Next
    
    mFindExcelFilesInArr = mArrFinal
    
    Erase mArrFinal
    Erase mTempArr1
    Erase mArrGoodExt
    
End Function

Public Function mKillOldErrSheets()
    
    Dim mArr()          As String
    Dim mIndArr         As Long
    
    For Each mSh In ThisWorkbook.Sheets
        If mSh.Name Like "[A-Z][A-Z]" Then
            If InStr(mSh.Range("A2"), "Это лист с перечнем ошибок") > 0 Then
                If -1 = Not mArr Then
                    ReDim mArr(0)
                Else
                    ReDim Preserve mArr(UBound(mArr) + 1)
                End If
                mArr(UBound(mArr)) = mSh.Name
            End If
        End If
    Next
    
    If Not (-1 = Not mArr) Then
        For mIndArr = 0 To UBound(mArr)
            mDeleteSheet ThisWorkbook, mArr(mIndArr)
        Next
    End If
    
    Erase mArr
    
End Function

Public Function mFindNumStartColsActions() As Long()
    
    Dim mNumCol             As Long
    Dim mArrRes()           As Long
    
    For mNumCol = 1 To mSheetI.mRange.Columns.Count
        'Debug.Print mSheetI.mRange(mSheetI.mNumRowStart - 1, mNumCol) & vbTab & mSheetI.mRange(mSheetI.mNumRowStart - 1, mNumCol).Address
        If mSheetI.mRange(mSheetI.mNumRowStart - 1, mNumCol) = mSheetI.mTextActionCol_1 Then
            If -1 = Not mArrRes Then
                ReDim mArrRes(0)
            Else
                ReDim Preserve mArrRes(UBound(mArrRes) + 1)
            End If
            mArrRes(UBound(mArrRes)) = mNumCol
        End If
    Next
    
    mFindNumStartColsActions = mArrRes
    Erase mArrRes
    
End Function

Public Function f_SortArray_arrvar(ByVal a As Variant) As Variant
    
    Dim i           As Long
    Dim j           As Long
    Dim t           As String
    
    If mColsInArr(a) = 1 Then
        For i = LBound(a) To UBound(a)
            For j = LBound(a) To UBound(a) - 1
                If a(j) > a(j + 1) Then
                    t = a(j)
                    a(j) = a(j + 1)
                    a(j + 1) = t
                End If
            Next
        Next
    Else
        For i = LBound(a, 1) To UBound(a, 1)
            For j = LBound(a, 1) To UBound(a, 1) - 1
                If a(j, 1) > a(j + 1, 1) Then
                    t = a(j, 1)
                    a(j, 1) = a(j + 1, 1)
                    a(j + 1, 1) = t
                End If
            Next
        Next
    End If
    
    f_SortArray_arrvar = a
    
End Function

Public Function mColsInArr(ByRef mArr As Variant) As Long
    
    On Error GoTo Line_Err
    
    Dim mIter         As Long
    Dim mSize         As Long
    
    Do: mIter = mIter + 1: mSize = UBound(mArr, mIter): Loop
    
Line_Err:
    
    mColsInArr = mIter - 1
    
End Function

Public Function f_SelectUniqueInArr_arrvar(ByVal mArrData As Variant) As Variant
    
    Dim mArr1           As Variant
    Dim mLng1           As Long
    Dim mLng2           As Long
    
    If UBound(mArrData) - LBound(mArrData) < 1 Then
        f_SelectUniqueInArr_arrvar = mArrData
        Exit Function
    End If
    
    For mLng1 = LBound(mArrData, 1) To UBound(mArrData, 1) - 1
        For mLng2 = mLng1 + 1 To UBound(mArrData, 1)
            If CStr(mArrData(mLng1, 1)) = CStr(mArrData(mLng2, 1)) Then
                mArrData(mLng2, 1) = ""
            End If
        Next
    Next
    
    mLng2 = 0
    For mLng1 = LBound(mArrData, 1) To UBound(mArrData, 1)
        If Len(mArrData(mLng1, 1)) > 0 Then mLng2 = mLng2 + 1
    Next
    
    ReDim mArr1(mLng2 - 1)
    
    mLng2 = 0
    For mLng1 = LBound(mArrData, 1) To UBound(mArrData, 1)
        If Len(mArrData(mLng1, 1)) > 0 Then
            mArr1(mLng2) = mArrData(mLng1, 1)
            mLng2 = mLng2 + 1
        End If
    Next
    
    f_SelectUniqueInArr_arrvar = mArr1
    
End Function

Public Sub mSetColWidthSheetI()
    
    'Dim mNumCol         As Long
    
    'For mNumCol = 6 To mSheetI.mRange.Columns.Count
    '    mSheetI.mRange.Columns(mNumCol).ColumnWidth = _
    '        mFindMaxLenText(mSheetI.mRange, mSheetI.mNumRowStart - 1, mSheetI.mRange.Rows.Count, mNumCol) + 3
    '    'mFindMaxLenText mSheetI.mRange, mSheetI.mNumRowStart - 1, mSheetI.mRange.Rows.Count, mNumCol
    'Next
    
    ThisWorkbook.Sheets(mSheetI.mName).Columns("F:AF").AutoFit
    
    ThisWorkbook.Sheets(mSheetI.mName).Columns(mSheetI.mColOstatokTekyshOZON.mNum).ColumnWidth = 7
    ThisWorkbook.Sheets(mSheetI.mName).Columns(mSheetI.mColOstatokIzMT.mNum).ColumnWidth = 7
    ThisWorkbook.Sheets(mSheetI.mName).Columns(mSheetI.mNacenProcentItog.mNum).ColumnWidth = 10
    
End Sub

Public Function mFindMaxLenText(ByRef mRange As Range, ByVal mStartRow As Long, _
ByVal mEndRow As Long, ByVal mCol As Long) As Long
    
    Dim mRes            As Long
    Dim mRow            As Long
    
    For mRow = mStartRow To mEndRow
        'Do While InStr(mRange(mRow, mCol).Text, "#") > 0
        '    mRange.Columns(mCol).EntireColumn.ColumnWidth = mRange.Columns(mCol).EntireColumn.ColumnWidth + 1
        'Loop
        
        If IsNumeric(mRange(mRow, mCol)) Or IsNumeric(mRange(mRow, mCol).Text) Then
            If IsNumeric(mRange(mRow, mCol)) Then
                If Len(Round(mRange(mRow, mCol), 2)) > mRes Then
                    mRes = Len(Round(mRange(mRow, mCol), 2))
                End If
            End If
            If IsNumeric(mRange(mRow, mCol).Text) Then
                If Len(Round(mRange(mRow, mCol).Text, 2)) > mRes Then
                    mRes = Len(Round(mRange(mRow, mCol).Text, 2))
                End If
            End If
        Else
            If Len(mRange(mRow, mCol)) > mRes Then
                mRes = Len(mRange(mRow, mCol))
            End If
            If Len(mRange(mRow, mCol).Text) > mRes Then
                mRes = Len(mRange(mRow, mCol).Text)
            End If
        End If
        
    Next
    
    If mRes < 5 Then mRes = 5
    mFindMaxLenText = mRes
    
End Function
