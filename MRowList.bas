Attribute VB_Name = "MRowList"
#If MSACCESS Then
Option Compare Database
#End If

Option Explicit

Private Const QBCOLOR_DUMP As Integer = 8 'GRAY

'Returns an array with 4 entries which are the lower and upper bounds of
'the first two dimensions of the array stored in pvVar.
'The number of dimensions is returned in piRetDims.
Public Sub GetVarArrayBounds(ByRef pvVar As Variant, ByRef piRetDims As Integer, ByRef pavRetBounds As Variant)
  Dim lLB1    As Long
  Dim lUB1    As Long
  Dim lLB2    As Long
  Dim lUB2    As Long
  
  On Error Resume Next
  piRetDims = 0
  lLB1 = LBound(pvVar)
  If Err.Number = 0 Then
    piRetDims = 1
    lUB1 = UBound(pvVar)
    lLB2 = LBound(pvVar, 2)
    If Err.Number = 0& Then
      piRetDims = 2
      lUB2 = UBound(pvVar, 2)
    End If
  End If
  pavRetBounds = Array(lLB1, lUB1, lLB2, lUB2)
End Sub

'
' Build a row from a serie of values; all columns will be unnamed.
'
Public Function MakeRow(ParamArray pavRowValues() As Variant)
  Dim lLB           As Long
  Dim lUB           As Long
  Dim i             As Long
  Dim j             As Long
  Dim rowRet        As CRow
  
  Set rowRet = New CRow
  
  'get the number of elements in pavPairs
  On Error Resume Next
  lLB = LBound(pavRowValues())
  lUB = UBound(pavRowValues())
  If Err.Number = 0& Then 'If we could get the lBound then we have at least one element
    j = 1
    For i = lLB To lUB
      rowRet.AddCol "Column " & j, pavRowValues(i), 0&, 0&
      j = j + 1
    Next i
  End If
  
  Set MakeRow = rowRet
  Set rowRet = Nothing
End Function

'
' Utilities for lists
'

Public Function MakeParamList(ParamArray pavPairs() As Variant) As CList
  Dim lLowerBound   As Long
  Dim lElemCount    As Long
  Dim lPairsCount   As Long
  Dim lPairIndex    As Long
  Dim lstRetPairs   As CList
  
  'get the number of elements in pavPairs
  On Error Resume Next
  lLowerBound = LBound(pavPairs())
  If Err.Number = 0& Then 'If we could get the lBound then we have at least one element
    lElemCount = UBound(pavPairs()) - lLowerBound + 1& 'Just in case 0 is not the lower bound
    If (lElemCount Mod 2&) = 0& Then 'We must have pairs (name, followed by value)
      Set lstRetPairs = New CList
      lstRetPairs.ArrayDefine Array("ParamName", "Value"), Array(vbString, vbVariant)
      lPairsCount = lElemCount \ 2&
      For lPairIndex = 1& To lPairsCount
        lstRetPairs.AddValues pavPairs((lLowerBound + (lPairIndex - 1&)) * 2&), _
                              pavPairs((lLowerBound + (lPairIndex - 1&)) * 2& + 1&)
      Next lPairIndex
      Set MakeParamList = lstRetPairs
    End If
  End If
End Function

Public Function MakeParamRow(ParamArray pavValues() As Variant) As CRow
  Dim lLowerBound   As Long
  Dim lUpperBound   As Long
  Dim i             As Long
  Dim iCol          As Long
  Dim rowValues     As CRow
  
  'get the number of elements in pavValues
  On Error Resume Next
  lLowerBound = LBound(pavValues())
  If Err.Number = 0& Then 'If we could get the lBound then we have at least one element
    lUpperBound = UBound(pavValues())
    Set rowValues = New CRow
    iCol = 1&
    For i = lLowerBound To lUpperBound
      rowValues.AddCol "#" & iCol, pavValues(i), 0&, 0&
      iCol = iCol + 1&
    Next i
    Set MakeParamRow = rowValues
  End If
End Function

'Split a string and insert elements in a list.
'Returns a new list object that must be freed setting it to nothing.
Public Function SplitToList(ByVal sToSplit As String, _
  Optional sSep As String = " ", _
  Optional lMaxItems As Long = 0&, _
  Optional eCompare As VbCompareMethod = vbBinaryCompare) _
  As CList

  Dim lPos        As Long
  Dim lDelimLen   As Long
  Dim oRetList    As CList
  
  On Error GoTo SplitToList_Err
  
  Set oRetList = New CList
  oRetList.Define "Item", vbString, 0&, 0&
  If Len(sToSplit) Then
    lDelimLen = Len(sSep)
    If lDelimLen Then
      lPos = InStr(1, sToSplit, sSep, eCompare)
      Do While lPos
        oRetList.AddValues left$(sToSplit, lPos - 1&)
        sToSplit = Mid$(sToSplit, lPos + lDelimLen)
        If lMaxItems Then
          If oRetList.Count = lMaxItems - 1& Then Exit Do
        End If
        lPos = InStr(1, sToSplit, sSep, eCompare)
      Loop
    End If
    oRetList.AddValues sToSplit
  End If
  Set SplitToList = oRetList
  Exit Function

SplitToList_Err:
  Set oRetList = Nothing
End Function

Public Function JoinList(ByRef lstToJoin As CList, Optional ByVal sColSep As String = ",", Optional ByVal sRowSep As String = vbCrLf, Optional ByVal psColFilter As String = "") As String
  Dim iCol        As Long
  Dim iRow        As Long
  Dim sRet        As String
  Dim asColName() As String
  Dim iColCt      As Long
  
  On Error Resume Next
  If Len(psColFilter) Then
    iColCt = SplitString(asColName(), psColFilter, ";")
  End If
  For iRow = 1& To lstToJoin.Count
    If iRow > 1& Then sRet = sRet & sRowSep
    If iColCt = 0& Then
      For iCol = 1& To lstToJoin.ColCount
        If iCol > 1& Then sRet = sRet & sColSep
        sRet = sRet & lstToJoin(iCol, iRow) & ""
      Next iCol
    Else
      For iCol = 1& To iColCt
        If iCol > 1& Then sRet = sRet & sColSep
        sRet = sRet & lstToJoin(asColName(iCol), iRow) & ""
      Next iCol
    End If
  Next iRow
  JoinList = sRet
End Function

Public Function VariantAsString(ByRef pvValue As Variant, Optional ByVal pfHexNumbers As Boolean = False) As String
  Dim vType         As VbVarType
  
  vType = VarType(pvValue)
  If vType = vbNull Then
    VariantAsString = "#null"
  ElseIf vType = vbObject Then
    VariantAsString = "#ref_" & TypeName(pvValue) & "_" & ObjPtr(pvValue)
  ElseIf vType = vbEmpty Then
    VariantAsString = "#empty"
  Else
    Select Case vType
    Case vbInteger, vbLong, vbByte, vbDouble, vbCurrency, vbDecimal
      If pfHexNumbers Then
        VariantAsString = "$" & LCase$(Hex$(pvValue))
      Else
        VariantAsString = pvValue
      End If
    Case Else
      VariantAsString = pvValue
    End Select
  End If
End Function

'Handle object and null values to go from variant to string.
'object --> "#ref"
'null --> "#null"
'empty --> "#empty"
Public Function LogListDump( _
    oList As CList, _
    Optional ByVal sTitle As String = "", _
    Optional ByVal psColWidths As String = "", _
    Optional ByVal plStartRow As Long = 0&, _
    Optional ByVal plEndRow As Long = 0&, _
    Optional ByVal pfDumpToString As Boolean = False, _
    Optional ByVal pfShowColTitles As Boolean = True, _
    Optional ByVal pfHexIntLongs As Boolean = False, _
    Optional ByVal pvForeColor As Variant = Null) As String
  Dim iRow      As Long
  Dim i         As Long
  Dim lCount    As Long
  Dim asColName()  As String
  Dim aiColWidth() As Integer
  Dim iLen      As Integer
  Dim iStart    As Long
  Dim iEnd      As Long
  Dim sRet      As String
  Dim lForeColor  As Long
  
  Dim iColWidthCt       As Integer
  Dim asColWidthSpec()  As String
  Dim iCol              As Integer
  Dim sColName          As String
  Dim sWidth            As String
  Dim iColon            As Integer
  
  On Error GoTo ListDump_Err
  
  If Not IsNull(pvForeColor) Then
    lForeColor = CLng(pvForeColor)
  Else
    lForeColor = QBCOLOR_DUMP
  End If
  
  lCount = oList.ColCount
  If lCount = 0& Then Exit Function
  ReDim aiColWidth(1 To lCount)
  If Len(sTitle) Then
    If pfDumpToString Then
      sRet = sRet & String$(Len(sTitle), "-") & "+" & vbCrLf
      sRet = sRet & sTitle & "|" & vbCrLf
    Else
      #If USE_CONOUT Then
        ConOutLn String$(Len(sTitle), "-") & "+", lForeColor
        ConOutLn sTitle & "|", lForeColor
      #Else
        Debug.Print String$(Len(sTitle), "-") & "+"
        Debug.Print sTitle & "|"
      #End If
    End If
  End If
  If Len(psColWidths) Then
    iColWidthCt = SplitString(asColWidthSpec(), psColWidths, ";")
    For i = 1 To iColWidthCt
      iColon = InStr(1, asColWidthSpec(i), ":")
      If iColon Then
        sColName = left$(asColWidthSpec(i), iColon - 1)
        sWidth = Right$(asColWidthSpec(i), Len(asColWidthSpec(i)) - iColon)
        If Len(sWidth) > 0 Then
          If Val(sWidth) > 0 Then
            iCol = oList.ColPos(sColName)
            If iCol Then
              aiColWidth(iCol) = Val(sWidth)
            End If
          End If
        End If
      End If
    Next i
  End If
  
  If pfShowColTitles Then
    'col titles row sep
    For i = 1 To lCount
      iLen = IIf(aiColWidth(i) = 0, Len(oList.ColName(i)), aiColWidth(i))
      If pfDumpToString Then
        sRet = sRet & String$(iLen, "-") & "+"
      Else
        #If USE_CONOUT Then
          ConOut String$(iLen, "-") & "+"
        #Else
          Debug.Print String$(iLen, "-") & "+";
        #End If
      End If
    Next i
    If pfDumpToString Then
      sRet = sRet & vbCrLf
    Else
      #If USE_CONOUT Then
        ConOutLn ""
      #Else
        Debug.Print
      #End If
    End If
    For i = 1 To lCount
      iLen = IIf(aiColWidth(i) = 0, Len(oList.ColName(i)), aiColWidth(i))
      If pfDumpToString Then
        sRet = sRet & StrBlock(oList.ColName(i), " ", iLen) & "|"
      Else
        #If USE_CONOUT Then
          ConOut StrBlock(oList.ColName(i), " ", iLen) & "|"
        #Else
          Debug.Print StrBlock(oList.ColName(i), " ", iLen) & "|";
        #End If
      End If
    Next i
    If pfDumpToString Then
      sRet = sRet & vbCrLf
    Else
      #If USE_CONOUT Then
        ConOutLn ""
      #Else
        Debug.Print
      #End If
    End If
    'col titles row sep
    For i = 1 To lCount
      iLen = IIf(aiColWidth(i) = 0, Len(oList.ColName(i)), aiColWidth(i))
      If pfDumpToString Then
        sRet = sRet & String$(iLen, "-") & "+"
      Else
        #If USE_CONOUT Then
          ConOut String$(iLen, "-") & "+"
        #Else
          Debug.Print String$(iLen, "-") & "+";
        #End If
      End If
    Next i
    If pfDumpToString Then
      sRet = sRet & vbCrLf
    Else
      #If USE_CONOUT Then
        ConOutLn ""
      #Else
        Debug.Print ""
      #End If
    End If
  End If
  
  'dump values
  '---------------------------------------------
  iStart = 1&
  iEnd = oList.Count
  If plStartRow > 0& Then
    If plStartRow <= oList.Count Then
      iStart = plStartRow
    End If
  End If
  If plEndRow > 0& Then
    If plEndRow <= oList.Count Then
      iEnd = plEndRow
    End If
  End If
  If iStart > iEnd Then
    'Swap
    Dim iTemp As Long
    iTemp = iEnd
    iEnd = iStart
    iStart = iTemp
  End If
  
  For iRow = iStart To iEnd
    For i = 1 To oList.ColCount
      iLen = IIf(aiColWidth(i) = 0, Len(oList.ColName(i)), aiColWidth(i))
      If pfDumpToString Then
        sRet = sRet & StrBlock(VariantAsString(oList(i, iRow), pfHexIntLongs) & "", " ", iLen) & "|"
      Else
        #If USE_CONOUT Then
          ConOut StrBlock(VariantAsString(oList(i, iRow), pfHexIntLongs) & "", " ", iLen) & "|"
        #Else
          Debug.Print StrBlock(VariantAsString(oList(i, iRow), pfHexIntLongs) & "", " ", iLen) & "|";
        #End If
      End If
    Next i
    If pfDumpToString Then
      sRet = sRet & vbCrLf
    Else
      #If USE_CONOUT Then
        ConOutLn ""
      #Else
        Debug.Print
      #End If
    End If
  Next iRow
  
ListDump_Exit:
  LogListDump = sRet
  Exit Function
  
ListDump_Err:
  'Stop
  Resume ListDump_Exit
  Resume
End Function

Private Function StringInVariantArray(ByVal psString As String, pavArray As Variant) As Boolean
  Dim i   As Integer
  For i = LBound(pavArray) To UBound(pavArray)
    If pavArray(i) = psString Then
      StringInVariantArray = True
      Exit Function
    End If
  Next i
End Function

'Define a list from another, with optional column filter.
'@pavColNames : column names from source list to copy definition; if not specified: all
Public Sub DefineListFromList( _
  ByVal plstToDefine As CList, _
  plstSource As CList, _
  Optional ByVal pavColNames As Variant = Null, _
  Optional ByVal psAdditionalObjectCol As String = "")
  
  Dim iSrc          As Integer
  Dim fIncludeCol   As Boolean
  Dim sColName      As String
  
  Dim asDefColName()  As String
  Dim aiDefColType()  As Integer
  Dim alDefColSize()  As Long
  Dim alDefColFlags() As Long
  Dim iDefColCount    As Integer
  
  For iSrc = 1 To plstSource.ColCount
    sColName = plstSource.ColName(iSrc)
    fIncludeCol = True
    If Not IsNull(pavColNames) Then
      If IsArray(pavColNames) Then
        fIncludeCol = StringInVariantArray(sColName, pavColNames)
      Else
        fIncludeCol = CBool(sColName = pavColNames)
      End If
    End If
    If fIncludeCol Then
      iDefColCount = iDefColCount + 1
      ReDim Preserve asDefColName(1 To iDefColCount) As String
      ReDim Preserve aiDefColType(1 To iDefColCount) As Integer
      ReDim Preserve alDefColSize(1 To iDefColCount) As Long
      ReDim Preserve alDefColFlags(1 To iDefColCount) As Long
      asDefColName(iDefColCount) = sColName
      aiDefColType(iDefColCount) = plstSource.ColType(iSrc)
      alDefColSize(iDefColCount) = plstSource.ColSize(iSrc)
      alDefColFlags(iDefColCount) = plstSource.ColFlags(iSrc)
    End If
  Next iSrc
  
  If Len(psAdditionalObjectCol) > 0 Then
    iDefColCount = iDefColCount + 1
    ReDim Preserve asDefColName(1 To iDefColCount) As String
    ReDim Preserve aiDefColType(1 To iDefColCount) As Integer
    ReDim Preserve alDefColSize(1 To iDefColCount) As Long
    ReDim Preserve alDefColFlags(1 To iDefColCount) As Long
    asDefColName(iDefColCount) = psAdditionalObjectCol
    aiDefColType(iDefColCount) = vbObject
    alDefColSize(iDefColCount) = 0
    alDefColFlags(iDefColCount) = 0
  End If
  If iDefColCount > 0 Then
    plstToDefine.ArrayDefine asDefColName(), aiDefColType(), alDefColSize(), alDefColFlags()
  End If
End Sub

'pasColNames, pavColValues : 1 based arrays
Public Function CompareListColsToValues(plstList As CList, ByVal piList As Integer, pasColNames() As String, pavColValues() As Variant) As Boolean
  Dim iColCount     As Integer
  Dim i             As Integer
  iColCount = UBound(pasColNames) - LBound(pasColNames) + 1
  For i = 1 To iColCount
    If plstList(pasColNames(i), piList) <> pavColValues(i) Then
      Exit Function
    End If
  Next i
  CompareListColsToValues = True
End Function

'Builds a new CList by grouping the rows of a CList on one or more of its
'column values.
'The resulting list has a first set of columns, with the names of the source
'list columns on which we group, and a supplemental column which is an embedded
'CList for each row, that will contain the source list occurences that are grouped
'under the group keys.
'The column containing the grouped occurences is named "__tuples", and the function
'will fail if any of the columns in the source list has this name.
'NOTES:
' Watch out with BIG lists, as this function works on a copy of the source list
' and generates another list, bigger than the source list.
' This function can raise errors !
' The source list must have columns of type vbObject !
' No column named "__tuples" in source list !
'@plstSource  : the source CList
'@pvGroupCols : can be either a col name or an array of col names from source
Public Function ListSortAndGroupBy( _
  ByRef plstSource As CList, _
  ByVal pvGroupCols As Variant, _
  Optional ByVal pvTuplesCols As Variant = "", _
  Optional ByVal psTuplesColName As String = "__tuples") As CList
  
  Dim iCol          As Integer
  
  If plstSource Is Nothing Then Exit Function
  If plstSource.Count = 0 Then Exit Function
  If plstSource.ColCount = 0 Then Exit Function
  
  'Check that we have no columns of type vbObject in the source list,
  'not supported at this time. Also test if we have a column with our reserved name.
  For iCol = 1 To plstSource.ColCount
    If plstSource.ColType(iCol) = vbObject Then
      Err.Raise 13, "ListSortAndGroupBy", "Unsupported object type of column [" & plstSource.ColName(iCol) & "]"
      Exit Function
    End If
    If StrComp(plstSource.ColName(iCol), psTuplesColName, vbTextCompare) = 0 Then
      Err.Raise 5, "ListSortAndGroupBy", "Forbidden column name [" & psTuplesColName & "] found in source list"
      Exit Function
    End If
  Next iCol
  
  'Build array of columns on which we want to group
  Dim asGroupCol()  As String
  Dim iGroupColsCt  As Integer
  Dim k             As Integer
  Dim sSortColumns  As String
  
  If IsArray(pvGroupCols) Then
    iGroupColsCt = UBound(pvGroupCols) - LBound(pvGroupCols) + 1
    ReDim asGroupCol(1 To iGroupColsCt) As String
    k = LBound(pvGroupCols)
    For iCol = 1 To iGroupColsCt
      If iCol > 1 Then
        sSortColumns = sSortColumns & ","
      End If
      asGroupCol(iCol) = pvGroupCols(k)
      sSortColumns = sSortColumns & asGroupCol(iCol)
      k = k + 1
    Next iCol
  Else
    iGroupColsCt = 1
    ReDim asGroupCol(1 To 1) As String
    asGroupCol(1) = CStr(pvGroupCols)
    sSortColumns = asGroupCol(1)
  End If
  
  'check that all key columns exist in lstSource
  For iCol = 1 To iGroupColsCt
    If Not plstSource.ColExists(asGroupCol(iCol)) Then
      Err.Raise 5, "ListSortAndGroupBy", "Column [" & asGroupCol(iCol) & "] not found in source list"
      Exit Function
    End If
  Next iCol
  
  Dim iTuplesColCount   As Integer
  ReDim asTupleColName(1 To 1) As String
  Dim iTupleCol         As Integer
  'Get tuples cols in array
  If IsArray(pvTuplesCols) Then
    iTuplesColCount = UBound(pvTuplesCols) - LBound(pvTuplesCols) + 1
    ReDim asTupleColName(1 To iTuplesColCount) As String
    iTupleCol = 1
    For k = LBound(pvTuplesCols) To UBound(pvTuplesCols)
      asTupleColName(iTupleCol) = pvTuplesCols(k)
      iTupleCol = iTupleCol + 1
    Next k
  Else
    If Len(pvTuplesCols) > 0 Then
      iTuplesColCount = 1
      asTupleColName(1) = pvTuplesCols
    End If
  End If
  'Check that all tuples cols exist in lstSource
  For iCol = 1 To iTuplesColCount
    If Not plstSource.ColExists(asTupleColName(iCol)) Then
      Err.Raise 5, "ListSortAndGroupBy", "Value Column [" & asTupleColName(iCol) & "] not found in source list"
      Exit Function
    End If
  Next iCol
  
  'Now we can safely execute the grouping
  Dim lstSource   As CList
  Set lstSource = New CList
  lstSource.CopyFrom plstSource
  lstSource.Sort sSortColumns
  
  Dim iList       As Long
  Dim lstResult   As CList
  Dim lstTuples   As CList
  Dim oRow        As CRow
  ReDim avValues(1 To iGroupColsCt) As Variant
  Dim fKeysChanged  As Boolean
  
  Set lstResult = New CList
  Dim oRowResult  As CRow
  DefineListFromList lstResult, lstSource, pvGroupCols, psTuplesColName
  Set lstTuples = New CList
  lstSource.DefineList lstTuples
  
  'create new result row
  Set oRowResult = New CRow
  lstResult.DefineRow oRowResult
  For k = 1 To iGroupColsCt
    avValues(k) = lstSource(asGroupCol(k), 1)
  Next k
  For iList = 1& To lstSource.Count
    fKeysChanged = Not CompareListColsToValues(lstSource, iList, asGroupCol, avValues)
    If fKeysChanged Then
      For k = 1 To iGroupColsCt
        oRowResult(asGroupCol(k)) = avValues(k)
      Next k
      Set oRowResult(psTuplesColName) = lstTuples
      lstResult.AddRow oRowResult
      
      'key(s) changed, create new result list row and child tuples list
      Set lstTuples = New CList
      If iTuplesColCount = 0 Then
        lstSource.DefineList lstTuples
      Else
        For iCol = 1 To iTuplesColCount
          lstTuples.AddCol asTupleColName(iCol), _
                           lstSource(asTupleColName(iCol), 1), _
                           lstSource.ColType(asTupleColName(iCol)), _
                           lstSource.ColFlags(asTupleColName(iCol))
        Next iCol
      End If
    
      'create new result row
      Set oRowResult = New CRow
      lstResult.DefineRow oRowResult
      
      'Get array of values to compare for detecting key change
      For k = 1 To iGroupColsCt
        avValues(k) = lstSource(asGroupCol(k), iList)
      Next k
      
    End If
    If iTuplesColCount = 0 Then
      lstSource.GetRow oRow, iList
    Else
      Set oRow = New CRow
      lstTuples.DefineRow oRow
      For iCol = 1 To iTuplesColCount
        oRow(iCol) = lstSource(asTupleColName(iCol), iList)
      Next iCol
    End If
    lstTuples.AddRow oRow
    'RowLineDump oRow
  Next iList
  For k = 1 To iGroupColsCt
    oRowResult(asGroupCol(k)) = avValues(k)
  Next k
  Set oRowResult(psTuplesColName) = lstTuples
  lstResult.AddRow oRowResult
  
  Set ListSortAndGroupBy = lstResult
End Function

'Builds a new CList by grouping the rows of a CList on one or more of its
'column values.
'The resulting list has a first set of columns, with the names of the source
'list columns on which we group, and a supplemental column which is an embedded
'CList for each row, that will contain the source list occurences that are grouped
'under the group keys.
'The column containing the grouped occurences is named "__tuples", and the function
'will fail if any of the columns in the source list has this name.
'NOTES:
' Watch out with BIG lists, as this function works on a copy of the source list
' and generates another list, bigger than the source list.
' This function can raise errors !
' The source list must have columns of type vbObject !
' No column named "__tuples" in source list !
'@plstSource  : the source CList
'@pvGroupCols : can be either a col name or an array of col names from source
Public Function ListGroupBy( _
  ByRef plstSource As CList, _
  ByVal pvGroupCols As Variant, _
  Optional ByVal pvTuplesCols As Variant = "", _
  Optional ByVal psTuplesColName As String = "__tuples") As CList
  
  Dim iCol          As Integer
  
  If plstSource Is Nothing Then Exit Function
  If plstSource.Count = 0 Then Exit Function
  If plstSource.ColCount = 0 Then Exit Function
  
  'Check that we have no columns of type vbObject in the source list,
  'not supported at this time. Also test if we have a column with our reserved name.
  For iCol = 1 To plstSource.ColCount
    If plstSource.ColType(iCol) = vbObject Then
      Err.Raise 13, "ListGroupBy", "Unsupported object type of column [" & plstSource.ColName(iCol) & "]"
      Exit Function
    End If
    If StrComp(plstSource.ColName(iCol), psTuplesColName, vbTextCompare) = 0 Then
      Err.Raise 5, "ListGroupBy", "Forbidden column name [" & psTuplesColName & "] found in source list"
      Exit Function
    End If
  Next iCol
  
  'Build array of columns on which we want to group
  Dim asGroupCol()  As String
  Dim iGroupColsCt  As Integer
  Dim k             As Integer
  Dim sSortColumns  As String
  
  If IsArray(pvGroupCols) Then
    iGroupColsCt = UBound(pvGroupCols) - LBound(pvGroupCols) + 1
    ReDim asGroupCol(1 To iGroupColsCt) As String
    k = LBound(pvGroupCols)
    For iCol = 1 To iGroupColsCt
      If iCol > 1 Then
        sSortColumns = sSortColumns & ","
      End If
      asGroupCol(iCol) = pvGroupCols(k)
      sSortColumns = sSortColumns & asGroupCol(iCol)
      k = k + 1
    Next iCol
  Else
    iGroupColsCt = 1
    ReDim asGroupCol(1 To 1) As String
    asGroupCol(1) = CStr(pvGroupCols)
    sSortColumns = asGroupCol(1)
  End If
  
  'check that all key columns exist in lstSource
  For iCol = 1 To iGroupColsCt
    If Not plstSource.ColExists(asGroupCol(iCol)) Then
      Err.Raise 5, "ListGroupBy", "Column [" & asGroupCol(iCol) & "] not found in source list"
      Exit Function
    End If
  Next iCol
  
  Dim iTuplesColCount   As Integer
  ReDim asTupleColName(1 To 1) As String
  Dim iTupleCol         As Integer
  'Get tuples cols in array
  If IsArray(pvTuplesCols) Then
    iTuplesColCount = UBound(pvTuplesCols) - LBound(pvTuplesCols) + 1
    ReDim asTupleColName(1 To iTuplesColCount) As String
    iTupleCol = 1
    For k = LBound(pvTuplesCols) To UBound(pvTuplesCols)
      asTupleColName(iTupleCol) = pvTuplesCols(k)
      iTupleCol = iTupleCol + 1
    Next k
  Else
    If Len(pvTuplesCols) > 0 Then
      iTuplesColCount = 1
      asTupleColName(1) = pvTuplesCols
    End If
  End If
  'Check that all tuples cols exist in lstSource
  For iCol = 1 To iTuplesColCount
    If Not plstSource.ColExists(asTupleColName(iCol)) Then
      Err.Raise 5, "ListGroupBy", "Value Column [" & asTupleColName(iCol) & "] not found in source list"
      Exit Function
    End If
  Next iCol
  
  'Now we can safely execute the grouping
  Dim lstSource   As CList
  Set lstSource = New CList
  Set lstSource = plstSource
  
  Dim iList       As Long
  Dim lstResult   As CList
  Dim lstTuples   As CList
  Dim oRow        As CRow
  ReDim avValues(1 To iGroupColsCt) As Variant
  Dim fKeysChanged  As Boolean
  
  Set lstResult = New CList
  Dim oRowResult  As CRow
  DefineListFromList lstResult, lstSource, pvGroupCols, psTuplesColName
  Set lstTuples = New CList
  lstSource.DefineList lstTuples
  
  'create new result row
  Set oRowResult = New CRow
  lstResult.DefineRow oRowResult
  
  Dim iResult As Integer
  For iList = 1& To lstSource.Count
    'search result list for line with key values
    For k = 1 To iGroupColsCt
      avValues(k) = lstSource(asGroupCol(k), iList)
    Next k
    iResult = lstResult.Find(asGroupCol, avValues)
    If iResult > 0 Then
      Set lstTuples = lstResult(psTuplesColName, iResult)
    Else
      'add result row
      Set oRowResult = New CRow
      lstResult.DefineRow oRowResult
      For k = 1 To iGroupColsCt
        oRowResult(asGroupCol(k)) = avValues(k)
      Next k
      
      Set lstTuples = New CList
      If iTuplesColCount = 0 Then
        lstSource.DefineList lstTuples
      Else
        For iCol = 1 To iTuplesColCount
          lstTuples.AddCol asTupleColName(iCol), _
                           lstSource(asTupleColName(iCol), 1), _
                           lstSource.ColType(asTupleColName(iCol)), _
                           lstSource.ColFlags(asTupleColName(iCol))
        Next iCol
      End If
      Set oRowResult(psTuplesColName) = lstTuples
      lstResult.AddRow oRowResult
    End If
    
    If iTuplesColCount = 0 Then
      lstSource.GetRow oRow, iList
    Else
      Set oRow = New CRow
      lstTuples.DefineRow oRow
      For iCol = 1 To iTuplesColCount
        oRow(iCol) = lstSource(asTupleColName(iCol), iList)
      Next iCol
    End If
    lstTuples.AddRow oRow
  Next iList
  
  Set ListGroupBy = lstResult
End Function

'Transforms a CRow to a CList with columns "ColName" and "ColValue"
Public Function RowToList(ByRef poRow As CRow) As CList
  Dim lstRes  As CList
  Dim iColCt  As Integer
  Dim i       As Integer
  Dim v       As Variant
  
  If poRow Is Nothing Then Exit Function
  
  Set lstRes = New CList
  lstRes.AddCol "ColName", "", 0, 0
  lstRes.AddCol "ColValue", v, 0, 0
  iColCt = poRow.ColCount
  For i = 1 To iColCt
    lstRes.AddValues poRow.ColName(i), poRow(i)
  Next i
  Set RowToList = lstRes
End Function
