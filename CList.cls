VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DEFAULT_GROWSIZE As Integer = 20
Private mlGrowSize          As Long

'We expose publicly the inner data array, to be able to
'let ADO populate it automatically with a GetRows() call.
Public DataArray              As Variant  '1st dim: Columns, 2nd dim: rows
Private mlDataArraySize       As Long     'Size of 2nd dim (rows) of DataArray
Private mlDataArraySlotCount  As Long     'Number of used rows (some may have been freed)

'Row handles storage
Private malIndex()          As Long
Private mlIdxArraySize      As Long     'Size of index array (upper bound)
Private mlIdxCount          As Long     'Number of elements in the DataArray

'For sorting the list on one or multiple columns
Private Type TSortInfo
  sColName        As String
  lColIndex       As Long
  fCaseSensitive  As Boolean
  fDesc           As Boolean
End Type
Private mfSorted            As Boolean  'If true, the class keeps the list elements sorted
Private msSortColumns       As String     'The sort definitions string ( [!]<colname>[+|-] )
Private matSortInfo()       As TSortInfo  'Information for each column on which we sort
Private miSortColCt         As Integer    'The number of columns on which we sort
'For searching the list
Enum eSearchListAlgorithm
  eSearchAlgoSequential
  eSearchAlgoDichotomic
End Enum
'[f]ind[i]nfo [f]lag values
Private Const FIF_CASESENSITIVE    As Long = 1&  'applies only to text fields.
Private Const FIF_HASJOKERS        As Long = 2&
Private Const FIF_HASWILDCARDS     As Long = 4&
Private Const FIF_EXACTSEARCH      As Long = 8&
Private Const FIF_ROOTSEARCH       As Long = 16&
Private Const FIF_SORTEDDESC       As Long = 32&
Private Type TFindInfo
  fUseSortKeys      As Boolean  'True if we can benefit from a sorted list, because we search on key fields
  lSearchFieldsCt   As Long     'Number of fields on which we search
  asSearchField()   As String   '1 to lSearchFieldsCt
  alColIndex()      As Long     '1 to lSearchFieldsCt
  avSearchValue()   As Variant  '1 to lSearchFieldsCt
  alFlags()         As Long     '1 to lSearchFieldsCt
  eSearchAlgo       As eSearchListAlgorithm
End Type
Private mtFindInfo  As TFindInfo

'The garbage queue for the rows array (circular array queue)
Private malGarbageQ()       As Long
Private mlGarbQSize         As Long
Private mlGarbQHead         As Long
Private mlGarbQTail         As Long
Private mlGarbQCount        As Long

'
' Columns
'
Private Type TColDef
  iDataType   As Integer
  lDataSize   As Long
  lFlags      As Long
  sColName    As String
End Type
Private matColDef()       As TColDef
Private malColIndex()     As Long
Private mlColCount        As Long
Private moColIndexMap     As CMapStringToLong 'Store colname-->colindex map

'Special column flags
Private Const CF_SORTEDCASESENSITIVE     As Long = &H10000
Private Const CF_SORTEDCASEINSENSITIVE   As Long = &H20000

'Interface for external comparison of 2 object refs or values
Private miiCompObjects    As IListCompare

#If Win64 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
              lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
#End If

'
' Constructor / destructor
'
Private Sub Class_Initialize()
  mlGrowSize = DEFAULT_GROWSIZE
  mfSorted = False
  'columns
  Set moColIndexMap = New CMapStringToLong  'case sensitive=false, by default
  moColIndexMap.Sorted = True
  Clear
End Sub

Private Sub Class_Terminate()
  'Debug.Print "CList::Terminate#" & ObjPtr(Me) '/**/
  Set miiCompObjects = Nothing
  Set moColIndexMap = Nothing
End Sub

'Clear removes the rows *AND* the column definitions.
'Use .Reset to remove the rows only
Public Sub Clear()
  Reset
  'clear sort info
  If miSortColCt Then Erase matSortInfo
  miSortColCt = 0
  'clear columns info
  If mlColCount Then
    Erase matColDef
    Erase malColIndex
    mlColCount = 0&
  End If
  moColIndexMap.Clear
End Sub

Public Sub Reset()
  If mlIdxCount Then Erase DataArray
  mlDataArraySlotCount = 0&
  mlDataArraySize = 0&
  'Clear index array
  mlIdxArraySize = mlGrowSize
  ReDim malIndex(1 To mlIdxArraySize)
  mlIdxCount = 0&
  'Clear garbage queue
  mlGarbQSize = mlGrowSize
  ReDim malGarbageQ(1 To mlGarbQSize)
  mlGarbQHead = 1&
  mlGarbQTail = 0&
  mlGarbQCount = 0&
  'clear special column flags
  ClearAllSysColFlags
End Sub

Public Property Get IIListCompare() As IListCompare
  Set IIListCompare = miiCompObjects
End Property

Public Property Set IIListCompare(ByRef piiListCompare As IListCompare)
  Set miiCompObjects = piiListCompare
End Property

Public Property Get ColCaseSensitive() As Boolean
  ColCaseSensitive = moColIndexMap.CaseSensitive
End Property

Public Property Let ColCaseSensitive(ByVal pfColCaseSensitive As Boolean)
  moColIndexMap.CaseSensitive = pfColCaseSensitive
End Property

'2020-11-02 FFO - Knowing that we don't have garbaged
'rows in out list is knwing that we can access the
'underlying data array directly (dangerous optimisations altough)
Public Property Get HasGarbage() As Boolean
  HasGarbage = CBool(mlGarbQCount > 0)
End Property

'2020-11-02 FFO - Again, dangerous one, but needed if we
'still want to access the Data Array "indirectly", if it has garbage
Public Property Get DataArrayRowIndex(ByVal plRowIndex As Long) As Long
  DataArrayRowIndex = malIndex(plRowIndex)
End Property

Public Sub SyncWithDataArray()
  Dim i     As Long
  On Error Resume Next
  mlIdxCount = UBound(DataArray, 2)
  If Err.Number = 0 Then
    mlIdxCount = mlIdxCount - LBound(DataArray) + 1&
  End If
  On Error GoTo 0
  If mlIdxCount Then
    If mlIdxCount > mlGrowSize Then
      mlIdxArraySize = ((mlIdxCount \ mlGrowSize) * mlGrowSize) + (mlIdxCount Mod mlGrowSize)
    Else
      mlIdxArraySize = mlGrowSize
    End If
    ReDim malIndex(1 To mlIdxArraySize)
    'For i = 1& To mlIdxCount
    For i = 1& To mlIdxCount
      malIndex(i) = i - 1&
    Next i
    mlDataArraySize = mlIdxCount
    mlDataArraySlotCount = mlIdxCount
  Else
    mlDataArraySize = 0&
    mlDataArraySlotCount = 0&
    mlIdxArraySize = mlGrowSize
    ReDim malIndex(1 To mlIdxArraySize)
  End If
  ClearAllSysColFlags
End Sub

Public Property Get ColCount() As Long
  ColCount = mlColCount
End Property

Public Sub AddCol(ByRef psColName As String, _
                  ByVal pvTemplateValue As Variant, _
                  ByVal plDataSize As Long, _
                  ByVal plFlags As Long, _
                  Optional ByVal plInsertAfter As Long = 0&, _
                  Optional ByVal plInsertBefore As Long = 0&)
  If Len(psColName) Then
    If moColIndexMap.Find(psColName) Then
      'This key is already associated with an element of this collection
      Err.Raise 457&, "CRow::AddCol", "Duplicate column names not allowed"
      Exit Sub
    End If
  End If
  
  'All column contents lost, because we modify the first dim
  If mlColCount Then
    ReDim DataArray(0 To mlColCount, 0 To 0)
  Else
    ReDim DataArray(0 To 0, 0 To 0)
  End If
  
  Dim lColIndex   As Long
  Dim i           As Long
  ReDim Preserve matColDef(1 To mlColCount + 1&)
  ReDim Preserve malColIndex(1 To mlColCount + 1&)
  If plInsertAfter Then
    lColIndex = plInsertAfter + 1&
    'Push down other elements
    If lColIndex <= mlColCount Then
      CopyMemory malColIndex(lColIndex + 1&), malColIndex(lColIndex), (mlColCount - plInsertAfter) * LenB(lColIndex)
      For i = 1 To moColIndexMap.Count
        If moColIndexMap.Item(i) >= lColIndex Then
          moColIndexMap.Item(i) = moColIndexMap.Item(i) + 1&
        End If
      Next i
    End If
  ElseIf plInsertBefore Then
    'Push down other elements
    lColIndex = plInsertBefore
    If mlColCount Then
      CopyMemory malColIndex(lColIndex + 1&), malColIndex(lColIndex), (mlColCount - lColIndex + 1&) * LenB(lColIndex)
      For i = 1 To moColIndexMap.Count
        If moColIndexMap.Item(i) >= lColIndex Then
          moColIndexMap.Item(i) = moColIndexMap.Item(i) + 1&
        End If
      Next i
    End If
  Else
    'Append at end of columns
    lColIndex = mlColCount + 1&
  End If
  mlColCount = mlColCount + 1&
  malColIndex(lColIndex) = mlColCount
  If Len(psColName) Then moColIndexMap.Add psColName, lColIndex
  'Add info
  With matColDef(mlColCount)
    .iDataType = VarType(pvTemplateValue)
    .lDataSize = plDataSize
    .lFlags = plFlags And &HFFFF&
    .sColName = psColName
  End With
End Sub

Public Function ColPos(ByVal psColName As String) As Long
  Dim lColIndex As Long
  If left$(psColName, 1) <> "#" Then
    lColIndex = moColIndexMap.Find(psColName)
    If lColIndex = 0& Then
      'Error 5 (Invalid procedure call)
      Err.Raise 5&, "CRow::Value [Get]", "<" & psColName & ">: Item not found"
      Exit Function
    End If
    ColPos = moColIndexMap.Item(lColIndex)
  Else
    ColPos = CLng(Val(Right$(psColName, Len(psColName) - 1)))
  End If
End Function

Public Function ColExists(ByVal psColName As String) As Boolean
  ColExists = CBool(moColIndexMap.Find(psColName))
End Function

Public Property Get ColName(ByVal plColIndex As Long) As String
  ColName = matColDef(malColIndex(plColIndex)).sColName
End Property

Public Property Let ColName(ByVal plColIndex As Long, ByVal psNewName As String)
  Dim lMappedLong   As Long
  Dim lMapIndex     As Long
  Dim i             As Long
  
  'Column name must be given
  If Len(psNewName) = 0 Then
    Err.Raise 5&, "CRow::ColName [Let]", "Missing column name."
    Exit Property
  End If
  If moColIndexMap.Find(psNewName) Then
    'This key is already associated with an element of this collection
    Err.Raise 457&, "CRow::ColName [Let]", "Duplicate column names not allowed."
    Exit Property
  End If
  'Sequentially search for the index in the map
  For i = 1 To moColIndexMap.Count
    If moColIndexMap.Item(i) = plColIndex Then
      lMapIndex = i
      Exit For
    End If
  Next i
  If lMapIndex Then
    moColIndexMap.Remove lMapIndex
    moColIndexMap.Add psNewName, plColIndex
    matColDef(malColIndex(plColIndex)).sColName = psNewName
  Else
    'Bad column index, error 9: "Subscript out of range"
    Err.Raise 9&, "CRow::ColName [Let]", VBA.Error$(9&)
  End If
End Property

Public Property Get ColType(ByVal pvIndex As Variant) As Integer
  Dim lColIndex     As Long
  If VarType(pvIndex) = vbString Then
    lColIndex = ColPos(pvIndex)
  Else
    lColIndex = pvIndex
  End If
  ColType = matColDef(malColIndex(lColIndex)).iDataType
End Property

Public Property Let ColType(ByVal pvIndex As Variant, ByVal piNewType As Integer)
  Dim lColIndex     As Long
  If VarType(pvIndex) = vbString Then
    lColIndex = ColPos(pvIndex)
  Else
    lColIndex = pvIndex
  End If
  matColDef(malColIndex(lColIndex)).iDataType = piNewType
End Property

Public Property Get ColSize(ByVal pvIndex As Variant) As Long
  Dim lColIndex     As Long
  If VarType(pvIndex) = vbString Then
    lColIndex = ColPos(pvIndex)
  Else
    lColIndex = pvIndex
  End If
  ColSize = matColDef(malColIndex(lColIndex)).lDataSize
End Property

Public Property Let ColSize(ByVal pvIndex As Variant, ByVal plNewSize As Long)
  Dim lColIndex     As Long
  If VarType(pvIndex) = vbString Then
    lColIndex = ColPos(pvIndex)
  Else
    lColIndex = pvIndex
  End If
  matColDef(malColIndex(lColIndex)).lDataSize = plNewSize
End Property

Public Property Get ColFlags(ByVal pvIndex As Variant) As Long
  Dim lColIndex     As Long
  If VarType(pvIndex) = vbString Then
    lColIndex = ColPos(pvIndex)
  Else
    lColIndex = pvIndex
  End If
  ColFlags = matColDef(malColIndex(lColIndex)).lFlags
End Property

Public Property Let ColFlags(ByVal pvIndex As Variant, ByVal plNewFlags As Long)
  Dim lColIndex     As Long
  If VarType(pvIndex) = vbString Then
    lColIndex = ColPos(pvIndex)
  Else
    lColIndex = pvIndex
  End If
  'bit values above &HFFFF& are reserved for class internal use,
  'so we just mask the passed flags with &HFFFF&
  matColDef(malColIndex(lColIndex)).lFlags = plNewFlags And &HFFFF&
End Property

'Define list from a flat array.
'Elements must be grouped 4 by 4 for a column definition, and in this
'order: name, type, size, flags.
Public Sub Define(ParamArray pavDefs() As Variant)
  Dim lLB       As Long
  Dim lUB       As Long
  Dim lCount    As Long
  Dim lIndex    As Long
  Dim sColName  As String
  Dim iDataType As Integer
  Dim lDataSize As Long
  Dim lFlags    As Long
  Dim lCol      As Long
  
  lLB = LBound(pavDefs)
  lUB = UBound(pavDefs)
  'There must be a multiple of 4 element count
  If (lUB - lLB + 1&) Mod 4 Then
    Err.Raise 5&, "CList::Define", "Incorrect number of array elements."
    Exit Sub
  End If
  lCount = (lUB - lLB + 1&) \ 4
  Clear
  lCol = 1&
  For lIndex = 1& To lCount
    sColName = pavDefs(lLB + (lCol - 1&) * 4&)
    iDataType = pavDefs(lLB + (lCol - 1&) * 4& + 1&)
    lDataSize = pavDefs(lLB + (lCol - 1&) * 4& + 2&)
    lFlags = pavDefs(lLB + (lCol - 1&) * 4& + 3&)
    If moColIndexMap.Find(sColName) = 0 Then
      ReDim Preserve matColDef(1 To mlColCount + 1&)
      ReDim Preserve malColIndex(1 To mlColCount + 1&)
      'Append at end of columns
      mlColCount = mlColCount + 1&
      malColIndex(mlColCount) = mlColCount
      moColIndexMap.Add sColName, mlColCount
      'Add info
      With matColDef(mlColCount)
        .iDataType = iDataType
        .lDataSize = lDataSize
        .lFlags = lFlags And &HFFFF&
        .sColName = sColName
      End With
    End If
    lCol = lCol + 1&
  Next lIndex
  ReDim DataArray(0 To mlColCount - 1, 0 To 0)
End Sub

Public Sub ArrayDefine(pavColName As Variant, _
                       Optional pavDataType As Variant, _
                       Optional pavDataSize As Variant, _
                       Optional pavDataFlags As Variant)
  Dim lLB       As Long
  Dim lUB       As Long
  Dim lIndex    As Long
  Dim sColName  As String
  Dim iDataType As Integer
  Dim lDataSize As Long
  Dim lFlags    As Long
  
  lLB = LBound(pavColName)
  lUB = UBound(pavColName)
  Clear
  For lIndex = lLB To lUB
    sColName = pavColName(lIndex)
    If Not IsMissing(pavDataType) Then
      iDataType = pavDataType(lIndex)
    Else
      iDataType = VbVarType.vbString
    End If
    If Not IsMissing(pavDataSize) Then lDataSize = pavDataSize(lIndex)
    If Not IsMissing(pavDataFlags) Then lFlags = pavDataFlags(lIndex)
    If moColIndexMap.Find(sColName) = 0 Then
      ReDim Preserve matColDef(1 To mlColCount + 1&)
      ReDim Preserve malColIndex(1 To mlColCount + 1&)
      'Append at end of columns
      mlColCount = mlColCount + 1&
      malColIndex(mlColCount) = mlColCount
      moColIndexMap.Add sColName, mlColCount
      'Add info
      With matColDef(mlColCount)
        .iDataType = iDataType
        .lDataSize = lDataSize
        .lFlags = lFlags And &HFFFF&
        .sColName = sColName
      End With
    End If
  Next lIndex
  ReDim DataArray(0 To mlColCount - 1, 0 To 0)
End Sub

'
' CList specific methods
'

Public Property Get GrowSize() As Long
  GrowSize = mlGrowSize
End Property

Public Property Let GrowSize(ByVal plGrowSize As Long)
  mlGrowSize = plGrowSize
End Property

Public Property Get Count() As Long
  Count = mlIdxCount
End Property

'
' Sorting, public methods
'

'Sort columns string format: [!]Column name[+|-][, [!]ColumnName2[+|-] [, ...] ]
Private Sub ParseSortColumns(ByRef psSortColumns As String)
  Dim asColName()     As String
  Dim sColName        As String
  Dim i               As Integer
  Dim sCaseSensitive  As String
  Dim fCaseSensitive  As Boolean
  Dim sAscDesc        As String
  
  'If the sort columns definition is empty, we reinit and exit
  If Len(psSortColumns) = 0 Then
    If miSortColCt Then Erase matSortInfo()
    miSortColCt = 0
    mfSorted = False
    Exit Sub
  End If
  'each column on which we'll have to sort is separated by a comma
  miSortColCt = SplitString(asColName(), psSortColumns, ",")
  ReDim matSortInfo(1 To miSortColCt)
  For i = 1 To miSortColCt
    sColName = asColName(i)
    sCaseSensitive = left$(sColName, 1)
    sAscDesc = Right$(sColName, 1)
    With matSortInfo(i)
      'The column name must be banged (preceeded by "!") for the sort to be case sensitive
      If sCaseSensitive = "!" Then
        .fCaseSensitive = True
        sColName = Right$(sColName, Len(sColName) - 1)
      Else
        .fCaseSensitive = False
      End If
      'Last char of column name may indicate asc or desc sorting
      If (sAscDesc = "+") Or (sAscDesc = "-") Then
        .fDesc = CBool(sAscDesc = "-")
        sColName = left$(sColName, Len(sColName) - 1)
      Else
        .fDesc = False
      End If
      .sColName = sColName
      On Error Resume Next
      .lColIndex = ColPos(sColName)
      If .lColIndex = 0& Then
        'this can't be a good thing: we can't sort on an unknown column
        '(we borrow VB error 5 (invalid procedure call); we could also
        'have used error 9 (subscript out of range))
        Err.Raise 5&, "CList::ParseSortColumns", sColName & ": is not an existing column; can't sort list"
        Exit Sub
      End If
    End With
  Next i
End Sub

'(Complementary function for handling sort info)
'Need to now if a column is part of the sort key
Private Function IsSortColumn(ByVal plColIndex As Long) As Boolean
  Dim i               As Integer
  Dim sCompColName    As String
  Dim eCompMethod     As VbCompareMethod
  
  If miSortColCt Then
    If moColIndexMap.CaseSensitive Then
      eCompMethod = vbBinaryCompare
    Else
      eCompMethod = vbTextCompare
    End If
    sCompColName = ColName(plColIndex)
    For i = 1 To miSortColCt
      If StrComp(matSortInfo(i).sColName, sCompColName, eCompMethod) = 0 Then
        IsSortColumn = True
        Exit Function
      End If
    Next i
  End If
End Function

Public Property Get Sorted() As Boolean
  Sorted = mfSorted
End Property

Public Sub SetUnsorted()
  If miSortColCt Then Erase matSortInfo()
  mfSorted = False: miSortColCt = 0: ClearAllSysColFlags
End Sub

Public Sub SetSorted(ByVal psSortColumns As String)
  Dim i         As Integer
  Dim iColPos   As Long
  
  ParseSortColumns psSortColumns
  msSortColumns = psSortColumns
  If miSortColCt Then
    For i = 1 To miSortColCt
      iColPos = matSortInfo(i).lColIndex
      If matSortInfo(i).fCaseSensitive Then
        SetSysColFlags iColPos, CF_SORTEDCASESENSITIVE
      Else
        SetSysColFlags iColPos, CF_SORTEDCASEINSENSITIVE
      End If
    Next i
    mfSorted = True
  End If
End Sub

Public Sub Sort(ByVal psSortColumns As String, Optional ByVal plStartRow As Long = 1&, Optional ByVal plEndRow As Long = 0&)
  Dim i         As Integer
  Dim iColPos   As Long
  
  If plEndRow = 0& Then plEndRow = mlIdxCount
  ClearAllSysColFlags
  ParseSortColumns psSortColumns 'Will raise an error if a column name doesn't exist or map
  msSortColumns = psSortColumns
  
  If miSortColCt Then
    mfSorted = False
    'resolve sort columns indices
    QuickSortRows plStartRow, plEndRow
    If (plStartRow = 1&) And (plEndRow = mlIdxCount) Then
      'set columns sys flags
      For i = 1 To miSortColCt
        iColPos = matSortInfo(i).lColIndex
        If matSortInfo(i).fCaseSensitive Then
          SetSysColFlags iColPos, CF_SORTEDCASESENSITIVE
        Else
          SetSysColFlags iColPos, CF_SORTEDCASEINSENSITIVE
        End If
      Next i
      mfSorted = True
    End If
  End If
End Sub

Public Property Get SortColumns()
  SortColumns = msSortColumns
End Property

Public Property Get Item(ByVal pvColIndex As Variant, ByVal plRowIndex As Long) As Variant
Attribute Item.VB_UserMemId = 0
  Dim lColIndex     As Long
  Dim lRealRowIndex As Long
  '/**/ check that row index (plRowIndex) is not out of bounds
  If VarType(pvColIndex) = vbString Then
    lColIndex = ColPos(pvColIndex) - 1&
  Else
    lColIndex = pvColIndex - 1&
  End If
  lRealRowIndex = malIndex(plRowIndex)
  If Not IsObject(DataArray(lColIndex, lRealRowIndex)) Then
    Item = DataArray(lColIndex, lRealRowIndex)
  Else
    Set Item = DataArray(lColIndex, lRealRowIndex)
  End If
End Property

'IsItemObject allows the class user to determine if the list item if an object reference, before reading it.
Public Function IsItemObject(ByVal pvColIndex As Variant, ByVal plRowIndex As Long) As Boolean
  Dim lColIndex     As Long
  Dim lRealRowIndex As Long
  If VarType(pvColIndex) = vbString Then
    lColIndex = ColPos(pvColIndex) - 1&
  Else
    lColIndex = pvColIndex - 1&
  End If
  lRealRowIndex = malIndex(plRowIndex)
  IsItemObject = IsObject(DataArray(lColIndex, lRealRowIndex))
End Function

Public Property Let Item(ByVal pvColIndex As Variant, ByVal plRowIndex As Long, ByRef pvCellValue As Variant)
  Dim lColIndex     As Long
  If VarType(pvColIndex) = vbString Then
    lColIndex = ColPos(pvColIndex)
  Else
    lColIndex = pvColIndex
  End If
  DataArray(lColIndex - 1, malIndex(plRowIndex)) = pvCellValue
  If IsSortColumn(lColIndex) Then mfSorted = False
End Property

Public Property Set Item(ByVal pvColIndex As Variant, ByVal plRowIndex As Long, ByRef poObjectRef As Object)
  Dim lColIndex     As Long
  If VarType(pvColIndex) = vbString Then
    lColIndex = ColPos(pvColIndex)
  Else
    lColIndex = pvColIndex
  End If
  Set DataArray(lColIndex - 1, malIndex(plRowIndex)) = poObjectRef
  If IsSortColumn(lColIndex) Then mfSorted = False
End Property

'
' Add row data
'

Public Function AddRow(prowToAdd As CRow, _
                       Optional ByVal plInsertAfter As Long = 0&, _
                       Optional ByVal plInsertBefore As Long = 0&) As Long
  Dim lDataIndex    As Long
  Dim lCol          As Long
  Dim lRowIndex     As Long
  
  'Grow index array if necessary
  If mlIdxCount = mlIdxArraySize Then
    ReDim Preserve malIndex(1 To mlIdxArraySize + mlGrowSize)
    mlIdxArraySize = mlIdxArraySize + mlGrowSize
  End If
  'Find a slot where to insert the new row data
  lDataIndex = AllocRow(prowToAdd)
  'Adding a row breaks sorted state
  SetUnsorted
  If plInsertAfter Then
    lRowIndex = plInsertAfter + 1&
    'Push down other index pointers
    If lRowIndex < mlIdxCount Then
      CopyMemory malIndex(lRowIndex + 1&), malIndex(lRowIndex), (mlIdxCount - plInsertAfter) * LenB(lRowIndex)
    End If
  ElseIf plInsertBefore Then
    lRowIndex = plInsertBefore
    If mlIdxCount Then
      CopyMemory malIndex(lRowIndex + 1&), malIndex(lRowIndex), (mlIdxCount - lRowIndex + 1&) * LenB(lRowIndex)
    End If
  Else
    'Simply insert the new element at the end of the index array
    lRowIndex = mlIdxCount + 1&
  End If
  mlIdxCount = mlIdxCount + 1&
  malIndex(lRowIndex) = lDataIndex
  'copy row data
  For lCol = 1& To mlColCount
    If Len(matColDef(lCol).sColName) Then
      If prowToAdd.ColExists(matColDef(lCol).sColName) Then
        If Not IsObject(prowToAdd(matColDef(lCol).sColName)) Then
          DataArray(lCol - 1&, lDataIndex) = prowToAdd(matColDef(lCol).sColName)
        Else
          Set DataArray(lCol - 1&, lDataIndex) = prowToAdd(matColDef(lCol).sColName)
        End If
      Else
        DataArray(lCol - 1&, lDataIndex) = Null
      End If
    Else
      DataArray(lCol - 1&, lDataIndex) = Null
    End If
  Next lCol
  AddRow = lRowIndex
End Function

'Return the index in DataArray() for the new row copy it
Private Function AllocRow(ByRef prowNew As CRow) As Long
  Dim lRetIndex   As Long
  
  lRetIndex = GarbQPop()
  If lRetIndex = -1& Then
    'No free slot in the garbage queue, add a new row in the DataArray
    If mlDataArraySlotCount = mlDataArraySize Then
      If mlColCount Then
        ReDim Preserve DataArray(0 To mlColCount - 1, 0 To mlDataArraySize + mlGrowSize - 1&)
      Else
        ReDim DataArray(0 To mlColCount - 1, 0 To mlDataArraySize + mlGrowSize - 1&)
      End If
      mlDataArraySize = mlDataArraySize + mlGrowSize
    End If
    lRetIndex = mlDataArraySlotCount
    mlDataArraySlotCount = mlDataArraySlotCount + 1&
  End If
  
  AllocRow = lRetIndex
End Function

'Return the index in DataArray() for a new row of data
Private Function AllocRowSlot() As Long
  Dim lRetIndex   As Long
  Dim lCol        As Long
  
  lRetIndex = GarbQPop()
  If lRetIndex = -1& Then
    'No free slot in the garbage queue, add a new row in the DataArray
    If mlDataArraySlotCount = mlDataArraySize Then
      If mlColCount Then
        ReDim Preserve DataArray(0 To mlColCount - 1, 0 To mlDataArraySize + mlGrowSize - 1&)
      Else
        ReDim DataArray(0 To mlColCount - 1, 0 To mlDataArraySize + mlGrowSize - 1&)
      End If
      mlDataArraySize = mlDataArraySize + mlGrowSize
    End If
    lRetIndex = mlDataArraySlotCount
    mlDataArraySlotCount = mlDataArraySlotCount + 1&
  End If
  
  AllocRowSlot = lRetIndex
End Function

Public Function AddValues(ParamArray pavValues() As Variant) As Long
  Dim lLB           As Long
  Dim lUB           As Long
  Dim lIndex        As Long
  Dim lListCol      As Long
  Dim lDataIndex    As Long
  Dim lRowIndex     As Long
  
  'Grow index array if necessary
  If mlIdxCount = mlIdxArraySize Then
    ReDim Preserve malIndex(1 To mlIdxArraySize + mlGrowSize)
    mlIdxArraySize = mlIdxArraySize + mlGrowSize
  End If
  'Find a slot where to insert the new row data
  lDataIndex = AllocRowSlot()
  'Adding a row breaks sorted state
  SetUnsorted
  'Simply insert the new element at the end of the index array
  lRowIndex = mlIdxCount + 1&
  mlIdxCount = mlIdxCount + 1&
  malIndex(lRowIndex) = lDataIndex
  'Assign passed variant array elements
  lLB = LBound(pavValues)
  lUB = UBound(pavValues)
  lListCol = 1&
  For lIndex = lLB To lUB
    If Not IsObject(pavValues(lIndex)) Then
      DataArray(malColIndex(lListCol) - 1, lDataIndex) = pavValues(lIndex)
    Else
      Set DataArray(malColIndex(lListCol) - 1, lDataIndex) = pavValues(lIndex)
    End If
    lListCol = lListCol + 1&
  Next lIndex
  AddValues = lRowIndex
End Function

Public Function AddValuesArray(pavValues As Variant, _
                               Optional ByVal plInsertAfter As Long = 0&, _
                               Optional ByVal plInsertBefore As Long = 0&) As Long
  Dim lLB           As Long
  Dim lUB           As Long
  Dim lIndex        As Long
  Dim lListCol      As Long
  Dim lDataIndex    As Long
  Dim lRowIndex     As Long
  
  'Grow index array if necessary
  If mlIdxCount = mlIdxArraySize Then
    ReDim Preserve malIndex(1 To mlIdxArraySize + mlGrowSize)
    mlIdxArraySize = mlIdxArraySize + mlGrowSize
  End If
  'Find a slot where to insert the new row data
  lDataIndex = AllocRowSlot()
  'Adding a row breaks sorted state
  SetUnsorted
  If plInsertAfter Then
    lRowIndex = plInsertAfter + 1&
    'Push down other index pointers
    If lRowIndex < mlIdxCount Then
      CopyMemory malIndex(lRowIndex + 1&), malIndex(lRowIndex), (mlIdxCount - plInsertAfter) * LenB(lRowIndex)
    End If
  ElseIf plInsertBefore Then
    lRowIndex = plInsertBefore
    If mlIdxCount Then
      CopyMemory malIndex(lRowIndex + 1&), malIndex(lRowIndex), (mlIdxCount - lRowIndex + 1&) * LenB(lRowIndex)
    End If
  Else
    'Simply insert the new element at the end of the index array
    lRowIndex = mlIdxCount + 1&
  End If
  mlIdxCount = mlIdxCount + 1&
  malIndex(lRowIndex) = lDataIndex
  'Assign passed variant array elements
  lLB = LBound(pavValues)
  lUB = UBound(pavValues)
  lListCol = 1&
  For lIndex = lLB To lUB
    If Not IsObject(pavValues(lIndex)) Then
      DataArray(malColIndex(lListCol) - 1, lDataIndex) = pavValues(lIndex)
    Else
      Set DataArray(malColIndex(lListCol) - 1, lDataIndex) = pavValues(lIndex)
    End If
    lListCol = lListCol + 1&
  Next lIndex
  AddValuesArray = lRowIndex
End Function

'
' Update row data
'

Public Sub AssignRow(ByVal plRowIndex As Long, ByRef prowSource As CRow)
  Dim lDataIndex    As Long
  Dim lCol          As Long
  Dim sColName      As String
  
  lDataIndex = malIndex(plRowIndex)
  'copy row data
  For lCol = 1& To mlColCount
    sColName = matColDef(lCol).sColName
    If Len(sColName) Then
      If prowSource.ColExists(sColName) Then
        If Not IsObject(prowSource(sColName)) Then
          DataArray(lCol - 1&, lDataIndex) = prowSource(sColName)
        Else
          Set DataArray(lCol - 1&, lDataIndex) = prowSource(sColName)
        End If
        'If the sorted column value is changed, then the sorted state is broken
        If IsSortColumn(lCol) And mfSorted Then SetUnsorted
      End If
    End If
  Next lCol
End Sub

Public Sub AssignValues(ByVal plRowIndex As Long, ParamArray pavValues() As Variant)
  Dim lLB           As Long
  Dim lUB           As Long
  Dim lIndex        As Long
  Dim lListCol      As Long
  Dim lDataIndex    As Long
  
  lDataIndex = malIndex(plRowIndex)
  'Assign passed variant array elements
  lLB = LBound(pavValues)
  lUB = UBound(pavValues)
  lListCol = 1&
  For lIndex = lLB To lUB
    If Not IsObject(pavValues(lIndex)) Then
      DataArray(malColIndex(lListCol) - 1, lDataIndex) = pavValues(lIndex)
    Else
      Set DataArray(malColIndex(lListCol) - 1, lDataIndex) = pavValues(lIndex)
    End If
    'If the sorted column value is changed, then the sorted state is broken
    If IsSortColumn(lListCol) And mfSorted Then SetUnsorted
    lListCol = lListCol + 1&
  Next lIndex
End Sub

Public Sub AssignValuesArray(ByVal plRowIndex As Long, ByRef pavValues() As Variant)
  Dim lLB           As Long
  Dim lUB           As Long
  Dim lIndex        As Long
  Dim lListCol      As Long
  Dim lDataIndex    As Long
  
  lDataIndex = malIndex(plRowIndex)
  'Assign passed variant array elements
  lLB = LBound(pavValues)
  lUB = UBound(pavValues)
  lListCol = 1&
  For lIndex = lLB To lUB
    If Not IsObject(pavValues(lIndex)) Then
      DataArray(malColIndex(lListCol) - 1, lDataIndex) = pavValues(lIndex)
    Else
      Set DataArray(malColIndex(lListCol) - 1, lDataIndex) = pavValues(lIndex)
    End If
    'If the sorted column value is changed, then the sorted state is broken
    If IsSortColumn(lListCol) And mfSorted Then SetUnsorted
    lListCol = lListCol + 1&
  Next lIndex
End Sub

'
' Remove rows
'

Public Sub Remove(ByVal plIndex As Long)
  Dim lSlot       As Long
  Dim lCol          As Long

  'Remove from the index array
  lSlot = malIndex(plIndex)
  
  If plIndex < mlIdxCount Then
    CopyMemory malIndex(plIndex), malIndex(plIndex + 1&), (mlIdxCount - 1&) * 4&
  End If
  mlIdxCount = mlIdxCount - 1&
  
  'release any hold object references
  For lCol = 0& To mlColCount - 1&
    If IsObject(DataArray(lCol, lSlot)) Then
      Set DataArray(lCol, lSlot) = Nothing
    End If
  Next lCol
  
  GarbQPush lSlot
End Sub

Private Sub FreeString(ByVal plSlot As Long)
  GarbQPush plSlot
End Sub

'
' Find data
'

Private Sub ClearFindInfo()
  With mtFindInfo
    .fUseSortKeys = False
    If .lSearchFieldsCt Then
      Erase .asSearchField
      Erase .alColIndex
      Erase .avSearchValue
      Erase .alFlags
    End If
    .lSearchFieldsCt = 0&
    .eSearchAlgo = eSearchAlgoSequential 'this one can always be applied
  End With
End Sub

'Warning, parameters may be 0 based arrays !
Private Sub ParseSearchColumns(ByRef pvSearchColumns As Variant, ByRef pvSearchCriteria As Variant)
  Dim iDims           As Integer
  Dim avBounds        As Variant
  Dim i               As Integer
  Dim eCompMethod     As VbCompareMethod
  Dim lParamIndex     As Long
  
  If moColIndexMap.CaseSensitive Then
    eCompMethod = vbBinaryCompare
  Else
    eCompMethod = vbTextCompare
  End If
  With mtFindInfo
    'If pvSearchColumns is an array, then we search on multiple columns
    GetVarArrayBounds pvSearchColumns, iDims, avBounds
    If iDims Then
      lParamIndex = avBounds(0)
      .lSearchFieldsCt = avBounds(1) - avBounds(0) + 1&
      ReDim .asSearchField(1 To .lSearchFieldsCt)
      ReDim .alColIndex(1 To .lSearchFieldsCt)
      ReDim .alFlags(1 To .lSearchFieldsCt)
      .asSearchField(1&) = pvSearchColumns(lParamIndex)
    Else
      ReDim .asSearchField(1 To 1)
      ReDim .alColIndex(1 To 1)
      ReDim .alFlags(1 To 1)
      .lSearchFieldsCt = 1&
      .asSearchField(1&) = pvSearchColumns
      lParamIndex = 1&
    End If
    'transform field names and compute derived infos
    .fUseSortKeys = True
    For i = 1& To .lSearchFieldsCt
      If i > 1& Then .asSearchField(i) = pvSearchColumns(lParamIndex)
      If left$(.asSearchField(i), 1) = "!" Then
        .alFlags(i) = .alFlags(i) Or FIF_CASESENSITIVE
        .asSearchField(i) = Right$(.asSearchField(i), Len(.asSearchField(i)) - 1&)
      Else
        .alFlags(i) = .alFlags(i) And (Not FIF_CASESENSITIVE)
      End If
      'see if we're in the sorted key fields sequence
      If .fUseSortKeys And (i <= miSortColCt) Then
        .fUseSortKeys = CBool(StrComp(.asSearchField(i), matSortInfo(i).sColName, eCompMethod) = 0)
        If .fUseSortKeys Then
          If matSortInfo(i).fDesc Then
            .alFlags(i) = .alFlags(i) Or FIF_SORTEDDESC
          Else
            .alFlags(i) = .alFlags(i) And (Not FIF_SORTEDDESC)
          End If
        Else
          .alFlags(i) = .alFlags(i) And (Not FIF_SORTEDDESC)
        End If
      Else
        .fUseSortKeys = False
      End If
      On Error Resume Next
      'We just test for a valid column name by computing its index
      .alColIndex(i) = ColPos(.asSearchField(i))
      If .alColIndex(i) = 0& Then
        'We can't accept a non matching column; we reuse VB error 5 (invalid proc call)
        Err.Raise 5&, "CList::ParseSearchColumns", .asSearchField(i) & " is not a valid column name"
      End If
      On Error GoTo 0
      lParamIndex = lParamIndex + 1&
    Next i
    'We must have a corresponding number of values
    GetVarArrayBounds pvSearchCriteria, iDims, avBounds
    If iDims Then
      lParamIndex = avBounds(0)
      If (avBounds(1) - avBounds(0) + 1&) < .lSearchFieldsCt Then
        'there is a problem, there are less search values than
        'there are search fields. We just decrement the number
        'of search fields we take into account, to synchronize.
        .lSearchFieldsCt = iDims
      End If
      ReDim .avSearchValue(1 To .lSearchFieldsCt)
      .avSearchValue(1&) = pvSearchCriteria(lParamIndex)
    Else
      ReDim .avSearchValue(1 To 1)
      .avSearchValue(1&) = pvSearchCriteria
      lParamIndex = 1&
    End If
    'interpret field values
    For i = 1& To .lSearchFieldsCt
      If i > 1& Then .avSearchValue(i) = pvSearchCriteria(lParamIndex)
      'Compute flags
      If VarType(.avSearchValue(i)) = vbString Then
        If CBool(InStr(1, .avSearchValue(i), "?")) Then
          .alFlags(i) = .alFlags(i) Or FIF_HASJOKERS
          'jokers/wildcards supported only on last search column
          If i < .lSearchFieldsCt Then .fUseSortKeys = False
        End If
        If CBool(InStr(1, .avSearchValue(i), "*")) Then
          .alFlags(i) = .alFlags(i) Or FIF_HASWILDCARDS
          'jokers/wildcards supported only on last search column
          If i < .lSearchFieldsCt Then .fUseSortKeys = False
        End If
        If (.alFlags(i) And FIF_HASJOKERS) = 0& Then
          If (.alFlags(i) And FIF_HASWILDCARDS) = 0& Then
            .alFlags(i) = .alFlags(i) Or FIF_EXACTSEARCH
          Else
            If left$(.avSearchValue(i), 1) <> "*" Then
              .alFlags(i) = .alFlags(i) Or FIF_ROOTSEARCH
              .avSearchValue(i) = left$(.avSearchValue(i), Len(.avSearchValue(i)) - 1&)
            End If
          End If
        End If
      End If
      lParamIndex = lParamIndex + 1&
    Next i
  End With 'mtFindInfo
End Sub

'Find a specific value and return its index. Returns 0 if not found.
'Note: when there are duplicates, any one of the duplicate's
'index may be returned. To get the first row of the list, when
'there are duplicates, use the FindFirst method.
Public Function Find(ByVal pvSearchColumns As Variant, ByVal pvSearchCriteria As Variant, Optional ByVal plStartFrom As Long = 1&, Optional ByVal plEndTo As Long = 0&) As Long
  Dim lMidIndex       As Long
  Dim lMinIndex       As Long
  Dim lMaxIndex       As Long
  Dim iComp           As Integer
  
  'If no items in array then immediately exit
  If mlIdxCount = 0& Then Exit Function
  'Fix last index to search to
  If plEndTo = 0& Then plEndTo = mlIdxCount
  'build find info
  ClearFindInfo
  ParseSearchColumns pvSearchColumns, pvSearchCriteria
  
  'use a sequential search if we can't benefit from the sorted state
  If (Not mfSorted) Or (Not mtFindInfo.fUseSortKeys) Then
    For lMidIndex = plStartFrom To plEndTo
      iComp = CompareRowForFind(malIndex(lMidIndex))
      If iComp = 0 Then
        Find = lMidIndex  'Item found, exit
        Exit Function
      End If
    Next lMidIndex
    'not found sequentially, the function terminates here.
    Exit Function
  End If
  
  'dichotomic search is possible
  lMinIndex = plStartFrom
  lMaxIndex = plEndTo
  While True
    lMidIndex = (lMinIndex + lMaxIndex) \ 2&
    If lMaxIndex < lMinIndex Then Exit Function
    iComp = CompareRowForFind(malIndex(lMidIndex))
    If iComp = 1 Then
      lMinIndex = lMidIndex + 1&
    Else
      If iComp = -1 Then
        lMaxIndex = lMidIndex - 1&
      Else
        Find = lMidIndex
        Exit Function
      End If
    End If
  Wend
End Function

'Find the first string and return its index. Returns 0 if not found.
Public Function FindFirst(ByVal pvSearchColumns As Variant, ByVal pvSearchCriteria As Variant, Optional ByVal plStartFrom As Long = 1&, Optional ByVal plEndTo As Long = 0&) As Long
  Dim lMidIndex       As Long
  Dim lMinIndex       As Long
  Dim lMaxIndex       As Long
  Dim lSaveIndex      As Long
  Dim iComp           As Integer
  
  'If no items in array then immediately exit
  If mlIdxCount = 0& Then Exit Function
  'Fix last index to search to
  If plEndTo = 0& Then plEndTo = mlIdxCount
  'build find info
  ClearFindInfo
  ParseSearchColumns pvSearchColumns, pvSearchCriteria
  
  'use a sequential search if we can't benefit from the sorted state
  If (Not mfSorted) Or (Not mtFindInfo.fUseSortKeys) Then
    For lMidIndex = plStartFrom To plEndTo
      iComp = CompareRowForFind(malIndex(lMidIndex))
      If iComp = 0 Then
        FindFirst = lMidIndex  'Item found, exit
        Exit Function
      End If
    Next lMidIndex
    'not found sequentially, the function terminates here.
    Exit Function
  End If
  
  'dichotomic search is possible
  lMinIndex = plStartFrom
  lMaxIndex = plEndTo
  While True
    lMidIndex = (lMinIndex + lMaxIndex) \ 2&
    If lMaxIndex < lMinIndex Then Exit Function
    iComp = CompareRowForFind(malIndex(lMidIndex))
    If iComp = 1 Then
      lMinIndex = lMidIndex + 1&
    Else
      If iComp = -1 Then
        lMaxIndex = lMidIndex - 1&
      Else
        If iComp = 1 Then
          lMaxIndex = lMidIndex - 1&
        Else
          If iComp = 0 Then
            'We've found a corresponding string. Now we bubble up
            'sequentially, until we reach the first one of its
            'duplicates (if any).
            Do While lMidIndex > 1&
              lSaveIndex = lMidIndex
              lMidIndex = lMidIndex - 1&
              iComp = CompareRowForFind(malIndex(lMidIndex))
              If iComp Then
                lMidIndex = lSaveIndex
                Exit Do
              End If
            Loop
            FindFirst = lMidIndex
            Exit Function
          Else
            lMinIndex = lMidIndex + 1&
          End If
        End If
      End If
    End If
  Wend
End Function

'The array must be sorted.
Public Sub RemoveDuplicates()
  If Not mfSorted Then
    'Raise "Invalid procedure call" standard vb error
    Err.Raise 5&, "CList::RemoveDuplicates", VBA.Error$(5&) & ". List must be sorted."
    Exit Sub
  End If
  
  'To have a duplicate we must have at least two elements...
  If mlIdxCount < 2& Then Exit Sub
  
  'First pass: identify duplicates and put a -1 index in their slot pointer (sentinels)
  Dim lIndex      As Long
  Dim vCompValue  As Variant
  Dim vCellValue  As Variant
  Dim iComp       As Integer
  Dim lNewIndex   As Long
  Dim lRefIndex   As Long
  Dim lCol        As Long
  
  lRefIndex = 1&
  For lIndex = 2& To mlIdxCount
    iComp = CompareRowsForSort(malIndex(lRefIndex), malIndex(lIndex))
    If iComp Then
      lRefIndex = lIndex
    Else
      'release any hold object references
      For lCol = 0& To mlColCount - 1&
        If IsObject(DataArray(lCol, lIndex)) Then
          Set DataArray(lCol, lIndex) = Nothing
        End If
      Next lCol
      'duplicate: free string and zero its index
      GarbQPush lIndex
      malIndex(lIndex) = -1&
    End If
  Next lIndex
  
  'second pass: compress the array and free string duplicates
  lNewIndex = 0&
  For lIndex = 1& To mlIdxCount
    If malIndex(lIndex) > -1& Then
      lNewIndex = lNewIndex + 1&
      malIndex(lNewIndex) = malIndex(lIndex)
    End If
  Next lIndex
  mlIdxCount = lNewIndex
End Sub

Public Property Get Row(ByVal plRowIndex As Long) As CRow
  Dim oRetRow           As CRow
  Set oRetRow = New CRow
  GetRow oRetRow, plRowIndex
  Set Row = oRetRow
  Set oRetRow = Nothing
End Property

Public Sub GetRow(prowNewRet As CRow, ByVal plRowIndex As Long)
  Dim lDataColIndex     As Long
  Dim lDataRowIndex     As Long
  Dim lColIndex         As Long
  Dim fValidRow         As Boolean
  lDataRowIndex = malIndex(plRowIndex)
  fValidRow = ((plRowIndex > 0&) And (plRowIndex <= mlIdxCount))
  
  Set prowNewRet = New CRow
  'Define row columns
  ReDim avColName(1 To mlColCount)
  ReDim avColType(1 To mlColCount)
  ReDim avColSize(1 To mlColCount)
  ReDim avColFlags(1 To mlColCount)
  ReDim avValues(1 To mlColCount)
  For lColIndex = 1& To mlColCount
    lDataColIndex = malColIndex(lColIndex)
    With matColDef(lDataColIndex)
      avColName(lColIndex) = .sColName
      avColType(lColIndex) = .iDataType
      avColSize(lColIndex) = .lDataSize
      avColFlags(lColIndex) = .lFlags
      avValues(lColIndex) = Null
      If fValidRow Then
        If Not IsObject(DataArray(lDataColIndex - 1&, lDataRowIndex)) Then
          avValues(lColIndex) = DataArray(lDataColIndex - 1&, lDataRowIndex)
        Else
          Set avValues(lColIndex) = DataArray(lDataColIndex - 1&, lDataRowIndex)
        End If
      End If
    End With
  Next lColIndex
  prowNewRet.ArrayDefine avColName, avColType, avColSize, avColFlags
  
  If fValidRow Then
    prowNewRet.ArrayAssign avValues
  End If
End Sub

Public Sub DefineRow(prowRet As CRow)
  prowRet.Clear
  If mlColCount = 0& Then Exit Sub
  
  Dim lDataColIndex     As Long
  Dim lDataRowIndex     As Long
  Dim lColIndex         As Long
  Dim fValidRow         As Boolean
  
  'Define row columns
  ReDim avColName(1 To mlColCount)
  ReDim avColType(1 To mlColCount)
  ReDim avColSize(1 To mlColCount)
  ReDim avColFlags(1 To mlColCount)
  ReDim avValues(1 To mlColCount)
  For lColIndex = 1& To mlColCount
    lDataColIndex = malColIndex(lColIndex)
    With matColDef(lDataColIndex)
      avColName(lColIndex) = .sColName
      avColType(lColIndex) = .iDataType
      avColSize(lColIndex) = .lDataSize
      avColFlags(lColIndex) = .lFlags And &HFFFF&
      avValues(lColIndex) = Null
    End With
  Next lColIndex
  prowRet.ArrayDefine avColName, avColType, avColSize, avColFlags
End Sub

Public Sub DefineList(plstRet As CList)
  plstRet.Clear
  If mlColCount = 0& Then Exit Sub
  
  Dim lDataColIndex     As Long
  Dim lDataRowIndex     As Long
  Dim lColIndex         As Long
  Dim fValidRow         As Boolean
  
  'Define row columns
  ReDim avColName(1 To mlColCount)
  ReDim avColType(1 To mlColCount)
  ReDim avColSize(1 To mlColCount)
  ReDim avColFlags(1 To mlColCount)
  For lColIndex = 1& To mlColCount
    lDataColIndex = malColIndex(lColIndex)
    With matColDef(lDataColIndex)
      avColName(lColIndex) = .sColName
      avColType(lColIndex) = .iDataType
      avColSize(lColIndex) = .lDataSize
      avColFlags(lColIndex) = .lFlags And &HFFFF&
    End With
  Next lColIndex
  plstRet.ArrayDefine avColName, avColType, avColSize, avColFlags
End Sub

Public Sub CopyFrom(ByRef plstSource As CList)
  Dim lColCount     As Long
  Dim iCol          As Long
  Dim lRow          As Long
  
  Clear
  lColCount = plstSource.ColCount
  If lColCount = 0& Then Exit Sub
  
  For iCol = 1& To lColCount
    With plstSource
      'vTemplateValue cannot be provided, so...
      Me.AddCol .ColName(iCol), _
                "", _
                .ColSize(iCol), _
                .ColFlags(iCol)
    End With
  Next iCol
  '...column types are synchronized afterwards
  For iCol = 1& To lColCount
    Me.ColType(iCol) = plstSource.ColType(iCol)
  Next iCol
  'There's no other solution than copying row by row
  If plstSource.Count Then
    For lRow = 1 To plstSource.Count
      Me.AddRow plstSource.Row(lRow)
    Next lRow
    If Len(plstSource.SortColumns) Then SetSorted plstSource.SortColumns
  End If
End Sub

Public Property Let Row(ByVal plRowIndex As Long, ByRef prowValues As CRow)
  If (plRowIndex < 1&) Or (plRowIndex > mlIdxCount) Then
    Exit Property
  End If
  Dim lCol          As Long
  Dim lDataIndex    As Long
  Dim lColPos       As Long
  'copy row data
  On Error Resume Next
  lDataIndex = malIndex(plRowIndex)
  For lCol = 1 To prowValues.ColCount
    lColPos = 0&
    If Len(prowValues.ColName(lCol)) Then lColPos = ColPos(prowValues.ColName(lCol))
    If lColPos Then
      If Not IsObject(prowValues(lCol)) Then
        DataArray(lColPos - 1&, lDataIndex) = prowValues(lCol)
      Else
        Set DataArray(lColPos - 1&, lDataIndex) = prowValues(lCol)
      End If
    End If
  Next lCol
End Property

'
' Garbage queue
'

Private Function GarbQPop() As Long
  If mlGarbQCount Then
    GarbQPop = malGarbageQ(mlGarbQHead)
    If mlGarbQHead < mlGarbQSize Then
      mlGarbQHead = mlGarbQHead + 1&
    Else
      mlGarbQHead = 1&
    End If
    mlGarbQCount = mlGarbQCount - 1&
  Else
    GarbQPop = -1&
  End If
End Function

Private Sub GarbQPush(ByVal plIndexValue As Long)
  If mlGarbQCount = mlGarbQSize Then
    Dim lMoveCount    As Long
    Dim lMoveIndex    As Long
    Dim lOldSize      As Long
    lOldSize = mlGarbQSize
    lMoveCount = mlGarbQSize - mlGarbQHead + 1
    'Grow the queue array
    ReDim Preserve malGarbageQ(1 To mlGarbQSize + mlGrowSize)
    mlGarbQSize = mlGarbQSize + mlGrowSize
    If mlGarbQTail < mlGarbQHead Then
      'Move at end of queue array
      'This throws a GPF: CopyMemory malGarbageQ(mlGarbQHead + mlGrowSize), malGarbageQ(mlGarbQHead), mlGrowSize * 4&
      For lMoveIndex = mlGarbQSize To mlGarbQSize - lMoveCount Step -1&
        malGarbageQ(lMoveIndex) = malGarbageQ(lMoveIndex - lOldSize)
      Next lMoveIndex
      mlGarbQHead = mlGarbQHead + mlGrowSize
    End If
  End If
  
  If mlGarbQTail < mlGarbQSize Then
    mlGarbQTail = mlGarbQTail + 1&
  Else
    mlGarbQTail = 1&
  End If
  malGarbageQ(mlGarbQTail) = plIndexValue
  mlGarbQCount = mlGarbQCount + 1&
End Sub

'
' Sorting
'

Private Sub QuickSortRows(ByVal plLowBound As Long, ByVal plUpBound As Long)
  Dim lPivot  As Long
  Dim lTemp   As Long
  Dim lFirst  As Long
  Dim lLast   As Long
  Dim lMiddle As Long
  Dim iComp   As Integer
  
  'Locate Pivot
  lFirst = plLowBound
  lLast = plUpBound
  lMiddle = (lFirst + lLast) / 2&
  lPivot = malIndex(lMiddle)
  
  Do  'Move pointers against each other
    iComp = CompareRowsForSort(malIndex(lFirst), lPivot)
    While iComp = -1
      lFirst = lFirst + 1&
      iComp = CompareRowsForSort(malIndex(lFirst), lPivot)
    Wend
    iComp = CompareRowsForSort(malIndex(lLast), lPivot)
    While iComp = 1
      lLast = lLast - 1&
      iComp = CompareRowsForSort(malIndex(lLast), lPivot)
    Wend
    
    If lFirst <= lLast Then
      'Swap string (faked) pointers
      lTemp = malIndex(lFirst)
      malIndex(lFirst) = malIndex(lLast)
      malIndex(lLast) = lTemp
      lFirst = lFirst + 1&
      lLast = lLast - 1&
    End If
  Loop Until lFirst > lLast
  
  If plLowBound < lLast Then
    Call QuickSortRows(plLowBound, lLast)
  End If
  If lFirst < plUpBound Then
    Call QuickSortRows(lFirst, plUpBound)
  End If
End Sub

Private Function CompNonObjectValues(ByRef pvValue1 As Variant, ByRef pvValue2 As Variant) As Integer
  If pvValue1 < pvValue2 Then
    CompNonObjectValues = -1
  ElseIf pvValue1 > pvValue2 Then
    CompNonObjectValues = 1
  Else
    CompNonObjectValues = 0
  End If
End Function

Private Function CompValues(ByRef pvValue1 As Variant, ByRef pvValue2 As Variant) As Integer
  If (Not IsObject(pvValue1)) And (Not IsObject(pvValue2)) Then
    CompValues = CompNonObjectValues(pvValue1, pvValue2)
  Else
    If Not miiCompObjects Is Nothing Then
      'it should be possible to compare an object with a value,
      'if we provide a comparer interface.
      CompValues = miiCompObjects.CompareValues(pvValue1, pvValue2)
    Else
      'We'll take the object address in memory as the value to compare
      If IsObject(pvValue1) Then
        CompValues = CompNonObjectValues(ObjPtr(pvValue1), pvValue2)
      Else
        CompValues = CompNonObjectValues(pvValue1, ObjPtr(pvValue2))
      End If
    End If
  End If
End Function

Private Function CompStrings(ByVal psString1 As String, ByVal psString2 As String, peCompMethod As VbCompareMethod) As Integer
  CompStrings = StrComp(psString1, psString2, peCompMethod)
End Function

'Used to sort rows. Returns -1, 0 or 1, as StrComp does
Private Function CompareRowsForSort(ByVal plRow1 As Long, ByVal plRow2 As Long) As Integer
  'we compare in the matSortInfo fields order.
  Dim iFld            As Long
  Dim lColIndex       As Long
  Dim iType           As String
  Dim iComp           As Integer
  Dim iDir            As Integer
  Dim eStrCompMethod  As VbCompareMethod
  
  'sequential search for this column value
  For iFld = 1 To miSortColCt
    If matSortInfo(iFld).fDesc Then iDir = -1 Else iDir = 1
    lColIndex = malColIndex(matSortInfo(iFld).lColIndex)
    iType = matColDef(lColIndex).iDataType
    lColIndex = lColIndex - 1&
    If iType = VbVarType.vbString Then
      If matSortInfo(iFld).fCaseSensitive Then
        eStrCompMethod = vbBinaryCompare
      Else
        eStrCompMethod = vbTextCompare
      End If
      iComp = CompStrings(DataArray(lColIndex, plRow1), DataArray(lColIndex, plRow2), eStrCompMethod)
    Else
      iComp = CompValues(DataArray(lColIndex, plRow1), DataArray(lColIndex, plRow2))
    End If
    'if there is a difference between comparison operators then we can stop comparing
    If iComp Then Exit For
  Next iFld
  CompareRowsForSort = iDir * iComp
End Function

Private Function CompareRowForFind(ByVal plRowIndex As Long) As Integer
  Dim iFld            As Long
  Dim lColIndex       As Long
  Dim iComp           As Integer
  Dim iDir            As Integer
  Dim eStrCompMethod  As VbCompareMethod
  Dim lLenSearch      As Long
  Dim vTemp           As Variant
  Dim iType           As Integer
  
  With mtFindInfo
    'We compare in the search fields order.
    'As soon as a field by field compare is <>0 then exit
    For iFld = 1& To .lSearchFieldsCt
      If .alFlags(iFld) And FIF_SORTEDDESC Then iDir = -1 Else iDir = 1
      lColIndex = malColIndex(.alColIndex(iFld))
      iType = matColDef(lColIndex).iDataType
      lColIndex = lColIndex - 1&
      vTemp = DataArray(lColIndex, plRowIndex)
      If iType = VbVarType.vbString Then
        'compare method for this search field derived from FIF_CASESENSITIVE flag
        If .alFlags(iFld) And FIF_CASESENSITIVE Then
          eStrCompMethod = vbBinaryCompare
        Else
          eStrCompMethod = vbTextCompare
        End If
        If iFld = .lSearchFieldsCt Then 'only on last search field we look for joker/wildcards
          If .alFlags(iFld) And FIF_EXACTSEARCH Then
            'raw string compare
            iComp = CompStrings(.avSearchValue(iFld), vTemp & "", eStrCompMethod)
          ElseIf .alFlags(iFld) And FIF_ROOTSEARCH Then
            lLenSearch = Len(.avSearchValue(iFld)) - 1&
            iComp = StrComp(.avSearchValue(iFld), left$(vTemp & "", lLenSearch), eStrCompMethod)
          Else
            iComp = 1
            If StrMatch(vTemp, .avSearchValue(iFld), eStrCompMethod) Then iComp = 0
          End If
        Else
          'raw string compare
          iComp = CompStrings(.avSearchValue(iFld), vTemp & "", eStrCompMethod)
        End If
      Else
        iComp = CompValues(.avSearchValue(iFld), vTemp)
      End If
      If iComp Then Exit For
    Next iFld
  End With 'mtFindInfo
  CompareRowForFind = iDir * iComp
End Function

'
' A routine for comparing strings using simple pattern matching operators,
' like "*" for root or suffix expansion and "?" as character joker.
'
'Returns True if psToMatch matches the psModel pattern.
Private Function StrMatch(ByVal psToMatch As String, ByVal psModel As String, ByVal piCompMethod As VbCompareMethod) As Boolean
  'If no wildcard or joker then simply return StrComp
  If (InStr(1, psModel, "*") = 0) And (InStr(1, psModel, "?") = 0) Then
    StrMatch = CBool(StrComp(psModel, psToMatch, piCompMethod) = 0)
    Exit Function
  End If
  'If no model or nothing to match, no match
  If Len(psModel) = 0 Then Exit Function
  If Len(psToMatch) = 0 Then Exit Function
  'If model = "*" then match is ok
  If psModel = "*" Then StrMatch = True: Exit Function
    
  'Wildcard at start of string or at end. If anywhere in the middle, then ignore.
  Dim fRootExpand     As Boolean  'True if wildcard at end
  Dim fSuffixExpand   As Boolean  'True if wildcard at start
  Dim sExpandedModel  As String
  Dim sShrinkedModel  As String
  Dim iLgToMatch      As Integer
  Dim iLgModel        As Integer
  Dim iJoker          As Integer
    
  iLgToMatch = Len(psToMatch)
  iLgModel = Len(psModel)
  fRootExpand = (Right$(psModel, 1) = "*")
  fSuffixExpand = (left$(psModel, 1) = "*")
    
  'Different algorithms for both expansions or for only one.
  If fRootExpand And fSuffixExpand Then
    'Note: This would be the slowest case if we handled possible jokers in psModel.
    iLgModel = iLgModel - 2
    If iLgToMatch > iLgModel Then
      'Example: A call like StrMatch("something", "*th*", vbTextCompare) has been made.
      'In this implementation, jokers ("?") are not supported in this case.
      sShrinkedModel = Mid$(psModel, 2, iLgModel)  'ie: "th" in our example
      StrMatch = CBool(InStr(1, psToMatch, sShrinkedModel, piCompMethod))
    ElseIf iLgToMatch < iLgModel Then
      'Example: A call like StrMatch("text", "*textual*", vbTextCompare) has been made.
      StrMatch = False
    Else 'If iLgToMatch = iLgModel Then
      'Example: A call like StrMatch("something", "*something*", vbTextCompare) has been made.
      StrMatch = CBool(StrComp(Mid$(psModel, 2, iLgModel), psToMatch, piCompMethod) = 0)
    End If
  Else
    iLgModel = iLgModel - 1 'don't count the *.
    If fRootExpand Then
      'Example: A call like StrMatch("something", "some*", vbTextCompare) has been made.
      If iLgToMatch > iLgModel Then
        iJoker = InStr(1, sExpandedModel, "?", vbBinaryCompare)
        If iJoker Then
          'We expand the model with missing characters borrowed to psToMatch.
          sExpandedModel = left$(psModel, iLgModel) & Right$(psToMatch, iLgToMatch - iLgModel)
          'Now replace all jokers ("?") with corresponding char from psToMatch
          While iJoker
            Mid$(sExpandedModel, iJoker, 1) = Mid$(psToMatch, iJoker, 1)
            iJoker = InStr(iJoker + 1, sExpandedModel, "?", vbBinaryCompare)
          Wend
          'And finally compare
          StrMatch = CBool(StrComp(sExpandedModel, psToMatch, piCompMethod) = 0)
        Else
          StrMatch = CBool(StrComp(left$(psModel, iLgModel), left$(psToMatch, iLgModel), piCompMethod) = 0)
        End If
      ElseIf iLgToMatch < iLgModel Then
        StrMatch = False
      Else 'If iLgToMatch = iLgModel Then
        StrMatch = CBool(StrComp(left$(psModel, iLgModel), psToMatch, piCompMethod) = 0)
      End If
    Else 'can only be suffix expand now.
      'Example: A call like StrMatch("something", "*thing", vbTextCompare) has been made.
      If iLgToMatch > iLgModel Then
        'We expand the model with missing characters borrowed to psToMatch.
        sExpandedModel = left$(psToMatch, iLgToMatch - iLgModel) & Right$(psModel, iLgModel)
        'Now replace all jokers ("?") with corresponding char from psToMatch
        iJoker = InStr(1, sExpandedModel, "?")
        While iJoker
          Mid$(sExpandedModel, iJoker, 1) = Mid$(psToMatch, iJoker, 1)
          iJoker = InStr(iJoker + 1, sExpandedModel, "?", vbBinaryCompare)
        Wend
        'And finally compare
        StrMatch = CBool(StrComp(sExpandedModel, psToMatch, piCompMethod) = 0)
      ElseIf iLgToMatch < iLgModel Then
        StrMatch = False
      Else 'If iLgToMatch = iLgModel Then
        StrMatch = CBool(StrComp(left$(psModel, iLgModel), psToMatch, piCompMethod) = 0)
      End If
    End If
  End If
End Function

'
' Handle system column flags (not used outside the class)
'

Private Sub ClearAllSysColFlags()
  Dim i     As Long
  On Error Resume Next
  For i = 1& To mlColCount
    matColDef(malColIndex(i)).lFlags = matColDef(malColIndex(i)).lFlags And &HFFFF&
  Next i
End Sub

Private Sub SetAllSysColFlags(ByVal plSysFlags As Long)
  Dim i     As Long
  On Error Resume Next
  plSysFlags = plSysFlags And &HFFFF0000
  For i = 1& To mlColCount
    matColDef(malColIndex(i)).lFlags = matColDef(malColIndex(i)).lFlags Or plSysFlags
  Next i
End Sub

Private Sub SetSysColFlags(ByVal plColIndex As Long, ByVal plSysFlags As Long)
  plSysFlags = plSysFlags And &HFFFF0000
  matColDef(malColIndex(plColIndex)).lFlags = matColDef(malColIndex(plColIndex)).lFlags Or plSysFlags
End Sub

Private Sub ClearSysColFlags(ByVal plColIndex As Long)
  matColDef(malColIndex(plColIndex)).lFlags = matColDef(malColIndex(plColIndex)).lFlags And &HFFFF&
End Sub
