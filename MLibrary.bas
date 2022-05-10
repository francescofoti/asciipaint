Attribute VB_Name = "MLibrary"
Option Compare Database
Option Explicit

Private Const HDRINFO_APPINFO       As String = "appinfo"
Private Const HDRINFO_DATECREATED   As String = "datecreated"
Private Const HDRINFO_GUID          As String = "guid"
Private Const HDRINFO_AUTHOR        As String = "author"
Private Const HDRINFO_COPYRIGHT     As String = "copyright"
Private Const HDRINFO_DESCRIPTION   As String = "description"

Public Const BLOCKTYPE_CROW    As Integer = 1
Public Const BLOCKTYPE_BINARY  As Integer = 10

Public Const BLOCKNAME_LIBHEADER     As String = "\"
Public Const BLOCKNAME_LIBCUSTPROPS  As String = ".customproperties"

'Bloc attributes
Public Const BA_DELETED            As Long = 4&
Public Const BA_NONMOVEABLE        As Long = 8&
Public Const BA_NONDELETEABLE      As Long = 16&

Public Const HBLOCK_INVALID        As Long = 0&
Public Const HLIB_INVALID          As Long = 0&

Public Enum eBlockPointer
  eNext = 0
  ePrev = 1
End Enum

'Error context
Private mlErr       As Long
Private msErr       As String
Private msErrCtx    As String

Private Sub ClearErr()
  mlErr = 0&
  msErr = ""
End Sub

Private Sub SetErr(ByVal psErrCtx As String, ByVal plErr As Long, ByVal psErr As String)
  msErrCtx = psErrCtx
  mlErr = plErr
  msErr = psErr
End Sub

Public Function LibLastErr() As Long
  LibLastErr = mlErr
End Function

Public Function LibLastErrDesc() As String
  LibLastErrDesc = msErr
End Function

Public Function LibLastErrCtx() As String
  LibLastErrCtx = msErrCtx
End Function

Private Sub AdvanceInt(ByVal phLib As Long)
  Dim iDummy As Integer
  Get #phLib, , iDummy
End Sub

Private Sub AdvanceLong(ByVal phLib As Long)
  Dim lDummy As Long
  Get #phLib, , lDummy
End Sub

Public Function OpenLibraryFile(ByVal psLibraryFile As String, phRetLib As Integer, Optional ByVal pfForWrite As Boolean = False) As Boolean
  Const LOCAL_ERR_CTX As String = "OpenLibraryFile"
  On Error GoTo OpenLibraryFile_Err
  ClearErr
  
  Dim fOK       As Boolean
  Dim hFile     As Integer
  
  phRetLib = HLIB_INVALID
  
  hFile = FreeFile
  'Always open in RW or failure when reopening same file on different mode
  'If Not pfForWrite Then
  '  Open psLibraryFile For Binary Access Read Lock Read Write As #hFile
  'Else
    Open psLibraryFile For Binary Access Read Write Lock Read Write As #hFile
  'End If
  
  phRetLib = hFile
  fOK = True
  
OpenLibraryFile_Exit:
  OpenLibraryFile = fOK
  If Not fOK Then
    If hFile <> HLIB_INVALID Then
      On Error Resume Next
      Close #hFile
    End If
  End If
  Exit Function
OpenLibraryFile_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  fOK = False
  Resume OpenLibraryFile_Exit
  Resume
End Function

Public Function CloseLibrary(ByVal phLib As Integer) As Boolean
  On Error Resume Next
  If phLib > HLIB_INVALID Then
    Close #phLib
    CloseLibrary = True
  End If
End Function

Private Sub DefineHeaderRow(ByRef poRow As CRow)
  poRow.ArrayDefine _
    Array( _
      HDRINFO_APPINFO, HDRINFO_DATECREATED, HDRINFO_GUID, HDRINFO_AUTHOR, _
      HDRINFO_COPYRIGHT, HDRINFO_DESCRIPTION _
    ), _
    Array( _
      vbString, vbString, vbString, vbString, _
      vbString, vbString _
    )
End Sub

Public Function LibGetMaxBlockNameLen() As Integer
  LibGetMaxBlockNameLen = 250
End Function

Private Function PadBlockName(ByVal psBlockName As String) As String
  If Len(psBlockName) < LibGetMaxBlockNameLen() Then
    psBlockName = psBlockName & Space$(LibGetMaxBlockNameLen() - Len(psBlockName))
  End If
  PadBlockName = psBlockName
End Function

'*************************
'
' Wrapping Blocks
'
'*************************

Private Function CreateBlock( _
    ByVal phLib As Integer, _
    ByVal psBlockName As String, _
    ByVal phPrevBlock As Long, _
    ByVal phNextBlock As Long, _
    ByVal piBlockType As Integer, _
    ByVal plAttribs As Long, _
    Optional ByVal psTag As String = "" _
  ) As Long
  Const LOCAL_ERR_CTX As String = "CreateBlock"
  Dim lBlockSize  As Long
  Dim hNewBlock   As Long
  
  On Error GoTo CreateBlock_Err
  ClearErr
  hNewBlock = HBLOCK_INVALID
  
  If Len(psBlockName) > LibGetMaxBlockNameLen() Then
    SetErr LOCAL_ERR_CTX, -1&, "Block element name too long"
    GoTo CreateBlock_Exit
  End If
  psBlockName = PadBlockName(psBlockName)
  
  'old, INCORRECT: iBlockSize = 18& + Len(psBlockName) * 2 + Len(psTag) * 2 + 4 '2 is for the string length added by FilePutUnicodeString
  lBlockSize = 0 ' blocksize is the len of data after the header
  hNewBlock = Seek(phLib)
  Put #phLib, , piBlockType
  Put #phLib, , lBlockSize
  Put #phLib, , phNextBlock
  Put #phLib, , phPrevBlock
  Put #phLib, , plAttribs
  FilePutUnicodeString phLib, psBlockName, True
  FilePutUnicodeString phLib, psTag, True
  
CreateBlock_Exit:
  CreateBlock = hNewBlock
  Exit Function
CreateBlock_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  hNewBlock = HBLOCK_INVALID
  Resume CreateBlock_Exit
  Resume
End Function

Private Sub SkipBlockHeader(ByVal phLib As Integer)
  Dim lCurPos As Long
  Dim sTag    As String
  AdvanceInt phLib 'block type
  AdvanceLong phLib 'BlockSize
  AdvanceLong phLib 'hNextBlock
  AdvanceLong phLib 'hPrevBlock
  AdvanceLong phLib 'lAttribs
  'seek over blockname and tag
  lCurPos = Seek(phLib)
  Seek phLib, lCurPos + LibGetMaxBlockNameLen() * 2 + 2
  sTag = FileGetUnicodeString(phLib)
End Sub

Public Function LibWriteBlockPointers( _
    ByVal phLib As Integer, _
    ByVal phBlock As Long, _
    ByVal phPrevBlock As Long, _
    ByVal phNextBlock As Long _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "LibWriteBlockPointers"
  Dim lCurPos     As Long
  
  On Error GoTo LibWriteBlockPointers_Err
  ClearErr
  
  lCurPos = Seek(phLib)
  
  Seek #phLib, phBlock
  AdvanceInt phLib  'iBlockType
  AdvanceLong phLib 'lBlockSize
  Put #phLib, , phNextBlock
  Put #phLib, , phPrevBlock
  
  Seek #phLib, lCurPos
  LibWriteBlockPointers = True

LibWriteBlockPointers_Exit:
  Exit Function
LibWriteBlockPointers_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LibWriteBlockPointers_Exit
End Function

Public Function LibReadBlockPointers( _
    ByVal phLib As Integer, _
    ByVal phBlock As Long, _
    ByRef phRetPrevBlock As Long, _
    ByRef phRetNextBlock As Long _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "LibReadBlockPointers"
  Dim lCurPos     As Long
  Dim iBlockType  As Integer
  Dim lBlockSize  As Long
  
  On Error GoTo LibReadBlockPointers_Err
  ClearErr
  
  phRetPrevBlock = HBLOCK_INVALID
  phRetNextBlock = HBLOCK_INVALID
  lCurPos = Seek(phLib)
  
  Seek #phLib, phBlock
  AdvanceInt phLib  'iBlockType
  AdvanceLong phLib 'lBlockSize
  Get #phLib, , phRetNextBlock
  Get #phLib, , phRetPrevBlock
  
  Seek #phLib, lCurPos
  LibReadBlockPointers = True

LibReadBlockPointers_Exit:
  Exit Function
LibReadBlockPointers_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LibReadBlockPointers_Exit
End Function

Public Function LibReadBlockPointer( _
    ByVal phLib As Integer, _
    ByVal phBlock As Long, _
    ByRef phRethBlock As Long, _
    ByVal peBlockPointer As eBlockPointer _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "LibReadBlockPointer"
  Dim lCurPos     As Long
  Dim iBlockType  As Integer
  Dim lBlockSize  As Long
  
  On Error GoTo LibReadBlockPointer_Err
  ClearErr
  
  phRethBlock = HBLOCK_INVALID
  lCurPos = Seek(phLib)
  
  Seek #phLib, phBlock
  AdvanceInt phLib  'iBlockType
  AdvanceLong phLib 'lBlockSize
  If peBlockPointer = eNext Then
    Get #phLib, , phRethBlock
  Else
    AdvanceLong phLib
    Get #phLib, , phRethBlock
  End If
  
  Seek #phLib, lCurPos
  LibReadBlockPointer = True

LibReadBlockPointer_Exit:
  Exit Function
LibReadBlockPointer_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LibReadBlockPointer_Exit
End Function

Public Function LibWriteBlockAttributes( _
    ByVal phLib As Integer, _
    ByVal phBlock As Long, _
    ByVal plAttributes As Long _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "LibWriteBlockAttributes"
  Dim lCurPos     As Long
  Dim iBlockType  As Integer
  Dim lBlockSize  As Long
  Dim hBlock      As Long
  
  On Error GoTo LibWriteBlockAttributes_Err
  ClearErr
  
  lCurPos = Seek(phLib)
  
  Seek #phLib, phBlock
  AdvanceInt phLib  'iBlockType
  AdvanceLong phLib 'lBlockSize
  AdvanceLong phLib 'next
  AdvanceLong phLib 'prev
  Put #phLib, , plAttributes
  
  Seek #phLib, lCurPos
  LibWriteBlockAttributes = True

LibWriteBlockAttributes_Exit:
  Exit Function
LibWriteBlockAttributes_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LibWriteBlockAttributes_Exit
End Function

Public Function LibRenameBlock( _
    ByVal phLib As Integer, _
    ByVal phBlock As Long, _
    ByVal psNewName As String _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "LibRenameBlock"
  Dim lCurPos     As Long
  Dim iBlockType  As Integer
  Dim lBlockSize  As Long
  Dim hBlock      As Long
  
  On Error GoTo LibRenameBlock_Err
  ClearErr
  
  psNewName = Trim$(psNewName)
  If Len(psNewName) > LibGetMaxBlockNameLen() Then
    SetErr LOCAL_ERR_CTX, -1&, "Block element name too long"
    GoTo LibRenameBlock_Exit
  End If
  If (psNewName = BLOCKNAME_LIBHEADER) Or (psNewName = BLOCKNAME_LIBCUSTPROPS) Then
    SetErr LOCAL_ERR_CTX, -2&, "Invalid block name"
    GoTo LibRenameBlock_Exit
  End If
  
  lCurPos = Seek(phLib)
  
  Seek #phLib, phBlock
  AdvanceInt phLib  'iBlockType
  AdvanceLong phLib 'lBlockSize
  AdvanceLong phLib 'next
  AdvanceLong phLib 'prev
  AdvanceLong phLib 'attribs
  FilePutUnicodeString phLib, psNewName, True
  
  Seek #phLib, lCurPos
  LibRenameBlock = True

LibRenameBlock_Exit:
  Exit Function
LibRenameBlock_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LibRenameBlock_Exit
End Function

Public Function LibReadBlockAttributes( _
    ByVal phLib As Integer, _
    ByVal phBlock As Long, _
    ByRef plRetAttribs As Long _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "LibReadBlockAttributes"
  Dim lCurPos     As Long
  Dim iBlockType  As Integer
  Dim lBlockSize  As Long
  Dim hBlock      As Long
  
  On Error GoTo LibReadBlockAttributes_Err
  ClearErr
  plRetAttribs = 0&
  
  lCurPos = Seek(phLib)
  
  Seek #phLib, phBlock
  AdvanceInt phLib  'iBlockType
  AdvanceLong phLib 'lBlockSize
  AdvanceLong phLib 'next
  AdvanceLong phLib 'prev
  Get #phLib, , plRetAttribs
  
  Seek #phLib, lCurPos
  LibReadBlockAttributes = True

LibReadBlockAttributes_Exit:
  Exit Function
LibReadBlockAttributes_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LibReadBlockAttributes_Exit
End Function

Private Function WriteBlockClose(ByVal phLib As Integer, ByVal phBlock As Long, ByVal plDataStartPos As Long) As Boolean
  Const LOCAL_ERR_CTX As String = "WriteBlockClose"
  Dim lPos          As Long
  Dim lBlockSize    As Long
  Dim iBlockType    As Integer
  
  On Error GoTo WriteBlockClose_Err
  ClearErr
  
  lPos = Seek(phLib)
  lBlockSize = lPos - plDataStartPos
  
  Seek #phLib, phBlock
  AdvanceInt phLib  'iBlockType
  Put #phLib, , lBlockSize
  Seek #phLib, lPos
  
  WriteBlockClose = True

WriteBlockClose_Exit:
  Exit Function
WriteBlockClose_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume WriteBlockClose_Exit
End Function

Public Function LibReadBlockHeader( _
    ByVal phLib As Integer, _
    ByVal phBlock As Long, _
    ByRef piRetBlockType As Integer, _
    ByRef plRetBlockSize As Long, _
    ByRef phRetNextBlock As Long, _
    ByRef phRetPrevBlock As Long, _
    ByRef plRetAttribs As Long, _
    ByRef psBlockName As String, _
    ByRef psTag As String _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "LibReadBlockHeader"
  'NO, dont move back: Dim lCurPos     As Long
  Dim iBlockType  As Integer
  Dim lBlockSize  As Long
  
  On Error GoTo LibReadBlockHeader_Err
  ClearErr
  
  piRetBlockType = 0&
  plRetBlockSize = 0&
  phRetNextBlock = HBLOCK_INVALID
  phRetPrevBlock = HBLOCK_INVALID
  plRetAttribs = 0&
  psBlockName = ""
  psTag = ""
  
  'NO, dont move back: lCurPos = Seek(phLib)
  
  Seek #phLib, phBlock
  Get #phLib, , piRetBlockType
  Get #phLib, , plRetBlockSize
  Get #phLib, , phRetNextBlock
  Get #phLib, , phRetPrevBlock
  Get #phLib, , plRetAttribs
  psBlockName = Trim$(FileGetUnicodeString(phLib))
  psTag = FileGetUnicodeString(phLib)
  
  'NO, dont move back: Seek #phLib, lCurPos
  LibReadBlockHeader = True

LibReadBlockHeader_Exit:
  Exit Function
LibReadBlockHeader_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LibReadBlockHeader_Exit
End Function

Public Function LibIsValidBlockHandle(ByVal phLib As Integer, ByVal phBlock As Long) As Boolean
  On Error GoTo LibIsValidBlockHandle_Err
  LibIsValidBlockHandle = CBool((phBlock <> HBLOCK_INVALID) And (phBlock <= LOF(phLib)))
  Exit Function
LibIsValidBlockHandle_Err:
End Function

'*************************
'
' CRow blocks
'
'*************************

Public Function LibCreateCRowBlock( _
    ByVal phLib As Integer, _
    ByVal psBlockName As String, _
    ByRef prowData As CRow, _
    Optional ByVal plAttribs As Long = 0& _
  ) As Long
  Const LOCAL_ERR_CTX As String = "LibCreateCRowBlock"
  On Error GoTo LibCreateCRowBlock_Err
  ClearErr
  
  Dim hBlock  As Long
  Dim iPairCt As Integer
  Dim iPair   As Integer
  Dim fOK     As Boolean
  Dim lBlockDataStartPos As Long
  
  'go to end of lib file
  Seek #phLib, LOF(phLib) + 1&
  
  hBlock = CreateBlock(phLib, psBlockName, HBLOCK_INVALID, HBLOCK_INVALID, BLOCKTYPE_CROW, plAttribs)
  If hBlock = HBLOCK_INVALID Then
    GoTo LibCreateCRowBlock_Exit
  End If
  
  lBlockDataStartPos = Seek(phLib)
  iPairCt = prowData.ColCount
  Put #phLib, , iPairCt
  
  For iPair = 1 To iPairCt
    FilePutUnicodeString phLib, prowData.ColName(iPair)
    FilePutUnicodeString phLib, prowData.ColValue(iPair)
  Next iPair
  
  fOK = WriteBlockClose(phLib, hBlock, lBlockDataStartPos)
  If Not fOK Then
    hBlock = HBLOCK_INVALID
  End If
  
LibCreateCRowBlock_Exit:
  LibCreateCRowBlock = hBlock
  Exit Function

LibCreateCRowBlock_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LibCreateCRowBlock_Exit
  Resume
End Function

Public Function ReadCRowBlock( _
    ByVal phLib As Integer, _
    ByVal phBlock As Long, _
    ByRef poRetNewRow As CRow _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "ReadCRowBlock"
  On Error GoTo ReadCRowBlock_Err
  ClearErr
  
  Dim sValue        As String
  Dim rowData       As CRow
  Dim iKeyPairCt    As Integer
  Dim i             As Integer
  Dim sColName      As String
  Dim lCurPos       As Long
  
  Set poRetNewRow = Nothing
  Set rowData = New CRow
  
  lCurPos = Seek(phLib)
  Seek #phLib, phBlock
  SkipBlockHeader phLib
  
  Get #phLib, , iKeyPairCt
  For i = 1 To iKeyPairCt
    sColName = FileGetUnicodeString(phLib)
    sValue = FileGetUnicodeString(phLib)
    'If we have a duplicate key, we append the content
    'to the existing column
    If Not rowData.ColExists(sColName) Then
      rowData.AddCol sColName, sValue, Len(sValue), 0&
    Else
      rowData(sColName) = rowData(sColName) & sValue
    End If
  Next i
  
  Set poRetNewRow = rowData
  ReadCRowBlock = True
  
ReadCRowBlock_Exit:
  On Error Resume Next
  If lCurPos > 0& Then
    Seek #phLib, lCurPos
  End If
  Set rowData = Nothing
  Exit Function

ReadCRowBlock_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume ReadCRowBlock_Exit
  Resume
End Function

Public Function LibGotoBlock(ByVal phLib As Integer, ByVal phBlock As Long) As Long
  Const LOCAL_ERR_CTX As String = "LibGotoBlock"
  On Error GoTo LibGotoBlock_Err
  ClearErr
  
  Dim lCurPos As Long
  Seek #phLib, phBlock
  lCurPos = Seek(phLib)
  LibGotoBlock = lCurPos
  
LibGotoBlock_Exit:
  Exit Function

LibGotoBlock_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LibGotoBlock_Exit
  Resume
End Function

Public Function LibGotoNextBlock(ByVal phLib As Integer, ByVal phBlock As Long) As Long
  Const LOCAL_ERR_CTX As String = "LibGotoNextBlock"
  On Error GoTo LibGotoNextBlock_Err
  ClearErr
  
  Dim hNextBlock As Long
  Dim fOK As Boolean
  
  fOK = LibReadBlockPointer(phLib, phBlock, hNextBlock, eNext)
  If fOK Then
    LibGotoNextBlock = LibGotoBlock(phLib, hNextBlock)
  Else
    LibGotoNextBlock = HBLOCK_INVALID
  End If
  
LibGotoNextBlock_Exit:
  Exit Function

LibGotoNextBlock_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LibGotoNextBlock_Exit
  Resume
End Function

'*************************
'
' Library header
'
'*************************

Private Function LibWriteHeader( _
    ByVal phLib As Integer, _
    ByVal psAuthorName As String, _
    ByVal psCopyright As String, _
    ByVal psDescription As String, _
    ByRef prowCustomProps As CRow _
  ) As Long
  Const LOCAL_ERR_CTX As String = "LibWriteHeader"
  On Error GoTo LibWriteHeader_Err
  ClearErr
  
  Dim rowInfo           As CRow
  Dim hBlock            As Long
  Dim hBlock2           As Long
  Dim fOK               As Boolean
  
  Set rowInfo = New CRow
  DefineHeaderRow rowInfo
  rowInfo(HDRINFO_APPINFO) = APP_NAME & " v" & APP_VERSION
  rowInfo(HDRINFO_DATECREATED) = Format$(Now, "yyyymmddhh:mm:ss")
  rowInfo(HDRINFO_GUID) = CreateGUID()
  rowInfo(HDRINFO_AUTHOR) = psAuthorName
  rowInfo(HDRINFO_COPYRIGHT) = psCopyright
  rowInfo(HDRINFO_DESCRIPTION) = psDescription
  
  fOK = True
  hBlock = LibCreateCRowBlock(phLib, BLOCKNAME_LIBHEADER, rowInfo, BA_NONMOVEABLE Or BA_NONDELETEABLE)
  If hBlock > HBLOCK_INVALID Then
    'Link the two blocks
    If Not prowCustomProps Is Nothing Then
      hBlock2 = LibCreateCRowBlock(phLib, BLOCKNAME_LIBCUSTPROPS, prowCustomProps)
      fOK = LibWriteBlockPointers(phLib, hBlock2, hBlock, HBLOCK_INVALID)
    End If
    If fOK Then
      fOK = LibWriteBlockPointers(phLib, hBlock, HBLOCK_INVALID, hBlock2)
      If Not fOK Then
        hBlock = HBLOCK_INVALID
        GoTo LibWriteHeader_Exit
      End If
    Else
      hBlock = HBLOCK_INVALID
      GoTo LibWriteHeader_Exit
    End If
  End If
  
  LibWriteHeader = hBlock
  
LibWriteHeader_Exit:
  Exit Function

LibWriteHeader_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LibWriteHeader_Exit
End Function

Public Function LibReadHeader( _
    ByVal phLib As Integer, _
    ByRef poRetNewRow As CRow, _
    ByRef prowRetNewCustProps As CRow, _
    Optional ByVal pfReadCustProps As Boolean = True _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "LibReadHeader"
  On Error GoTo LibReadHeader_Err
  ClearErr
  
  Dim sValue      As String
  Dim rowData     As CRow
  Dim fOK         As Boolean
  Dim hNextBlock  As Long
  
  Set poRetNewRow = Nothing
  Set prowRetNewCustProps = Nothing
  
  'The lib header is the CRow block at position 1
  fOK = ReadCRowBlock(phLib, 1&, poRetNewRow)
  If pfReadCustProps Then
    fOK = LibReadBlockPointer(phLib, 1&, hNextBlock, eNext)
    If fOK Then
      fOK = ReadCRowBlock(phLib, hNextBlock, prowRetNewCustProps)
    End If
  End If
  LibReadHeader = fOK
  
LibReadHeader_Exit:
  Set rowData = Nothing
  Exit Function

LibReadHeader_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LibReadHeader_Exit
  Resume
End Function

Public Function LibCreateLibrary( _
    ByVal psLibraryFile As String, _
    ByVal psAuthorName As String, _
    ByVal psCopyright As String, _
    ByVal psDescription As String, _
    ByRef prowCustomProps As CRow _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "LibCreateLibrary"
  On Error GoTo LibCreateLibrary_Err
  ClearErr
  
  Dim fIsOpen     As Boolean
  Dim fOK         As Boolean
  Dim hLibFile    As Integer
  
  If ExistFile(psLibraryFile) Then
    SetErr LOCAL_ERR_CTX, -1&, "Library file already exists"
    GoTo LibCreateLibrary_Exit
  End If
  
  hLibFile = FreeFile
  Open psLibraryFile For Output Access Write Lock Read Write As #hLibFile
  Close hLibFile
  hLibFile = HLIB_INVALID
  
  fOK = OpenLibraryFile(psLibraryFile, hLibFile, True)
  If Not fOK Then
    GoTo LibCreateLibrary_Exit
  End If
  
  fOK = LibWriteHeader(hLibFile, psAuthorName, psCopyright, psDescription, prowCustomProps)
  
  CloseLibrary hLibFile
  hLibFile = HLIB_INVALID
  
  LibCreateLibrary = fOK
  
LibCreateLibrary_Exit:
  On Error Resume Next
  If hLibFile > HLIB_INVALID Then
    CloseLibrary hLibFile
  End If
  Exit Function

LibCreateLibrary_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LibCreateLibrary_Exit
  Resume
End Function

'*************************
'
' Add content
'
'*************************

Public Function IsValidBlockName(psBlockName) As Boolean
  IsValidBlockName = CBool( _
                      (Len(psBlockName) > 0) And _
                      (psBlockName <> BLOCKNAME_LIBHEADER) And _
                      (psBlockName <> BLOCKNAME_LIBCUSTPROPS) _
                     )
End Function

'returns hBlock
Public Function LibAddFile( _
    ByVal phLib As Long, _
    ByVal psBlockName As String, _
    ByVal psFilename As String, _
    ByRef prowCustomProps As CRow _
  ) As Long
  Const LOCAL_ERR_CTX As String = "LibAddFile"
  On Error GoTo LibAddFile_Err
  ClearErr
  
  Dim abData()    As Byte
  Dim hFile       As Integer
  Dim fIsOpen     As Boolean
  Dim fOK         As Boolean
  Dim hBlock      As Long
  Dim hBlock2     As Long
  Dim lBlockDataStartPos As Long
  
  If Not IsValidBlockName(psBlockName) Then
    SetErr LOCAL_ERR_CTX, -1&, "[" & psBlockName & "] is not a valid block name"
    LibAddFile = False
    GoTo LibAddFile_Exit
  End If
  
  hFile = FreeFile
  Open psFilename For Binary Access Read As #hFile
  fIsOpen = True
  
  'read all the bytes
  ReDim abData(1 To LOF(hFile)) As Byte
  Get #hFile, , abData
  Close hFile
  fIsOpen = False
  
  'go to end of lib file
  Seek #phLib, LOF(phLib) + 1&
  
  'create binary block
  hBlock = CreateBlock(phLib, psBlockName, HBLOCK_INVALID, HBLOCK_INVALID, BLOCKTYPE_BINARY, 0&)
  If hBlock <> HBLOCK_INVALID Then
    lBlockDataStartPos = Seek(phLib)
    Put #phLib, , abData
    fOK = WriteBlockClose(phLib, hBlock, lBlockDataStartPos)
    If fOK Then
      'write custom props and link the blocks
      If Not prowCustomProps Is Nothing Then
        hBlock2 = LibCreateCRowBlock(phLib, BLOCKNAME_LIBCUSTPROPS, prowCustomProps)
        fOK = LibWriteBlockPointers(phLib, hBlock2, hBlock, HBLOCK_INVALID)
      End If
      If fOK Then
        fOK = LibWriteBlockPointers(phLib, hBlock, HBLOCK_INVALID, hBlock2)
        If Not fOK Then
          hBlock = HBLOCK_INVALID
          GoTo LibAddFile_Exit
        End If
      Else
        hBlock = HBLOCK_INVALID
        GoTo LibAddFile_Exit
      End If
    Else
      LibAddFile = HBLOCK_INVALID
      GoTo LibAddFile_Exit
    End If
  Else
    LibAddFile = HBLOCK_INVALID
    GoTo LibAddFile_Exit
  End If
  
  LibAddFile = hBlock
  
LibAddFile_Exit:
  On Error Resume Next
  If fIsOpen Then
    Close #hFile
  End If
  Exit Function

LibAddFile_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LibAddFile_Exit
  Resume
End Function

Private Function ExportCRowBlock(ByVal phFile As Integer, ByRef prowData As CRow) As Boolean
  Const LOCAL_ERR_CTX As String = "ExportCRowBlock"
  On Error GoTo ExportCRowBlock_Err
  Dim fOK As Boolean
  Dim i   As Integer
  ClearErr
  
  If prowData Is Nothing Then
    fOK = True
    GoTo ExportCRowBlock_Exit
  End If
  
  For i = 1 To prowData.ColCount
    FilePutUnicodeString phFile, prowData.ColName(i) & "=" & prowData(i) & vbCrLf, False
  Next i
  
  fOK = True
  
ExportCRowBlock_Exit:
  ExportCRowBlock = fOK
  Exit Function

ExportCRowBlock_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume ExportCRowBlock_Exit
  Resume
End Function

'*************************
'
' Directory
'
'*************************

Private Sub DefineDirectoryList(plstDir As CList)
  plstDir.ArrayDefine Array( _
                      "hblock", _
                      "blocktype", _
                      "blocksize", _
                      "nextsibling", _
                      "prevsibling", _
                      "attribs", _
                      "blockname", _
                      "tag" _
                    ), Array( _
                      vbLong, _
                      vbInteger, _
                      vbLong, _
                      vbLong, _
                      vbLong, _
                      vbLong, _
                      vbString, _
                      vbString _
                    )
End Sub

Public Function LibLoadDirectory( _
    ByVal phLib As Integer, _
    ByRef plstRetNewDir As CList _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "LibLoadDirectory"
  On Error GoTo LibLoadDirectory_Err
  ClearErr
  
  Dim lstDir      As CList
  Dim fOK         As Boolean
  
  Dim rowHeader       As CRow
  Dim i               As Integer
  Dim rowCustProps    As New CRow
  
  Dim hBlock          As Long
  Dim iBlockType      As Integer
  Dim lBlockSize      As Long
  Dim hNextSibling    As Long
  Dim hPrevSibling    As Long
  Dim hNextBlock      As Long
  Dim lAttribs        As Long
  Dim sBlockName      As String
  Dim sTag            As String
  Dim iLinkedBlockCt  As Integer
  Dim iFind           As Long
  
  'Read library header (but not the custom props)
  fOK = LibReadHeader(phLib, rowHeader, Nothing, False)
  If Not fOK Then
    GoTo LibLoadDirectory_Exit
  End If
  
  Set lstDir = New CList
  DefineDirectoryList lstDir
  
  hBlock = 1&
  Do While LibIsValidBlockHandle(phLib, hBlock)
    iFind = lstDir.Find("hblock", hBlock)
    If iFind > 0 Then
      'If we're reading the same block twice, then there's a block chaining error somewhere
      'and we can only stop reading the affected library.
      SetErr LOCAL_ERR_CTX, -1&, "Reading same block #" & hBlock & " twice. Library integrity is broken"
      GoTo LibLoadDirectory_Exit
    End If
    
    fOK = LibReadBlockHeader(phLib, hBlock, iBlockType, lBlockSize, hNextSibling, hPrevSibling, lAttribs, sBlockName, sTag)
    If Not fOK Then
      GoTo LibLoadDirectory_Exit
    End If
    'Jump over block data
    hNextBlock = Seek(phLib) + lBlockSize
    
    'we list only "parent" blocks, blocks that not chained or first
    'of a chain, ie blcks that have no previous sibling.
    If hPrevSibling = HBLOCK_INVALID Then
      lstDir.AddValues hBlock, iBlockType, lBlockSize, hNextSibling, hPrevSibling, lAttribs, sBlockName, sTag
    End If
    
    hBlock = hNextBlock
  Loop
  
  If lstDir.Count > 1& Then
    lstDir.Sort "blockname"
  End If
  Set plstRetNewDir = lstDir
  LibLoadDirectory = True
  
LibLoadDirectory_Exit:
  Set lstDir = Nothing
  Exit Function

LibLoadDirectory_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LibLoadDirectory_Exit
  Resume
End Function

Public Function LibReadCustProps( _
    ByVal phLib As Integer, _
    ByVal phBlock As Long, _
    ByRef phRetNextPropBlock As Long, _
    ByRef prowRetNewCustProps As CRow _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "LibReadCustProps"
  On Error GoTo LibReadCustProps_Err
  ClearErr
  
  Dim hNextBlock    As Long
  Dim fOK           As Boolean
  Dim rowProps      As CRow
  
  phRetNextPropBlock = HBLOCK_INVALID
  Set prowRetNewCustProps = Nothing
  fOK = LibReadBlockPointer(phLib, phBlock, hNextBlock, eNext)
  If fOK Then
    If hNextBlock <> HBLOCK_INVALID Then
      fOK = LibReadBlockPointer(phLib, hNextBlock, phRetNextPropBlock, eNext)
      If Not fOK Then
        GoTo LibReadCustProps_Exit
      End If
      fOK = ReadCRowBlock(phLib, hNextBlock, rowProps)
      If fOK Then
        Set prowRetNewCustProps = rowProps
      End If
    End If
  End If
  
  LibReadCustProps = True
  
LibReadCustProps_Exit:
  Set rowProps = Nothing
  Exit Function

LibReadCustProps_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LibReadCustProps_Exit
  Resume
End Function

Public Function LibReadPropsBlock( _
    ByVal phLib As Integer, _
    ByVal phBlock As Long, _
    ByRef phRetNextPropBlock As Long, _
    ByRef prowRetNewProps As CRow _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "LibReadPropsBlock"
  On Error GoTo LibReadPropsBlock_Err
  ClearErr
  
  Dim fOK             As Boolean
  Dim rowProps        As CRow
  Dim iBlockType      As Integer
  Dim lBlockSize      As Long
  Dim hNextSibling    As Long
  Dim hPrevSibling    As Long
  Dim hNextBlock      As Long
  Dim lAttribs        As Long
  Dim sBlockName      As String
  Dim sTag            As String
  
  phRetNextPropBlock = HBLOCK_INVALID
  Set prowRetNewProps = Nothing
  fOK = LibReadBlockHeader(phLib, phBlock, iBlockType, lBlockSize, hNextSibling, hPrevSibling, lAttribs, sBlockName, sTag)
  If Not fOK Then
    GoTo LibReadPropsBlock_Exit
  End If
  fOK = ReadCRowBlock(phLib, phBlock, rowProps)
  If fOK Then
    phRetNextPropBlock = hNextSibling
    Set prowRetNewProps = rowProps
  End If
  
  LibReadPropsBlock = True
  
LibReadPropsBlock_Exit:
  Set rowProps = Nothing
  Exit Function

LibReadPropsBlock_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LibReadPropsBlock_Exit
  Resume
End Function

Public Function LibAttribsToText(ByVal plAttributes As Long) As String
  Dim sText As String
  
  If plAttributes And BA_NONMOVEABLE Then
    If Len(sText) > 0 Then
      sText = sText & ", "
    End If
    sText = sText & "fixed"
  End If
  
  If plAttributes And BA_NONDELETEABLE Then
    If Len(sText) > 0 Then
      sText = sText & ", "
    End If
    sText = sText & "permanent"
  End If
  
  If plAttributes And BA_DELETED Then
    If Len(sText) > 0 Then
      sText = sText & ", "
    End If
    sText = sText & "deleted"
  End If
  
  LibAttribsToText = sText
End Function

Public Function LibExtractFile( _
    ByVal phLib As Long, _
    ByVal phBlock As Long, _
    ByVal psTargetFilename As String _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "LibExtractFile"
  On Error GoTo LibExtractFile_Err
  ClearErr
  
  Dim abData()    As Byte
  Dim hFile       As Integer
  Dim fIsOpen     As Boolean
  Dim fOK         As Boolean
  Dim lstBlocks   As CList
  Dim i           As Integer
  Dim hBlock      As Long
  
  Dim iBlockType  As Integer
  Dim lBlockSize  As Long
  Dim hNextBlock  As Long
  Dim hPrevBlock  As Long
  Dim lAttribs    As Long
  Dim sBlockName  As String
  Dim sTag        As String
  Dim fNoData     As Boolean
  Dim rowProps    As CRow
  Dim iStartBlockType As Integer
  
  If ExistFile(psTargetFilename) Then
    SetErr LOCAL_ERR_CTX, -1&, "File already exists"
    Exit Function
  End If
  
  Set lstBlocks = New CList
  DefineDirectoryList lstBlocks
  
  'Create target file
  hFile = FreeFile
  Open psTargetFilename For Binary Access Read Write Lock Read Write As #hFile
  fIsOpen = True
  
  If iBlockType = BLOCKTYPE_CROW Then
    FilePutUnicodeString hFile, ";" & APP_NAME & " v" & APP_VERSION & " " & Format$(Now) & vbCrLf, False
  End If
  
  hBlock = phBlock
  Do
    fOK = LibReadBlockHeader(phLib, hBlock, iBlockType, lBlockSize, hNextBlock, hPrevBlock, lAttribs, sBlockName, sTag)
    If fOK Then
      lstBlocks.AddValues hBlock, iBlockType, lBlockSize, hNextBlock, hPrevBlock, lAttribs, sBlockName, sTag
      If iStartBlockType = 0 Then
        iStartBlockType = iBlockType
      End If
    Else
      GoTo LibExtractFile_Exit
    End If
    If lBlockSize <= 0& Then
      fNoData = True
    End If
    
    'write all the bytes
    If Not fNoData Then
      If iBlockType = BLOCKTYPE_CROW Then
        FilePutUnicodeString hFile, "[block " & hBlock & "]" & vbCrLf, False
        fOK = ReadCRowBlock(phLib, hBlock, rowProps)
        If Not fOK Then
          GoTo LibExtractFile_Exit
        End If
        fOK = ExportCRowBlock(hFile, rowProps)
      Else
        ReDim abData(1 To lBlockSize) As Byte
        Get #phLib, , abData
        Put #hFile, , abData
        fOK = True
        'don't read any other blocks (only for CRows)
        hNextBlock = HBLOCK_INVALID
      End If
    Else
      fOK = True
      If iBlockType <> BLOCKTYPE_CROW Then
        hNextBlock = HBLOCK_INVALID
      End If
    End If
    
    'avoid endless loop:
    If hNextBlock <> HBLOCK_INVALID Then
      i = lstBlocks.Find("hblock", hNextBlock)
      If i = 0 Then
        hBlock = hNextBlock
      Else
        FilePutUnicodeString hFile, ";ERROR: Circular link detected at block #" & hBlock & " referencing next block #" & hNextBlock & vbCrLf, False
        fOK = False
      End If
    End If
  Loop Until (hNextBlock = HBLOCK_INVALID) Or (Not fOK)
  
  If iStartBlockType = BLOCKTYPE_CROW Then
    FilePutUnicodeString hFile, vbCrLf & "[block list]" & vbCrLf, False
    FilePutUnicodeString hFile, "count=" & lstBlocks.Count & vbCrLf, False
    For i = 1 To lstBlocks.Count
      FilePutUnicodeString hFile, "name" & i & "=" & lstBlocks("blockname", i) & vbCrLf, False
      FilePutUnicodeString hFile, "blockid" & i & "=" & lstBlocks("hblock", i) & vbCrLf, False
    Next i
  End If
  
  Close hFile
  fIsOpen = False
  
  LibExtractFile = fOK
  
LibExtractFile_Exit:
  On Error Resume Next
  If fIsOpen Then
    Close #hFile
  End If
  Set lstBlocks = Nothing
  Exit Function

LibExtractFile_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LibExtractFile_Exit
  Resume
End Function


