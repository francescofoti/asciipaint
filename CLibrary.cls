VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mhLib       As Integer
Private mfReadOnly  As Boolean
Private msFilename  As String

Private mlstDir     As CList

'Class error context
Private mlErrNo   As Long
Private msErrCtx  As String
Private msErrDesc As String

Private Sub ClearErr()
  mlErrNo = 0&
  msErrCtx = ""
  msErrDesc = ""
End Sub

Private Sub SetErr(ByVal psErrCtx As String, ByVal plErrNum As Long, ByVal psErrDesc As String)
  mlErrNo = plErrNum
  msErrCtx = psErrCtx
  msErrDesc = psErrDesc
End Sub

Public Property Get LastErr() As Long
  LastErr = mlErrNo
End Property

Public Property Get LastErrDesc() As String
  LastErrDesc = msErrDesc
End Property

Private Sub Class_Initialize()
  mhLib = 0
End Sub

Private Sub Class_Terminate()
  CloseLibrary
End Sub

Private Function CheckRW() As Boolean
  If mfReadOnly Then
    SetErr "CheckRW", -111&, "Library is in read only mode"
  Else
    CheckRW = True
  End If
End Function

Public Sub CloseLibrary()
  If mhLib <> 0 Then
    Close mhLib
    mfReadOnly = False
  End If
End Sub

Public Property Get IsOpen() As Boolean
  IsOpen = CBool(mhLib <> 0)
End Property

Public Property Get HLibrary() As Integer
  HLibrary = mhLib
End Property

Public Property Get Directory() As CList
  Set Directory = mlstDir
End Property

Public Property Get IsReadOnly() As Boolean
  IsReadOnly = mfReadOnly
End Property

Public Property Get Filename() As String
  Filename = msFilename
End Property

Public Property Get FileSize() As Long
  If mhLib > 0 Then
    FileSize = LOF(mhLib)
  End If
End Property

Public Function OpenLibrary(ByVal psLibFilename As String, ByVal pfReadOnly As Boolean) As Boolean
  Const LOCAL_ERR_CTX As String = "OpenLibrary"
  On Error GoTo OpenLibrary_Err
  ClearErr
  
  Dim fOK       As Boolean
  Dim lstDir    As CList
  Dim hLib      As Integer
  
  CloseLibrary
  fOK = OpenLibraryFile(psLibFilename, mhLib)
  If fOK Then
    mfReadOnly = pfReadOnly
    msFilename = psLibFilename
  Else
    SetErr LOCAL_ERR_CTX, LibLastErr(), LibLastErrDesc()
  End If
  
  OpenLibrary = fOK
  
OpenLibrary_Exit:
  Exit Function

OpenLibrary_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume OpenLibrary_Exit
  Resume
End Function

Public Function LoadDirectory() As Boolean
  Const LOCAL_ERR_CTX As String = "LoadDirectory"
  On Error GoTo LoadDirectory_Err
  ClearErr
  
  Dim fOK       As Boolean
  Dim lstDir    As CList
  
  fOK = LibLoadDirectory(mhLib, lstDir)
  If fOK Then
    Set mlstDir = lstDir
  Else
    SetErr LOCAL_ERR_CTX, LibLastErr(), LibLastErrDesc()
  End If
  
  LoadDirectory = fOK
  
LoadDirectory_Exit:
  Exit Function

LoadDirectory_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume LoadDirectory_Exit
  Resume
End Function

Public Function ReadProperties( _
    ByVal phBlock As Long, _
    ByRef phRetNextPropBlock As Long, _
    ByRef prowCustomProps As CRow _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "ReadProperties"
  Dim fOK       As Boolean
  
  Set prowCustomProps = Nothing
  fOK = LibReadCustProps(mhLib, phBlock, phRetNextPropBlock, prowCustomProps)
  If Not fOK Then
    SetErr LOCAL_ERR_CTX, LibLastErr(), LibLastErrDesc()
  End If
  ReadProperties = fOK
End Function

Public Function ReadPropsBlock( _
    ByVal phBlock As Long, _
    ByRef phRetNextPropBlock As Long, _
    ByRef prowRetNewProps As CRow _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "ReadPropsBlock"
  Dim fOK       As Boolean
  
  Set prowRetNewProps = Nothing
  fOK = LibReadPropsBlock(mhLib, phBlock, phRetNextPropBlock, prowRetNewProps)
  If Not fOK Then
    SetErr LOCAL_ERR_CTX, LibLastErr(), LibLastErrDesc()
  End If
  ReadPropsBlock = fOK
End Function

Public Function ReadHeader(ByRef prowProps As CRow) As Boolean
  Const LOCAL_ERR_CTX As String = "ReadHeader"
  Dim fOK       As Boolean
  Set prowProps = Nothing
  fOK = LibReadHeader(mhLib, prowProps, Nothing, False)
  If Not fOK Then
    SetErr LOCAL_ERR_CTX, LibLastErr(), LibLastErrDesc()
  End If
  ReadHeader = fOK
End Function

Public Function AddFile( _
    ByVal psBlockName As String, _
    ByVal psFilename As String, _
    ByRef prowCustomProps As CRow _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "AddFile"
  Dim fOK     As Boolean
  Dim hBlock  As Long
  
  If Not CheckRW() Then
    Exit Function
  End If
  
  hBlock = LibAddFile(mhLib, psBlockName, psFilename, prowCustomProps)
  If hBlock <> HBLOCK_INVALID Then
    AddFile = True
  Else
    SetErr LOCAL_ERR_CTX, LibLastErr(), LibLastErrDesc()
  End If
End Function

Public Function ExtractFile( _
    ByVal phBlock As Long, _
    ByVal psTargetFilename As String _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "LibExtractFile"
  Dim fOK   As Boolean
  fOK = LibExtractFile(mhLib, phBlock, psTargetFilename)
  If fOK Then
    ExtractFile = True
  Else
    SetErr LOCAL_ERR_CTX, LibLastErr(), LibLastErrDesc()
  End If
End Function

Public Function CreateLibrary( _
    ByVal psLibraryFile As String, _
    ByVal psAuthorName As String, _
    ByVal psCopyright As String, _
    ByVal psDescription As String, _
    ByRef prowCustomProps As CRow _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "CreateLibrary"
  Dim fOK   As Boolean
  fOK = LibCreateLibrary(psLibraryFile, psAuthorName, psCopyright, psDescription, prowCustomProps)
  If fOK Then
    CreateLibrary = True
  Else
    SetErr LOCAL_ERR_CTX, LibLastErr(), LibLastErrDesc()
  End If
End Function

Public Function RenameBlock(ByVal phBlock As Long, ByVal psNewName As String) As Boolean
  Const LOCAL_ERR_CTX As String = "RenameBlock"
  Dim fOK   As Boolean
  fOK = LibRenameBlock(mhLib, phBlock, psNewName)
  If fOK Then
    RenameBlock = True
  Else
    SetErr LOCAL_ERR_CTX, LibLastErr(), LibLastErrDesc()
  End If
End Function

Public Function ExportProperties(ByVal phBlock As Long, ByVal psFilename As String) As Boolean
  Const LOCAL_ERR_CTX As String = "ExportProperties"
  Dim fOK             As Boolean
  Dim hNextSibling    As Long
  
  'almost same as ExtractFile(), but we position on the first properties block
  fOK = LibReadBlockPointer(mhLib, phBlock, hNextSibling, eNext)
  If fOK Then
    fOK = LibExtractFile(mhLib, hNextSibling, psFilename)
    If fOK Then
      ExportProperties = True
    Else
      SetErr LOCAL_ERR_CTX, LibLastErr(), LibLastErrDesc()
    End If
  Else
    SetErr LOCAL_ERR_CTX, LibLastErr(), LibLastErrDesc()
  End If
End Function

Public Function AddProperties( _
    ByVal phBlock As Long, _
    ByVal psBlockName As String, _
    ByVal prowProps As CRow _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "AddProperties"
  On Error GoTo AddProperties_Err
  Dim lstBlocks       As CList
  Dim hBlock          As Long
  Dim hBlock2         As Long
  Dim hNextBlock      As Long
  Dim hPrevBlock      As Long
  Dim fOK             As Boolean
  Dim i               As Integer
  
  ClearErr
  
  Set lstBlocks = New CList
  lstBlocks.AddCol "hblock", 0&, 4&, 0&
  
  'go to the last linked sibling
  hBlock = phBlock
  fOK = False
  Do
    lstBlocks.AddValues hBlock
    fOK = LibReadBlockPointers(mhLib, hBlock, hPrevBlock, hNextBlock)
    If Not fOK Then
      SetErr LOCAL_ERR_CTX, LibLastErr(), LibLastErrDesc()
      GoTo AddProperties_Exit
    End If
    'Debug.Print "1) phBlock=" & phBlock, "hBlock=" & hBlock, "hPrevBlock=" & hPrevBlock, "hNextBlock=" & hNextBlock
    
    'avoid endless loop:
    If hNextBlock <> HBLOCK_INVALID Then
      i = lstBlocks.Find("hblock", hNextBlock)
      If i = 0 Then
        hBlock = hNextBlock
        fOK = LibReadBlockPointers(mhLib, hBlock, hPrevBlock, hNextBlock)
        'Debug.Print "2) phBlock=" & phBlock, "hBlock=" & hBlock, "hPrevBlock=" & hPrevBlock, "hNextBlock=" & hNextBlock
        If Not fOK Then
          SetErr LOCAL_ERR_CTX, LibLastErr(), LibLastErrDesc()
          GoTo AddProperties_Exit
        End If
      Else
        SetErr LOCAL_ERR_CTX, -1&, "Circular link detected at block #" & hBlock & " referencing next block #" & hNextBlock
        GoTo AddProperties_Exit
      End If
    End If
  Loop Until (hNextBlock = HBLOCK_INVALID) Or (Not fOK)
  
  hBlock2 = LibCreateCRowBlock(mhLib, psBlockName, prowProps)
  If hBlock2 <> HBLOCK_INVALID Then
    fOK = LibWriteBlockPointers(mhLib, hBlock2, hBlock, HBLOCK_INVALID)
    If Not fOK Then
      SetErr LOCAL_ERR_CTX, LibLastErr(), LibLastErrDesc()
    End If
    'Debug.Print "phBlock=" & phBlock, "hBlock=" & hBlock, "hPrevBlock=" & hPrevBlock, "hBlock2=" & hBlock2
    fOK = LibWriteBlockPointers(mhLib, hBlock, hPrevBlock, hBlock2)
    If Not fOK Then
      SetErr LOCAL_ERR_CTX, LibLastErr(), LibLastErrDesc()
    End If
  Else
    SetErr LOCAL_ERR_CTX, LibLastErr(), LibLastErrDesc()
  End If
  
  AddProperties = fOK
  
AddProperties_Exit:
  Set lstBlocks = Nothing
  Exit Function

AddProperties_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume AddProperties_Exit
  Resume
End Function
