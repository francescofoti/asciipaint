VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Librarian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IConsoul
Implements ICsMouseEventSink
Implements IMessageReceiver

Private mconLib As CConsoul
Attribute mconLib.VB_VarHelpID = -1

Private Const FILEEXT_LIBRARY  As String = "aplib"
Private Const SUBDIR_LIBRARIES As String = "libraries"
Private Const SUBDIR_EXTRACT   As String = "extractions"

Private Const DEFAULT_FONTNAME    As String = "Courier New"
Private Const DEFAULT_FONTSIZE    As String = "11"
Private Const DEFAULT_MAXCAP      As String = "1000"
Private Const DEFAULT_FORECOLOR As String = "&HD0D0D0"
Private Const DEFAULT_BACKCOLOR As String = "0"

Private moLib     As New CLibrary

Private mlstMRU   As CList
Private miMaxMRU  As Integer
Private Const TOPIC_LIBMRU        As String = "LibrariesMRU"
Private Const INIPARAM_MAXMRU     As String = "MaxMRU"
Private Const DEFAULT_MAXMRU      As Integer = 20
Private Const INIPARAM_MRUCT      As String = "MRUcount"
Private Const INIPARAM_MRUPREFIX  As String = "mru"
Private Const INIPARAM_LASTLIB    As String = "LastLibrary"

'To handle console sizing whether new dialog shown or not
Private mfInNewDialog     As Boolean
Private mlOriginBackColor As Long

Private Const RIBBONTAB_FILE      As String = "File"
Private Const RIBBONTAB_LIBRARY   As String = "Library"
Private Const RIBBONTAB_ELEMENT   As String = "Element"
Private moRibbon As CRibbon

Private Sub cboElement_Click()
  SetControlStates
End Sub

Private Sub cmdBrowseLib_Click()
  Dim iChoice     As Integer
  Dim fLoaded     As Boolean
  Dim sFilename   As String
  Dim sInitialDir As String
  
  On Error GoTo cmdBrowseLib_Click_Err
  If Not LockUI() Then Exit Sub
  
  AppIniFile.GetOption INIOPT_LASTLIBBROWSEPATH, sInitialDir, ""
  sInitialDir = Trim$(sInitialDir)
  If Len(sInitialDir) = 0 Then
    sInitialDir = CombinePath _
      ( _
        CombinePath _
        ( _
          GetSpecialFolder(Application.hWndAccessApp, CSIDL_PERSONAL), _
          APP_NAME _
        ), _
        SUBDIR_LIBRARIES _
      )
  End If
  If Not ExistDir(sInitialDir) Then
    If Not CreatePath(sInitialDir) Then
      sInitialDir = GetSpecialFolder(Application.hWndAccessApp, CSIDL_PERSONAL)
    End If
  End If
  
  With Application.FileDialog(msoFileDialogFilePicker)
    .Title = "Select library file to open"
    .InitialFileName = NormalizePath(sInitialDir)
    .Filters.Clear
    .Filters.Add APP_NAME & " library", "*." & FILEEXT_LIBRARY
    .FilterIndex = 1
    iChoice = .Show()
    If iChoice <> 0 Then
      sFilename = .SelectedItems(1)
      AppIniFile.SetOption INIOPT_LASTLIBBROWSEPATH, (StripFileName(sFilename))
      Me.cboLibrary = sFilename
    End If
  End With

cmdBrowseLib_Click_Exit:
  UnlockUI
  Exit Sub

cmdBrowseLib_Click_Err:
  ShowUFError "An error occured while browsing for a library", Err.Description
  Resume cmdBrowseLib_Click_Exit
End Sub

Private Sub cmdCancel_Click()
  mfInNewDialog = False
  SetControlStates
  ShowNewLibControls False
End Sub

Private Sub cmdCreateLib_Click()
  Const LOCAL_ERR_CTX As String = "cmdCreateLib"
  On Error GoTo cmdCreateLib_Err
  
  Dim sFilename       As String
  Dim sAuthor         As String
  Dim sCopyright      As String
  Dim sDescription    As String
  Dim sMsg            As String
  Dim iRet            As VbMsgBoxResult
  Dim fOK             As Boolean
  
  sFilename = Trim$(Me.txtLibraryName & "")
  sAuthor = Trim$(Me.txtAuthor & "")
  sCopyright = Trim$(Me.txtCopyright & "")
  sDescription = Me.txtDescription & ""
  
  If Len(sFilename) = 0 Then
    MsgBox "Please specify a filename for the new library", vbCritical
    Me.txtLibraryName.SetFocus
    Exit Sub
  End If
  
  If (Len(sAuthor) = 0) Or (Len(sCopyright) = 0) Or (Len(sDescription) = 0) Then
    sMsg = "Warning" & vbCrLf & vbCrLf & "You left some informations empty." & vbCrLf & vbCrLf & _
           "Once created, you will no longer be able to update these informations." & vbCrLf & vbCrLf & _
           "Are you sure you want to continue ?"
    iRet = MsgBox(sMsg, vbExclamation + vbOKCancel + vbDefaultButton2, "Create library")
    If iRet = vbCancel Then
      Exit Sub
    End If
  End If
  
  ConOutLn "Creating library " & sFilename & "..."
  fOK = moLib.CreateLibrary(sFilename, sAuthor, sCopyright, sDescription, Nothing)
  If Not fOK Then
    DisplayLibError
    Exit Sub
  End If
  ConOutLn "Success"
  
  AddToMRU sFilename
  RefreshLibraryCombo
  ShowNewLibControls False
  mfInNewDialog = False
  Form_Resize
  
cmdCreateLib_Exit:
  Exit Sub

cmdCreateLib_Err:
  ShowUFError "Unexpected error while creating library", Err.Description
  Resume cmdCreateLib_Exit
  Resume
End Sub

Private Sub cmdElemDelete_Click()
  '/**/
End Sub

Private Sub cmdElemInfo_Click()
  DisplayElemInfo
End Sub

Private Sub cmdElemProps_Click()
  DisplayElemProps
End Sub

Private Sub cmdElemRename_Click()
  RenameElem
End Sub

Private Sub cmdElemXtract_Click()
  ExtractElem
End Sub

Private Sub cmdLibAddFile_Click()
  Dim iChoice     As Integer
  Dim fLoaded     As Boolean
  Dim sFilename   As String
  Dim sInitialDir As String
  Dim lstFilters  As CList
  Dim fOK         As Boolean
  
  On Error GoTo cmdLibAddFile_Click_Err
  If Not LockUI() Then Exit Sub
  
  Set lstFilters = NewSelectFileFilterList()
  lstFilters.AddValues "AsciiPaint", "*.ascp"
  lstFilters.AddValues "Text files", "*.txt;*.asc;*.vt100"
  lstFilters.AddValues "All", "*.*"
  fOK = SelectLoadFile(INIOPT_LASTLOADPATH, "Add File", lstFilters, sFilename)
  If Not fOK Then
    GoTo cmdLibAddFile_Click_Exit
  End If
  
  Dim lFileLen    As Long
  Dim sBlockName  As String
  
  lFileLen = FileLen(sFilename)
  mconLib.Clear
  ConOutLn "Import file (" & FileSizeAsText(lFileLen) & ") " & sFilename
  
  sBlockName = Trim$(InputBox$("Enter element name (empty cancels):", "Add file", sFilename))
  If Len(sBlockName) = 0 Then
    GoTo cmdLibAddFile_Click_Exit
  End If
  
  ConOutLn "Working..."
  fOK = moLib.AddFile(sBlockName, sFilename, Nothing)
  If fOK Then
    ConOutLn "Success"
    RefreshElements
  Else
    DisplayLibError
  End If
  
cmdLibAddFile_Click_Exit:
  UnlockUI
  Exit Sub

cmdLibAddFile_Click_Err:
  ShowUFError "An error occured while browsing for a library", Err.Description
  Resume cmdLibAddFile_Click_Exit
End Sub

Private Sub cmdLibDir_Click()
  mconLib.Clear
  DisplayDirectory
End Sub

Private Sub cmdLibInfo_Click()
  DisplayLibInfo
End Sub

Private Function GetSelectedElementIndex() As Integer
  If Not IsNull(Me.cboElement) Then
    GetSelectedElementIndex = Me.cboElement.ListIndex + 1
  End If
End Function

Private Sub SetControlStates()
  If moLib.IsOpen And (Not mfInNewDialog) Then
    Me.cmdLibInfo.Enabled = True
    Me.cmdLibProps.Enabled = True
    Me.cmdLibDir.Enabled = True
    Me.cmdLibAddFile.Enabled = Not moLib.IsReadOnly
    'Me.cmdLibAddFolder.Enabled = Not moLib.IsReadOnly
    
    Dim iElem     As Integer
    Dim fPropsOn  As Boolean
    
    iElem = GetSelectedElementIndex()
    If iElem > 0 Then
      Me.cmdElemInfo.Enabled = True
      Me.cmdElemProps.Enabled = True
      Me.cmdElemRename.Enabled = Not moLib.IsReadOnly
      Me.cmdElemXtract.Enabled = True
      Me.cmdElemDelete.Enabled = Not moLib.IsReadOnly
      
      fPropsOn = (moLib.Directory("nextsibling", iElem) <> HBLOCK_INVALID)
      Me.cmdPropsExport.Enabled = fPropsOn
      Me.cmdPropsImport.Enabled = True
    Else
      Me.cmdElemInfo.Enabled = False
      Me.cmdElemProps.Enabled = False
      Me.cmdElemRename.Enabled = False
      Me.cmdElemXtract.Enabled = False
      Me.cmdElemDelete.Enabled = False
    
      Me.cmdPropsExport.Enabled = False
      Me.cmdPropsImport.Enabled = False
    End If
    Me.cboElement.Enabled = True
  Else
    Me.cboElement.Enabled = False
    
    Me.cmdLibInfo.Enabled = False
    Me.cmdLibProps.Enabled = False
    Me.cmdLibDir.Enabled = False
    Me.cmdLibAddFile.Enabled = False
    'Me.cmdLibAddFolder.Enabled = False
    
    Me.cmdElemInfo.Enabled = False
    Me.cmdElemProps.Enabled = False
    Me.cmdElemRename.Enabled = False
    Me.cmdElemXtract.Enabled = False
    Me.cmdElemDelete.Enabled = False
    
    Me.cmdPropsExport.Enabled = False
    Me.cmdPropsImport.Enabled = False
  End If
End Sub

Private Sub ShowNewLibControls(ByVal pfShow As Boolean)
  Dim vCtl  As Variant
  
  Me.cmdDummy.Visible = True
  Me.cmdDummy.SetFocus
  If pfShow Then
    Me.Section(AcSection.acDetail).BackColor = mlOriginBackColor
    Me.cmdBrowseLib.Enabled = False
    Me.cmdNewLib.Enabled = False
    Me.chkReadWrite.Enabled = False
    Me.cmdOpenLib.Enabled = False
  Else
    Me.Section(AcSection.acDetail).BackColor = mconLib.BackColor
    Me.cmdBrowseLib.Enabled = True
    Me.cmdNewLib.Enabled = True
    Me.chkReadWrite.Enabled = True
    Me.cmdOpenLib.Enabled = True
  End If
  For Each vCtl In Me.Section(AcSection.acDetail).Controls
    If Not vCtl Is Me.cboLibrary Then
      If GetTagParam(vCtl.Tag, "newlibctl") = "newlibctl" Then
        vCtl.Visible = pfShow
      End If
    End If
  Next
  If pfShow Then
    Me.txtLibraryName.SetFocus
  End If
  
  Form_Resize
End Sub

Private Sub CloseLibrary()
  On Error Resume Next
  moLib.CloseLibrary
  SetControlStates
  Me.cboElement = Null
  Me.cboElement.RowSource = ""
End Sub

Private Sub cmdLibProps_Click()
  DisplayLibProps
End Sub

Private Sub cmdNewLib_Click()
  mfInNewDialog = True
  SetControlStates
  ShowNewLibControls True
End Sub

Private Sub cmdOpenLib_Click()
  Const LOCAL_ERR_CTX As String = "cmdOpenLib"
  On Error GoTo cmdOpenLib_Err
  
  Dim sLibFilename  As String
  Dim fReadOnly     As Boolean
  Dim fOK           As Boolean
  
  sLibFilename = Trim$(Me.cboLibrary & "")
  If Len(sLibFilename) = 0 Then
    MsgBox "Please specify the library file to open", vbCritical
    Me.cboLibrary.SetFocus
    Exit Sub
  End If
  If Not ExistFile(sLibFilename) Then
    MsgBox "Library file [" & sLibFilename & "] not found", vbCritical
    Me.cboLibrary.SetFocus
    Exit Sub
  End If
  
  CloseLibrary
  
  fReadOnly = Not Me.chkReadWrite
  
  mconLib.Clear
  ConOutLn "Opening library (" & IIf(fReadOnly, "RO", "RW") & ") " & StripFilePath(sLibFilename)
  fOK = moLib.OpenLibrary(sLibFilename, fReadOnly)
  If fOK Then
    'save last opened library in ini
    AppIniFile.Section = TOPIC_LIBMRU
    AppIniFile.SetString INIPARAM_LASTLIB, sLibFilename
  Else
    DisplayLibError
    GoTo cmdOpenLib_Exit
  End If
  
  RefreshElements
  
  SetControlStates
  
  DisplayLibInfo
  DisplayDirectory
  
  AddToMRU sLibFilename
  RefreshLibraryCombo
  
cmdOpenLib_Exit:
  Exit Sub

cmdOpenLib_Err:
  ShowUFError "Unexpected error while opening library", Err.Description
  Resume cmdOpenLib_Exit
  Resume
End Sub

Private Sub cmdPropsExport_Click()
  ExportElemProps
End Sub

Private Sub cmdPropsImport_Click()
  ImportElemProps
End Sub

Private Sub cmdSelectNewLib_Click()
  Dim iChoice     As Integer
  Dim fLoaded     As Boolean
  Dim sFilename   As String
  Dim sInitialDir As String
  Dim fOK         As Boolean
  
  On Error GoTo cmdSelectNewLib_Click_Err
  Call LockUI   'Don't test if it succeeded if the canvas form is not loaded
  
  AppIniFile.GetOption INIOPT_LASTLIBBROWSEPATH, sInitialDir, ""
  sInitialDir = Trim$(sInitialDir)
  If Len(sInitialDir) = 0 Then
    sInitialDir = CombinePath _
      ( _
        CombinePath _
        ( _
          GetSpecialFolder(Application.hWndAccessApp, CSIDL_PERSONAL), _
          APP_NAME _
        ), _
        SUBDIR_LIBRARIES _
      )
  End If
  
  fOK = VBGetSaveFileName( _
    sFilename, _
    "", _
    True, _
    APP_NAME & " library|*." & FILEEXT_LIBRARY, _
    1, _
    sInitialDir, _
    "New library...", _
    FILEEXT_LIBRARY, _
    Me.hWnd)
  If fOK Then
    Me.txtLibraryName = sFilename
  End If

cmdSelectNewLib_Click_Exit:
  UnlockUI
  Exit Sub

cmdSelectNewLib_Click_Err:
  ShowUFError "An error occured while selecting the new library name", Err.Description
  Resume cmdSelectNewLib_Click_Exit
End Sub

'**************************
'
' Form events
'
'**************************

Private Sub Form_Load()
  CreateRibbon
  Set mconLib = New CConsoul
  mlOriginBackColor = Me.Section(AcSection.acDetail).BackColor
  CreateConsoul
  InitMRU
  PositionRibbonControls
  MessageManager.SubscribeMulti Me, Array(MSGTOPIC_LOCKUI, MSGTOPIC_UNLOCKUI, MSGTOPIC_UNLOADNOW)
End Sub

Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
  On Error Resume Next
  Dim hWnd                  As LongPtr
  Dim ptMouse               As POINTAPI
  Dim i                     As Long
  
  apiGetCursorPos ptMouse
  hWnd = apiWindowFromPoint(ptMouse.x, ptMouse.y)
  If hWnd = moRibbon.TabConsole.hWnd Then
    moRibbon.OnMousWheel Count
  End If
End Sub

Private Sub Form_Resize()
  If mconLib Is Nothing Then Exit Sub
  
  Const MARGIN As Integer = 2
  Dim rcClient      As RECT
  GetClientRect Me.hWnd, rcClient
  
  Dim iHdHeight   As Integer
  Dim iFtHeight   As Integer
  Dim iLinTop     As Integer
  Dim iTop        As Integer
  Dim iHeight     As Integer
  
  moRibbon.TabConsole.MoveWindow 0, 0, rcClient.Right - rcClient.left, TwipsToPixelsY(Me.rectRibbon.Height)
  moRibbon.OnResize
  Me.rectToolbar.Width = PixelsToTwipsX(rcClient.Right)
  Me.linToolbar.left = 0
  Me.linToolbar.Top = Me.rectToolbar.Top + Me.rectToolbar.Height - PixelsToTwipsY(1)
  Me.linToolbar.Width = Me.Width
  
  iHdHeight = TwipsToPixelsY(Me.Section(AcSection.acHeader).Height)  'Height of the header in pixels
  iHdHeight = iHdHeight + TwipsToPixelsY(Me.rectToolbar.Top + Me.rectToolbar.Height)
  
  iTop = iHdHeight
  If mfInNewDialog Then
    iTop = TwipsToPixelsY(Me.cmdCreateLib.Top + Me.cmdCreateLib.Height) + MARGIN * 2
  End If
  iHeight = (rcClient.Bottom - rcClient.Top) - iTop - iFtHeight - 2 * MARGIN
  
  If iHeight >= mconLib.LineHeight Then
    mconLib.MoveWindow MARGIN, iTop, _
                       rcClient.Right - 2 * MARGIN, _
                       iHeight
    mconLib.ShowWindow True
  Else
    mconLib.ShowWindow False
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  SaveMRU
  MessageManager.Unsubscribe Me.Name, ""
  mconLib.Detach
  Set moRibbon = Nothing
  Set mconLib = Nothing
End Sub

'**************************
'
' IConsoul implementation
'
'**************************

Private Property Let IConsoul_AutoRedraw(ByVal RHS As Boolean)
  If Not mconLib Is Nothing Then
    IConsoul_AutoRedraw = mconLib.AutoRedraw
  End If
End Property

Private Property Get IConsoul_AutoRedraw() As Boolean
  If Not mconLib Is Nothing Then
    IConsoul_AutoRedraw = mconLib.AutoRedraw
  End If
End Property

Private Sub IConsoul_ConOut(ByVal psInfo As String)
  ConOut psInfo
End Sub

Private Sub IConsoul_ConOutLn(ByVal psInfo As String, Optional ByVal piQBColorText As Integer = -1)
  ConOutLn psInfo
End Sub

Private Function IConsoul_GetConsoul() As CConsoul
  If Not mconLib Is Nothing Then
    Set IConsoul_GetConsoul = mconLib
  End If
End Function

Private Sub IConsoul_RefreshWindow()
  If Not mconLib Is Nothing Then
    mconLib.RefreshWindow
  End If
End Sub

'**************************
'
' Private methods
'
'**************************

Private Sub ConOut(ByVal psText As String)
  If Not mconLib Is Nothing Then
    mconLib.Output psText
  End If
End Sub

Private Sub ConOutLn(ByVal psText As String)
  If Not mconLib Is Nothing Then
    mconLib.OutputLn psText
  End If
End Sub

Private Function CreateConsoul() As Boolean
  Dim hwndParent  As LongPtr
  
  On Error GoTo CreateConsouls_Err
  
  hwndParent = Me.hWnd
  
  'We repeatedly call this function to create/destroy the console windows
  If Not mconLib Is Nothing Then
    'ConsoulEventDispatcher.UnregisterEventSink mconLib.hWnd
    Set mconLib = Nothing
  End If
  
  'The console window displaying the character set
  Set mconLib = New CConsoul
  Dim sConFontName  As String
  Dim sConFontSize  As String
  Dim sConMaxCap    As String
  Dim sConForeColor As String
  Dim sConBackColor As String
  Dim lIniValue     As Long
  
  AppIniFile.GetOption INIOPT_CONLIBFONTNAME, sConFontName, DEFAULT_FONTNAME
  AppIniFile.GetOption INIOPT_CONLIBFONTSIZE, sConFontSize, DEFAULT_FONTSIZE
  AppIniFile.GetOption INIOPT_CONLIBMAXCAP, sConMaxCap, DEFAULT_MAXCAP
  AppIniFile.GetOption INIOPT_CONLIBFORECOLOR, sConForeColor, DEFAULT_FORECOLOR
  AppIniFile.GetOption INIOPT_CONLIBBACKCOLOR, sConBackColor, DEFAULT_BACKCOLOR
  With mconLib
    .LineSpacing(elsTop) = 1
    .FontName = sConFontName
    lIniValue = Val(sConFontSize)
    If (lIniValue > 18) Then
      lIniValue = 18
    End If
    If lIniValue < 8 Then
      lIniValue = 8
    End If
    .FontSize = lIniValue
    lIniValue = Val(sConMaxCap)
    If (lIniValue > 2000) Then
      lIniValue = 2000
    End If
    If lIniValue < 250 Then
      lIniValue = 250
    End If
    .MaxCapacity = lIniValue
    .ForeColor = Val(sConForeColor)
    .BackColor = Val(sConBackColor)
    Me.Section(AcSection.acDetail).BackColor = .BackColor
  End With
  'Create the console window and tell the library that we want click feedback
  'If Not mconLib.Attach(hWndParent, 0, 0, 0, 0, AddressOf MSupport.OnConsoulMouseButton, piCreateAttributes:=LW_RENDERMODEBYLINE Or LW_TRACK_ZONES) Then
  If Not mconLib.Attach(hwndParent, 0, 0, 0, 0, 0, piCreateAttributes:=LW_RENDERMODEBYLINE Or LW_TRACK_ZONES) Then
    MsgBox "Failed to create console window", vbCritical
    GoTo CreateConsouls_Exit
  End If
  
  'Let the system know that click for our canvas console should arrive here
  'ConsoulEventDispatcher.RegisterEventSink mconLib.hWnd, Me, eCsMouseEvent
  'Show the console window
  mconLib.ShowWindow True
  
  CreateConsoul = True
  
CreateConsouls_Exit:
  Exit Function

CreateConsouls_Err:
  MsgBox "Failed to create consoul's output windows"
End Function

Private Function FormatH1(ByVal psTitle As String) As String
  FormatH1 = VT_BOLD_ON() & VT_FCOLOR(vbYellow) & psTitle & VT_BOLD_OFF()
End Function

Private Sub DisplayKeyPair(ByVal psName As String, ByVal psText As String, ByVal pfFixed As Boolean)
  If pfFixed Then
    ConOutLn VT_FCOLOR(vbWhite) & StrBlock(psName, " ", 12) & ": " & VT_FCOLOR(mconLib.ForeColor) & psText
  Else
    ConOutLn VT_FCOLOR(vbWhite) & psName & "=" & VT_FCOLOR(mconLib.ForeColor) & psText
  End If
End Sub

Private Sub DisplayLibError()
  ConOutLn VT_FCOLOR(vbMagenta) & VT_UNDL_ON() & "Failed" & VT_UNDL_OFF() & ": " & moLib.LastErrDesc
End Sub

Private Sub DisplayProps(poRow As CRow, ByVal pfFixed As Boolean)
  If Not poRow Is Nothing Then
    Dim i       As Integer
    For i = 1 To poRow.ColCount
      DisplayKeyPair poRow.ColName(i), poRow(i), pfFixed
    Next i
  Else
    ConOutLn "(empty list)"
  End If
End Sub

Private Sub DisplayLibInfo()
  Dim fOK       As Boolean
  Dim rowProps  As CRow
  
  mconLib.Clear
  
  ConOut "File " & StripFilePath(moLib.Filename)
  ConOut ", size: " & FileSizeAsText(moLib.FileSize)
  ConOutLn " (" & moLib.FileSize & " bytes)"
  ConOutLn "Access mode is " & IIf(moLib.IsReadOnly, "read only", "read/write")
  ConOutLn FormatH1("Library information:")
  
  fOK = moLib.ReadHeader(rowProps)
  If Not fOK Then
    DisplayLibError
    Exit Sub
  End If
  DisplayProps rowProps, True
End Sub

Private Sub DisplayLibProps()
  Dim fOK         As Boolean
  Dim rowProps    As CRow
  Dim hNextBlock  As Long
  
  mconLib.Clear
  
  ConOutLn FormatH1("Library properties:")
  DisplayPropsChain 1&
End Sub

'https://www.mrexcel.com/board/threads/convert-filesize-to-kb-mb-gb-etc-in-vba.829884/
Private Function FileSizeAsText(ByVal plBytes As Long) As String
  Dim sUnit As String
  Dim sRet  As String
  
  Select Case plBytes
  Case 0 To 1023
    sRet = Format$(plBytes, "0") & " bytes"
  Case 1024 To 1048575
    sRet = Format$(plBytes / 1024, "0") & " kb"
  Case 1048576 To 1073741823
    sRet = Format$(plBytes / 1048576, "0") & " mb"
  Case 1073741824 To 1.11111111111074E+20 'Not going to happen often here, we handle a Long value
    sRet = Format(plBytes / 1073741823, "0.00") & " gb"
  End Select

  FileSizeAsText = sRet
End Function

Private Sub DisplayDirectory()
  Dim lstDir      As CList
  Dim fEmpty      As Boolean
  Dim i           As Long
  Dim iSeqColLen  As Integer
  
  Set lstDir = moLib.Directory
  If lstDir Is Nothing Then
    fEmpty = True
  Else
    If lstDir.Count = 0 Then
      fEmpty = True
    End If
  End If
  
  If Not fEmpty Then
    iSeqColLen = Len(CStr(lstDir.Count))
    ConOutLn FormatH1("Library directory")
    ConOutLn lstDir.Count & " element(s) in " & StripFilePath(moLib.Filename) & ":"
    ConOutLn VT_FCOLOR(vbWhite) & "# " & Space$(iSeqColLen - 1) & StrBlock("size", " ", 12) & "name"
    For i = 1 To lstDir.Count
      ConOut Format$(i, String$(iSeqColLen, "0")) & " "
      ConOut StrBlock(FileSizeAsText(lstDir("blocksize", i)), " ", 11) & " "
      ConOutLn lstDir("blockname", i)
    Next i
  Else
    ConOutLn "(no entries found)"
  End If
End Sub

Private Sub DisplayElemInfo()
  Dim iElemIndex    As Integer
  Dim rowElem       As CRow
  Dim rowDisp   As New CRow
  
  iElemIndex = GetSelectedElementIndex()
  If iElemIndex < 1 Then
    Exit Sub
  End If
  
  moLib.Directory.GetRow rowElem, iElemIndex
  
  mconLib.Clear
  ConOutLn FormatH1(rowElem("blockname"))
  
  rowDisp.AddCol "ElementID", rowElem("hblock"), 0, 0
  Select Case rowElem("blocktype")
  Case BLOCKTYPE_CROW
    rowDisp.AddCol "Type", "Property list", 0, 0
  Case BLOCKTYPE_BINARY
    rowDisp.AddCol "Type", "Binary data", 0, 0
  End Select
  rowDisp.AddCol "Size", FileSizeAsText(rowElem("blocksize")) & " (" & rowElem("blocksize") & " bytes)", 0, 0
  rowDisp.AddCol "Next sibling", rowElem("nextsibling"), 0, 0
  rowDisp.AddCol "Prev sibling", rowElem("prevsibling"), 0, 0
  rowDisp.AddCol "Attributes", LibAttribsToText(rowElem("attribs")), 0, 0
  rowDisp.AddCol "Tag", rowElem("tag"), 0, 0
  
  DisplayProps rowDisp, True
  
End Sub

Private Sub DisplayPropsChain(ByVal phBlock As Long)
  Dim fOK           As Boolean
  Dim hNextBlock    As Long
  Dim lstBlocks     As CList
  Dim i             As Integer
  Dim rowProps      As CRow
  Dim hBlock        As Long
  Dim iBlockCt      As Integer
  
  Set lstBlocks = New CList
  lstBlocks.AddCol "hblock", 0&, 4&, 0&
  
  iBlockCt = 1
  fOK = moLib.ReadProperties(phBlock, hNextBlock, rowProps)
  If fOK Then
    lstBlocks.AddValues phBlock
    ConOutLn VT_UNDL_ON() & "Property set #" & iBlockCt & VT_UNDL_OFF()
    DisplayProps rowProps, False
  Else
    DisplayLibError
    Exit Sub
  End If
  If hNextBlock <> HBLOCK_INVALID Then
    
    Do While hNextBlock <> HBLOCK_INVALID
      hBlock = hNextBlock
      i = lstBlocks.Find("hblock", hBlock)
      If i = 0 Then
        lstBlocks.AddValues hBlock
        iBlockCt = iBlockCt + 1
        ConOutLn VT_UNDL_ON() & "Property set #" & iBlockCt & VT_UNDL_OFF()
        fOK = moLib.ReadPropsBlock(hBlock, hNextBlock, rowProps)
        If fOK Then
          DisplayProps rowProps, False
        Else
          DisplayLibError
          Exit Sub
        End If
      Else
        ConOutLn "Circular link detected at block #" & hBlock
        Exit Do
      End If
    Loop
  End If
End Sub

Private Sub DisplayElemProps()
  Dim iElemIndex    As Integer
  Dim rowProps      As CRow
  Dim fOK           As Boolean
  Dim hBlock        As Long
  
  iElemIndex = GetSelectedElementIndex()
  If iElemIndex < 1 Then
    Exit Sub
  End If
  
  hBlock = moLib.Directory("hblock", iElemIndex)
  mconLib.Clear
  If hBlock <> HBLOCK_INVALID Then
    ConOutLn FormatH1("Properties of " & moLib.Directory("blockname", iElemIndex) & ":")
    DisplayPropsChain hBlock
  Else
    ConOutLn "This element has no properties"
  End If
  Set rowProps = Nothing
End Sub

Private Sub RenameElem()
  Dim iElemIndex    As Integer
  Dim fOK           As Boolean
  Dim hBlock        As Long
  Dim sNewName      As String
  
  iElemIndex = GetSelectedElementIndex()
  If iElemIndex < 1 Then
    Exit Sub
  End If
  
  hBlock = moLib.Directory("hblock", iElemIndex)
  If hBlock <> HBLOCK_INVALID Then
    sNewName = moLib.Directory("blockname", iElemIndex)
    If (sNewName = BLOCKNAME_LIBHEADER) Or (sNewName = BLOCKNAME_LIBCUSTPROPS) Then
      MsgBox "This element cannot be renamed", vbCritical
      Exit Sub
    End If
    sNewName = InputBox$("Enter a new name for the element:", "Rename element", sNewName)
    If Len(sNewName) = 0 Then Exit Sub
    fOK = moLib.RenameBlock(hBlock, sNewName)
    If fOK Then
      RefreshElements
      On Error Resume Next
      Me.cboElement = sNewName
      DisplayElemInfo
    Else
      DisplayLibError
      Exit Sub
    End If
  End If
End Sub

Private Sub ExtractElem()
  Dim iElemIndex    As Integer
  Dim fOK           As Boolean
  Dim hBlock        As Long
  Dim sFilename     As String
  Dim sInitialDir   As String
  
  iElemIndex = GetSelectedElementIndex()
  If iElemIndex < 1 Then
    Exit Sub
  End If
  
  sFilename = moLib.Directory("blockname", iElemIndex)
  hBlock = moLib.Directory("hblock", iElemIndex)
  
  sFilename = StripFilePath(sFilename)
  
  AppIniFile.GetOption INIOPT_LASTLIBXTRACTPATH, sInitialDir, ""
  sInitialDir = Trim$(sInitialDir)
  If Len(sInitialDir) = 0 Then
    sInitialDir = CombinePath _
      ( _
        CombinePath _
        ( _
          GetSpecialFolder(Application.hWndAccessApp, CSIDL_PERSONAL), _
          APP_NAME _
        ), _
        SUBDIR_EXTRACT _
      )
  End If
  
  fOK = VBGetSaveFileName( _
    sFilename, _
    "", _
    False, _
    "All files|*.*", _
    1, _
    sInitialDir, _
    "Extract element", _
    "", _
    Me.hWnd)
  If fOK Then
    If Not ExistFile(sFilename) Then
      fOK = moLib.ExtractFile(hBlock, sFilename)
      If fOK Then
        MsgBox "Element extracted into file [" & sFilename & "]", vbInformation
      Else
        DisplayLibError
      End If
    Else
      MsgBox "The file [" & sFilename & "] already exists. Please choose another file", vbCritical
    End If
  End If
End Sub

Private Sub ExportElemProps()
  Dim iElemIndex    As Integer
  Dim fOK           As Boolean
  Dim hBlock        As Long
  Dim sFilename     As String
  Dim sInitialDir   As String
  
  iElemIndex = GetSelectedElementIndex()
  If iElemIndex < 1 Then
    Exit Sub
  End If
  
  sFilename = moLib.Directory("blockname", iElemIndex)
  hBlock = moLib.Directory("hblock", iElemIndex)
  
  sFilename = "props_" & StripFilePath(sFilename) & ".txt"
  
  AppIniFile.GetOption INIOPT_LASTLIBXTRACTPATH, sInitialDir, ""
  sInitialDir = Trim$(sInitialDir)
  If Len(sInitialDir) = 0 Then
    sInitialDir = CombinePath _
      ( _
        CombinePath _
        ( _
          GetSpecialFolder(Application.hWndAccessApp, CSIDL_PERSONAL), _
          APP_NAME _
        ), _
        SUBDIR_EXTRACT _
      )
  End If
  
  fOK = VBGetSaveFileName( _
    sFilename, _
    "", _
    False, _
    "Text files|*.txt", _
    1, _
    sInitialDir, _
    "Extract element", _
    "txt", _
    Me.hWnd)
  If fOK Then
    If Not ExistFile(sFilename) Then
      fOK = moLib.ExportProperties(hBlock, sFilename)
      If fOK Then
        MsgBox "Properties extracted into file [" & sFilename & "]", vbInformation
      Else
        DisplayLibError
      End If
    Else
      MsgBox "The file [" & sFilename & "] already exists. Please choose another file", vbCritical
    End If
  End If
End Sub

Private Sub ImportElemProps()
  Dim iElemIndex    As Integer
  Dim fOK           As Boolean
  Dim hBlock        As Long
  Dim sFilename     As String
  Dim sInitialDir   As String
  Dim lstFilters    As CList
  Dim oIniFile      As New CIniFile
  Dim iSectionCt    As Long
  Dim i             As Long
  Dim asIniKey()    As String
  Dim lFileLen      As Long
  Dim rowProps      As New CRow
  Dim sBlockName    As String
  
  Const MAX_BUFSIZE As Long = 8192&
  
  iElemIndex = GetSelectedElementIndex()
  If iElemIndex < 1 Then
    Exit Sub
  End If
  
  Set lstFilters = NewSelectFileFilterList()
  lstFilters.AddValues "Text files", "*.txt"
  fOK = SelectLoadFile(INIOPT_LASTLIBIMPORTPATH, "Import properties", lstFilters, sFilename)
  If Not fOK Then
    GoTo ImportElemProps_Exit
  End If

  If Not ExistFile(sFilename) Then
    MsgBox "The file [" & sFilename & "] doesn't exist or cannot be accessed", vbCritical
    GoTo ImportElemProps_Exit
  End If
  
  lFileLen = FileLen(sFilename)
  If lFileLen > MAX_BUFSIZE Then
    MsgBox "The byte size of a property file cannot exceed " & MAX_BUFSIZE & " bytes." & vbCrLf & _
           "File [" & sFilename & "] is " & lFileLen & " bytes and cannot be imported", vbCritical
    GoTo ImportElemProps_Exit
  End If
  
  oIniFile.Filename = sFilename
  oIniFile.Section = "properties"
  iSectionCt = oIniFile.GetSectionEntries(asIniKey(), MAX_BUFSIZE)
  For i = 1 To iSectionCt
    rowProps.AddCol asIniKey(i), oIniFile.GetString(asIniKey(i), MAX_BUFSIZE), 0, 0
  Next i
  
  hBlock = moLib.Directory("hblock", iElemIndex)
  sBlockName = StripFileExt(StripFilePath(sFilename))
  fOK = moLib.AddProperties(hBlock, sBlockName, rowProps)
  If fOK Then
    sBlockName = moLib.Directory("blockname", iElemIndex)
    MsgBox "Properties added to element [" & sBlockName & "] element", vbInformation
    RefreshElements
    DisplayElemProps
  Else
    DisplayLibError
  End If
ImportElemProps_Exit:
End Sub

'**************************
'
' MRU
'
'**************************

Private Function IsInMRU(ByVal psFilename As String) As Boolean
  Dim iFind   As Long
  iFind = mlstMRU.Find("filename", psFilename)
  IsInMRU = CBool(iFind > 0)
End Function

Private Sub RefreshLibraryCombo()
  Dim i         As Integer
  Dim sSource   As String
  
  If mlstMRU.Count > 0 Then
    For i = 1 To mlstMRU.Count
      If i > 1 Then
        sSource = sSource & ";"
      End If
      sSource = sSource & mlstMRU("filename", i)
    Next i
  End If
  
  On Error Resume Next
  Me.cboLibrary.RowSource = sSource
End Sub

Private Sub InitMRU()
  On Error GoTo InitMRU_Err
  Dim i         As Integer
  Dim iCount    As Integer
  Dim sEntryKey As String
  Dim sFilename As String
  
  If Len(AppIniFile.Filename) = 0 Then
    ConnectIniFile
  End If
  Set mlstMRU = New CList
  mlstMRU.ArrayDefine Array("filename"), Array(vbString)
  
  AppIniFile.Section = TOPIC_LIBMRU
  miMaxMRU = AppIniFile.GetInt(INIPARAM_MAXMRU, DEFAULT_MAXMRU)
  If miMaxMRU > 50 Then
    miMaxMRU = 50
  End If
  
  iCount = AppIniFile.GetInt(INIPARAM_MRUCT, 0)
  If iCount > miMaxMRU Then
    iCount = miMaxMRU
  End If
  
  For i = 1 To iCount
    sEntryKey = INIPARAM_MRUPREFIX & i
    sFilename = AppIniFile.GetString(sEntryKey)
    AddToMRU sFilename
  Next i
  
  RefreshLibraryCombo
  On Error Resume Next
  
  'Now see if there's a last opened library file in ini
  sFilename = AppIniFile.GetString(INIPARAM_LASTLIB)
  If Len(sFilename) > 0 Then
    Me.cboLibrary = sFilename
  Else
    If mlstMRU.Count > 0 Then
      Me.cboLibrary = mlstMRU("filename", mlstMRU.Count)
    End If
  End If
InitMRU_Err:
End Sub

Private Sub SaveMRU()
  Dim i         As Integer
  Dim sEntryKey As String
  
  If AppIniFile Is Nothing Then Exit Sub
  If mlstMRU Is Nothing Then Exit Sub
  
  AppIniFile.Section = TOPIC_LIBMRU
  AppIniFile.SetString INIPARAM_MRUCT, CStr(mlstMRU.Count)
  
  If mlstMRU.Count > 0 Then
    For i = 1 To mlstMRU.Count
      sEntryKey = INIPARAM_MRUPREFIX & i
      AppIniFile.SetString sEntryKey, mlstMRU("filename", i)
    Next i
  End If
End Sub

Private Sub AddToMRU(ByVal psFilename As String)
  On Error GoTo AddToMRU_Err
  
  If mlstMRU.Count >= miMaxMRU Then
    Exit Sub
  End If

  If Not IsInMRU(psFilename) Then
    mlstMRU.AddValues psFilename
  End If

AddToMRU_Err:
End Sub

Private Sub RefreshElements()
  On Error GoTo RefreshElements_Err
  Dim fOK       As Boolean
  Dim i         As Integer
  Dim sSource   As String
  
  'Populate elements combo
  Me.cboElement = Null
  fOK = moLib.LoadDirectory()
  If Not fOK Then
    ShowUFError "Error loading library directory", moLib.LastErrDesc
    GoTo RefreshElements_Exit
  End If
  If moLib.Directory.Count > 0 Then
    For i = 1 To moLib.Directory.Count
      If i > 1 Then
        sSource = sSource & ";"
      End If
      sSource = sSource & moLib.Directory("blockname", i)
    Next i
    Me.cboElement.RowSource = sSource
    Me.cboElement.Enabled = True
  Else
    Me.cboElement = Null
    Me.cboElement.RowSource = ""
    Me.cboElement.Enabled = False
  End If
RefreshElements_Exit:
  Exit Sub
RefreshElements_Err:
  On Error Resume Next
  Me.cboElement = Null
  Me.cboElement.RowSource = ""
  Me.cboElement.Enabled = False
  Resume RefreshElements_Exit
End Sub

'*****************
'
' Ribbon related
'
'*****************

Private Function CreateRibbon() As Boolean
  On Error GoTo CreateRibbon_Err
  
  Dim fOK         As Boolean
  
  Set moRibbon = New CRibbon
  fOK = moRibbon.Init( _
          Me, _
          vbBlack, _
          Me.Section(AcSection.acDetail).BackColor, _
          Me.rectToolbar.BackColor, _
          Me, Me.cmdDummy _
        )
  If fOK Then
    'create consol for ribbon tab strip
    moRibbon.TabControl.AddTab RIBBONTAB_FILE, "File"
    moRibbon.TabControl.AddTab RIBBONTAB_LIBRARY, "Library"
    moRibbon.TabControl.AddTab RIBBONTAB_ELEMENT, "Element"
  Else
    ShowUFError "Failed to create ribbon control", moRibbon.LastErrDesc
  End If
  
  CreateRibbon = fOK
  
CreateRibbon_Exit:
  Exit Function

CreateRibbon_Err:
  ShowUFError "Failed to create the ribbon", Err.Description
  Resume CreateRibbon_Exit
  Resume
End Function

Private Sub PositionRibbonControls()
  Dim iTop        As Integer
  iTop = Me.rectRibRef1.Top
  moRibbon.PositionBandControls iTop
End Sub

'****************************
'
' ICsMouseEventSink interface
'
'****************************

Private Function ICsMouseEventSink_OnMouseButton(ByVal phWnd As Long, ByVal piEvtCode As Integer, ByVal pwParam As Integer, ByVal piZoneID As Integer, ByVal piRow As Integer, ByVal piCol As Integer, ByVal piPosX As Integer, ByVal piPosY As Integer) As Integer
  If phWnd = moRibbon.TabConsole.hWnd Then
    ICsMouseEventSink_OnMouseButton = moRibbon.OnTabsMouseButton(phWnd, piEvtCode, pwParam, piZoneID, piRow, piCol, piPosX, piPosY)
  End If
End Function

'****************************
'
' IMessageReceiver interface
'
'****************************

Private Property Get IMessageReceiver_ClientID() As String
  IMessageReceiver_ClientID = Me.Name
End Property

Private Function IMessageReceiver_OnMessageReceived(ByVal psSenderID As String, ByVal psTopic As String, pvData As Variant) As Long
  Dim fOK     As Boolean
  Select Case psTopic
  Case MSGTOPIC_LOCKUI
    If psSenderID <> Me.Name Then
      fOK = MForms.FormSetAllowEdits(Me, False, "", "", 0)
    End If
  Case MSGTOPIC_UNLOCKUI
    If psSenderID <> Me.Name Then
      fOK = MForms.FormSetAllowEdits(Me, True, "", "", 0)
    End If
  Case MSGTOPIC_CANUNLOAD
    IMessageReceiver_OnMessageReceived = 0& 'can always unload
  Case MSGTOPIC_UNLOADNOW
    DoCmd.Close acForm, Me.Name
  End Select
End Function
