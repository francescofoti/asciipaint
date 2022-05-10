Attribute VB_Name = "MCommonDlg"
Option Compare Database
Option Explicit

#If Win64 Then
Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
'#if (_WIN32_WINNT >= 0x0500)
    pvReserved As LongPtr
    dwReserved As Long
    FlagsEx As Long
'#endif // (_WIN32_WINNT >= 0x0500)
End Type

Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" _
    Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare PtrSafe Function GetSaveFileName Lib "comdlg32.dll" _
    Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare PtrSafe Function GetFileTitle Lib "comdlg32.dll" _
    Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer

Private Type BROWSEINFO
  hwndOwner      As LongPtr
  pidlRoot       As LongPtr
  pszDisplayName As String
  lpszTitle      As String
  ulFlags        As Long
  lpfnCallback   As LongPtr
  lParam         As Long
  iImage         As Long
End Type

Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32" (lpBI As BROWSEINFO) As LongPtr
Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Any, ByVal lpBuffer As String) As LongPtr
'Private Declare PtrSafe Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As LongPtr

Private Declare PtrSafe Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As TChooseColor) As LongPtr
#Else
Private Type OPENFILENAME
  lStructSize As Long          ' Filled with UDT size
  hwndOwner As Long            ' Tied to Owner
  hInstance As Long            ' Ignored (used only by templates)
  lpstrFilter As String        ' Tied to Filter
  lpstrCustomFilter As String  ' Ignored
  nMaxCustFilter As Long       ' Ignored
  nFilterIndex As Long         ' Tied to FilterIndex
  lpstrFile As String          ' Tied to FileName
  nMaxFile As Long             ' Handled internally
  lpstrFileTitle As String     ' Tied to FileTitle
  nMaxFileTitle As Long        ' Handled internally
  lpstrInitialDir As String    ' Tied to InitDir
  lpstrTitle As String         ' Tied to DlgTitle
  Flags As Long                ' Tied to Flags
  nFileOffset As Integer       ' Ignored
  nFileExtension As Integer    ' Ignored
  lpstrDefExt As String        ' Tied to DefaultExt
  lCustData As Long            ' Ignored (needed for hooks)
  lpfnHook As Long             ' Ignored
  lpTemplateName As Long       ' Ignored
End Type

Private Declare Function GetOpenFileName Lib "COMDLG32" _
    Alias "GetOpenFileNameA" (file As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32" _
    Alias "GetSaveFileNameA" (file As OPENFILENAME) As Long
Private Declare Function GetFileTitle Lib "COMDLG32" _
    Alias "GetFileTitleA" (ByVal szFile As String, _
    ByVal szTitle As String, ByVal cbBuf As Long) As Long

Private Type BROWSEINFO
  hwndOwner      As Long
  pidlRoot       As Long
  pszDisplayName As Long
  lpszTitle      As Long
  ulFlags        As Long
  lpfnCallback   As Long
  lParam         As Long
  iImage         As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32" _
                                  (lpBI As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                  (ByVal pidList As Long, _
                                  ByVal lpBuffer As String) As Long
'Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                  (ByVal lpString1 As String, ByVal _
                                  lpString2 As String) As Long

Private Declare PtrSafe Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As TChooseColor) As LongPtr
#End If

Private Enum eOpenFile
  OFN_READONLY = &H1
  OFN_OVERWRITEPROMPT = &H2
  OFN_HIDEREADONLY = &H4
  OFN_NOCHANGEDIR = &H8
  OFN_SHOWHELP = &H10
  OFN_ENABLEHOOK = &H20
  OFN_ENABLETEMPLATE = &H40
  OFN_ENABLETEMPLATEHANDLE = &H80
  OFN_NOVALIDATE = &H100
  OFN_ALLOWMULTISELECT = &H200
  OFN_EXTENSIONDIFFERENT = &H400
  OFN_PATHMUSTEXIST = &H800
  OFN_FILEMUSTEXIST = &H1000
  OFN_CREATEPROMPT = &H2000
  OFN_SHAREAWARE = &H4000
  OFN_NOREADONLYRETURN = &H8000
  OFN_NOTESTFILECREATE = &H10000
  OFN_NONETWORKBUTTON = &H20000
  OFN_NOLONGNAMES = &H40000
  OFN_EXPLORER = &H80000
  OFN_NODEREFERENCELINKS = &H100000
  OFN_LONGNAMES = &H200000
End Enum

Private Const MAX_PATH As Integer = 260
Private Const MAX_FILE As Integer = 260
Private Const BIF_RETURNONLYFSDIRS  As Long = 1&
Private Const BIF_DONTGOBELOWDOMAIN As Long = 2&

Private Const CC_SOLIDCOLOR = &H80
Private Type TChooseColor
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    rgbResult As Long
    lpCustColors As LongPtr
    Flags As Long
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
End Type
Private malUdColors(1 To 16) As Long

#If Win64 Then
Public Function BrowseForFolder(ByVal hWnd As LongPtr, ByVal szTitle As String) As String
  Dim lpIDList As LongPtr
#Else
Public Function BrowseForFolder(ByVal hWnd As Long, ByVal szTitle As String) As String
  Dim lpIDList As Long
#End If
  Dim sBuffer As String
  Dim tBrowseInfo As BROWSEINFO
  
  With tBrowseInfo
    .hwndOwner = hWnd
    '.lpszTitle = lstrcat(szTitle, "")
    .lpszTitle = szTitle
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
  End With
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    BrowseForFolder = sBuffer
  End If
End Function

#If Win64 Then
Function VBGetOpenFileName(Filename As String, _
                           Optional FileTitle As String, _
                           Optional FileMustExist As Boolean = True, _
                           Optional MultiSelect As Boolean = False, _
                           Optional ReadOnly As Boolean = False, _
                           Optional HideReadOnly As Boolean = False, _
                           Optional filter As String = "All (*.*)| *.*", _
                           Optional FilterIndex As Long = 1, _
                           Optional InitDir As String, _
                           Optional DlgTitle As String, _
                           Optional DefaultExt As String, _
                           Optional Owner As Long = -1, _
                           Optional Flags As Long = 0) As Boolean
#Else
Function VBGetOpenFileName(Filename As String, _
                           Optional FileTitle As String, _
                           Optional FileMustExist As Boolean = True, _
                           Optional MultiSelect As Boolean = False, _
                           Optional ReadOnly As Boolean = False, _
                           Optional HideReadOnly As Boolean = False, _
                           Optional filter As String = "All (*.*)| *.*", _
                           Optional FilterIndex As Long = 1, _
                           Optional InitDir As String, _
                           Optional DlgTitle As String, _
                           Optional DefaultExt As String, _
                           Optional Owner As LongPtr = -1, _
                           Optional Flags As Long = 0) As Boolean
#End If
  Dim opfile As OPENFILENAME, s As String, afFlags As Long
  With opfile
  #If Win64 Then
    .lStructSize = LenB(opfile)
  #Else
    .lStructSize = Len(opfile)
  #End If
    
    ' Add in specific flags and strip out non-VB flags
    .Flags = (-FileMustExist * OFN_FILEMUSTEXIST) Or _
             (-MultiSelect * OFN_ALLOWMULTISELECT) Or _
             (-ReadOnly * OFN_READONLY) Or _
             (-HideReadOnly * OFN_HIDEREADONLY) Or _
             (Flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
    ' Owner can take handle of owning window
    If Owner <> -1 Then .hwndOwner = Owner
    ' InitDir can take initial directory string
    .lpstrInitialDir = InitDir
    ' DefaultExt can take default extension
    .lpstrDefExt = DefaultExt
    ' DlgTitle can take dialog box title
    .lpstrTitle = DlgTitle
    
    ' To make Windows-style filter, replace | and : with nulls
    Dim ch As String, i As Integer
    For i = 1 To Len(filter)
      ch = Mid$(filter, i, 1)
      If ch = "|" Or ch = ":" Then
        s = s & vbNullChar
      Else
        s = s & ch
      End If
    Next
    ' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = FilterIndex

    ' Pad file and file title buffers to maximum path
    s = Filename & String$(MAX_PATH - Len(Filename), 0)
    .lpstrFile = s
    .nMaxFile = MAX_PATH
    s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = MAX_FILE
    ' All other fields set to zero
    
    If GetOpenFileName(opfile) Then
      VBGetOpenFileName = True
      Filename = CtoVB(.lpstrFile)
      FileTitle = CtoVB(.lpstrFileTitle)
      Flags = .Flags
      ' Return the filter index
      FilterIndex = .nFilterIndex
      ' Look up the filter the user selected and return that
      filter = FilterLookup(.lpstrFilter, FilterIndex)
      If (.Flags And OFN_READONLY) Then ReadOnly = True
    Else
      VBGetOpenFileName = False
      Filename = ""
      FileTitle = ""
      Flags = 0
      FilterIndex = -1
      filter = ""
    End If
  End With
End Function

Function VBGetSaveFileName(Filename As String, _
                           Optional FileTitle As String, _
                           Optional OverWritePrompt As Boolean = True, _
                           Optional filter As String = "All (*.*)| *.*", _
                           Optional FilterIndex As Long = 1, _
                           Optional InitDir As String, _
                           Optional DlgTitle As String, _
                           Optional DefaultExt As String, _
                           Optional Owner As Long = -1, _
                           Optional Flags As Long) As Boolean
            
  Dim opfile As OPENFILENAME, s As String
  With opfile
    .lStructSize = Len(opfile)
    
    ' Add in specific flags and strip out non-VB flags
    .Flags = (-OverWritePrompt * OFN_OVERWRITEPROMPT) Or _
             OFN_HIDEREADONLY Or _
             (Flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
    ' Owner can take handle of owning window
    If Owner <> -1 Then .hwndOwner = Owner
    ' InitDir can take initial directory string
    .lpstrInitialDir = InitDir
    ' DefaultExt can take default extension
    .lpstrDefExt = DefaultExt
    ' DlgTitle can take dialog box title
    .lpstrTitle = DlgTitle
    
    ' Make new filter with bars (|) replacing nulls and double null at end
    Dim ch As String, i As Integer
    For i = 1 To Len(filter)
      ch = Mid$(filter, i, 1)
      If ch = "|" Or ch = ":" Then
        s = s & vbNullChar
      Else
        s = s & ch
      End If
    Next
    ' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = FilterIndex

    ' Pad file and file title buffers to maximum path
    s = Filename & String$(MAX_PATH - Len(Filename), 0)
    .lpstrFile = s
    .nMaxFile = MAX_PATH
    s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = MAX_FILE
    ' All other fields zero
    
    If GetSaveFileName(opfile) Then
      VBGetSaveFileName = True
      Filename = CtoVB(.lpstrFile)
      FileTitle = CtoVB(.lpstrFileTitle)
      Flags = .Flags
      ' Return the filter index
      FilterIndex = .nFilterIndex
      ' Look up the filter the user selected and return that
      filter = FilterLookup(.lpstrFilter, FilterIndex)
    Else
      VBGetSaveFileName = False
      Filename = ""
      FileTitle = ""
      Flags = 0
      FilterIndex = 0
      filter = ""
    End If
  End With
End Function

Private Function FilterLookup(ByVal sFilters As String, ByVal iCur As Long) As String
  Dim iStart As Long, iEnd As Long, s As String
  iStart = 1
  If sFilters = "" Then Exit Function
  Do
    ' Cut out both parts marked by null character
    iEnd = InStr(iStart, sFilters, vbNullChar)
    If iEnd = 0 Then Exit Function
    iEnd = InStr(iEnd + 1, sFilters, vbNullChar)
    If iEnd Then
      s = Mid$(sFilters, iStart, iEnd - iStart)
    Else
      s = Mid$(sFilters, iStart)
    End If
    iStart = iEnd + 1
    If iCur = 1 Then
      FilterLookup = s
      Exit Function
    End If
    iCur = iCur - 1
  Loop While iCur
End Function

Function VBGetFileTitle(sFile As String) As String
  Dim sFileTitle As String, cFileTitle As Integer

  cFileTitle = MAX_PATH
  sFileTitle = String$(MAX_PATH, 0)
  cFileTitle = GetFileTitle(sFile, sFileTitle, MAX_PATH)
  If cFileTitle Then
    VBGetFileTitle = ""
  Else
    VBGetFileTitle = left$(sFileTitle, InStr(sFileTitle, vbNullChar) - 1)
  End If
End Function

'Inspired by https://social.msdn.microsoft.com/Forums/office/en-US/3b95d3bf-1ecb-4c8e-a946-cd87f6194cf9/color-picker-for-access-project?forum=accessdev
Public Function PickColor(ByVal hWnd As LongPtr, ByRef plRetColor As Long, Optional ByVal plDefColor As Long = 0&) As Boolean
  Dim lRet    As LongPtr
  Dim tCC     As TChooseColor
  If hWnd = 0 Then hWnd = Application.hWndAccessApp
  With tCC
    .lStructSize = LenB(tCC)
    .hwndOwner = hWnd
    .Flags = CC_SOLIDCOLOR
    .lpCustColors = VarPtr(malUdColors(1))
    lRet = ChooseColor(tCC)
    If lRet Then
      plRetColor = CLng(.rgbResult)
      PickColor = True
    End If
  End With
End Function


