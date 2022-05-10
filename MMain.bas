Attribute VB_Name = "MMain"
Option Compare Database
Option Explicit

'When      |Version |Who|What
'----------+--------+---+-----------------------------------------------------------
'24.03.2020|02.00.00|FFO| Beginning versioning at v2 02.00.00
'          |        |   | - Adding transparency and background image features
'          |        |   | - Adding managed dirty state to CColorPalette
'04.04.2020|02.00.01|FFO| CBitmap:SaveConsoleAsBitmap() fully functional, saves
'          |        |   | only the part of the console that's the drawing, with
'          |        |   | pixel precision, thanks to consoul 01.01.00
'          |        |   | Enhancements & introducing progress bar and event locking
'04.04.2020|02.00.03|FFO| Lot of fixes, adding goto page on character map.
'          |        |   |
'          |        |   |
'          |        |   |

Public Const APP_NAME As String = "AsciiPaint"
Public Const APP_VERSION As String = "02.01.11"

Public Const IMSG_APP_ALREADY_RUNNING As String = "This application is already running."

Public AppIniFile     As New CIniFile

Public Const INISECTION_CANVAS          As String = "Canvas"
Public Const INIKEY_DEFCANVASROWS       As String = "DefaultRows"
Public Const INIKEY_DEFCANVASCOLS       As String = "DefaultCols"
Public Const INIKEY_DEFFORECOLOR        As String = "DefaultForeColor"
Public Const INIKEY_DEFBACKCOLOR        As String = "DefaultBackColor"
Public Const INIKEY_DEFFONTNAME         As String = "DefaultFontName"
Public Const INIKEY_DEFFONTSIZE         As String = "DefaultFontSize"
Public Const INIKEY_CANVASWINDOWWIDTH   As String = "WindowWidth"
Public Const INIKEY_CANVASWINDOWHEIGHT  As String = "WindowHeight"
Public Const INIKEY_CANVASTRANSPARENT   As String = "Transparent"
Public Const INIKEY_CANVASTRANSMODE     As String = "TransparencyMode"
Public Const INIKEY_CANVASTRANSCOLOR    As String = "TransparencyColor"
Public Const INIKEY_CANVASALPHAPCT      As String = "AlphaChannelPct"
Public Const INIKEY_CANVASBKGNDIMAGE    As String = "BackgoundImage"
Public Const INIKEY_CURSORFOLLOWMOUSE   As String = "CursorFollowsMouse"
Public Const INIKEY_SHOWGRID            As String = "ShowGrid"
Public Const INIKEY_FONTFAMILY          As String = "FontFamily"
'V02.00.01
Public Const INIOPT_LASTLOADPATH        As String = "LastLoadPath"
Public Const INIOPT_LASTSAVEPATH        As String = "LastSavePath"
Public Const INIOPT_CLIPLASTLOADPATH    As String = "ClipboardLastLoadPath"
Public Const INIOPT_PALLASTLOADPATH     As String = "PaletteLastLoadPath"
Public Const INIOPT_PALLASTSAVEPATH     As String = "PaletteLastSavePath"
Public Const INIOPT_LASTBKIMAGEPATH     As String = "LastBkImagePath"
Public Const INIOPT_CGLASTSAVEPATH      As String = "CodeGenLastSavePath"
Public Const INIOPT_CGLASTLOADPATH      As String = "CodeGenLastLoadPath"
Public Const INIOPT_LASTINSERTPATH      As String = "LastInsertPath"
Public Const INIOPT_LASTLIBBROWSEPATH   As String = "LastLibBrowsePath"
Public Const INIOPT_LASTLIBXTRACTPATH   As String = "LastLibExtractPath"
Public Const INIOPT_LASTLIBIMPORTPATH   As String = "LastLibImportPath"
'librarian
Public Const INIOPT_CONLIBFONTNAME      As String = "LibFontName"
Public Const INIOPT_CONLIBFONTSIZE      As String = "LibFontSize"
Public Const INIOPT_CONLIBMAXCAP        As String = "LibMaxCap"
Public Const INIOPT_CONLIBFORECOLOR     As String = "LibForeColor"
Public Const INIOPT_CONLIBBACKCOLOR     As String = "LibBackColor"

Public Const INISECTION_DEBUGCONSOLE    As String = "DebugConsole"
Public Const INIKEY_DISPLAYLINECT       As String = "DisplayLineCount"
Public Const INIKEY_CHARSPERLINE        As String = "CharsPerLine"

'A simple mechanism for avoiding reentrancy with control events
Private mfEventsLocked   As Boolean

Public Function ConnectIniFile() As Boolean
  Dim sIniFile      As String
  Dim sTargetFile1  As String
  Dim fh            As Integer
  
  On Error Resume Next
  
  With AppIniFile
    If Len(.Filename) = 0 Then
      'Search order:
      '1. User hive \ APP_NAME directory
      '2. Current (progdb) directory
      'If not found, create in User hive first
      sIniFile = APP_NAME & ".ini"
      sTargetFile1 = CombinePath(GetSpecialFolder(Application.hWndAccessApp, CSIDL_APPDATA), sIniFile)
      If Not ExistFile(sTargetFile1) Then
        fh = FreeFile
        Open sTargetFile1 For Output Access Write Lock Write As #fh
        Close fh
      End If
      .Filename = sTargetFile1
    End If
  End With
  
  ConnectIniFile = True
End Function

Public Function IsOffice64() As Boolean
  #If Win64 Then
    IsOffice64 = True
  #Else
    IsOffice64 = False
  #End If
End Function

Public Function FindConsoulLibrary() As Boolean
  'First look into current project path, or in \bin subdirectory
  Dim sLookupPath     As String
  Dim sFilename       As String
  Dim fOK             As Boolean
  Dim fInPath         As Boolean
  
  On Error GoTo FindConsoulLibrary_Err
  
#If Win64 Then
  Const CONSOUL_DLL_NAME As String = "consoul_010203_64.dll"
#Else
  Const CONSOUL_DLL_NAME As String = "consoul_010203_32.dll"
#End If
  'Look into project path
  sLookupPath = CurrentProject.Path
  sFilename = CombinePath(sLookupPath, CONSOUL_DLL_NAME)
  fOK = ExistFile(sFilename)
  If Not fOK Then
    'Look into \bin subdir
    sLookupPath = CombinePath(sLookupPath, "bin")
    sFilename = CombinePath(sLookupPath, CONSOUL_DLL_NAME)
    fOK = ExistFile(sFilename)
    If Not fOK Then
      Dim iPathCt     As Integer
      Dim asPath()    As String
      Dim i           As Integer
      iPathCt = SplitString(asPath(), Environ$("PATH"), ";")
      For i = 1 To iPathCt
        sLookupPath = asPath(i)
        sFilename = CombinePath(sLookupPath, CONSOUL_DLL_NAME)
        fOK = ExistFile(sFilename)
        If fOK Then
          fInPath = True
          Exit For
        End If
      Next i
    End If
  End If

  If fOK Then
    If Not fInPath Then
      Debug.Print "Found, chdir to : "; sLookupPath
      ChDir sLookupPath
    Else
      Debug.Print "Found in PATH : "; sLookupPath
    End If
  Else
    ShowUFError "Cannot find " & CONSOUL_DLL_NAME, CONSOUL_DLL_NAME & " couldn't be found in application directory, \bin subdirectory or directories in PATH"
    'ShowError "FindConsoulLibrary", -1&, "Cannot find " & CONSOUL_DLL_NAME
  End If
  
FindConsoulLibrary_Exit:
  FindConsoulLibrary = fOK
  Exit Function

FindConsoulLibrary_Err:
  ShowUFError "Searching for " & CONSOUL_DLL_NAME & " failed.", Err.Description
  'ShowError "FindConsoulLibrary", Err.Number, Err.Description
  Resume FindConsoulLibrary_Exit
End Function

'TestDDELink
'Return 1 if this database (psDatabaseName) is already opened by another MSAccess instance.
'Found on the Internet, source lost, please tweet us @idevinfo if you can attribute it.
Function TestDDELink(ByVal psDatabaseName As String) As Integer
  Dim vDDEChannel As Long
  On Error Resume Next
  Application.SetOption "Ignore DDE Requests", True
  ' Open a channel between database instances
  vDDEChannel = DDEInitiate("MSAccess", psDatabaseName)
  'If the database is NOT already opened, then it will not be possible to create the DDE channel
  If Err Then
    TestDDELink = 0
  Else
    TestDDELink = 1
    DDETerminate vDDEChannel
    DDETerminateAll
  End If
  Application.SetOption ("Ignore DDE Requests"), False
End Function

'For some database properties, that only the jet database engine can modify
Public Sub SetDBProperty(ByRef pDB As DAO.Database, ByVal sPropName As String, ByVal vPropValue As Variant)
  Dim db    As DAO.Database
  Dim prp   As DAO.Property
  
  On Error Resume Next
  pDB.Properties(sPropName) = vPropValue
  If Err.Number = 3270 Then
    Set prp = pDB.CreateProperty(sPropName, dbText, vPropValue)
    pDB.Properties.Append prp
  End If
End Sub

Private Function IsAdminMode() As Boolean
  'Only by code thou shalt change that
  IsAdminMode = False
End Function

Public Sub DisableAccessFeatures()
  Dim fAdminMode As Boolean
  
  fAdminMode = IsAdminMode()
  If Not fAdminMode Then
    DoCmd.RunCommand acCmdWindowHide
  Else
    MsgBox "WARNING" & vbCrLf & vbCrLf & _
           "This application is being executed in admin mode.", _
           vbCritical
  End If
  
  'Do some user / admin setup for the main Access Window
  On Error Resume Next
  SetDBProperty CurrentDb(), "StartupShowDBWindow", fAdminMode
  SetDBProperty CurrentDb(), "AllowBuiltInToolbars", fAdminMode
  SetDBProperty CurrentDb(), "StartUpShowStatusBar", fAdminMode
  SetDBProperty CurrentDb(), "AllowShortcutMenus", fAdminMode
  SetDBProperty CurrentDb(), "AllowToolbarChanges", fAdminMode
  SetDBProperty CurrentDb(), "AllowSpecialKeys", fAdminMode
  SetDBProperty CurrentDb(), "AllowBypassKey", fAdminMode
  SetDBProperty CurrentDb(), "AllowFullMenus", fAdminMode
  SetDBProperty CurrentDb(), "AllowBreakIntoCode", fAdminMode
  
  DoCmd.SetDisplayedCategories fAdminMode
  DoCmd.LockNavigationPane (Not fAdminMode)
  If Not fAdminMode Then
    DoCmd.ShowToolbar "Menu Bar", acToolbarNo
    DoCmd.ShowToolbar "Ribbon", acToolbarNo
  End If
  
End Sub

Public Sub DisablePrint()
'Unfortunately doesn't work
'TODO: keep looking for a solution
  Dim MainBar As CommandBar
  Dim subBar  As CommandBarPopup
  
  On Error Resume Next
  Set MainBar = CommandBars("Menu Bar")
  For Each subBar In MainBar.Controls
    subBar.Enabled = False
    subBar.Visible = False
  Next
  
  Set subBar = Nothing
  Set MainBar = Nothing
End Sub

Public Function Main() As Boolean
  On Error Resume Next
  
  'Test for single instance application
  If TestDDELink(Application.CurrentDb.Name) Then
    MsgBox IMSG_APP_ALREADY_RUNNING, vbInformation
    DoCmd.Quit acQuitSaveNone
  End If
  
  If Not FindConsoulLibrary() Then
    DoCmd.Quit acQuitSaveNone
    Exit Function
  End If
  
  DisableAccessFeatures
  
  Call ConnectIniFile
  DoCmd.OpenForm GetCanvasFormName(), acNormal
End Function

Public Function Max(ByVal V1 As Variant, ByVal V2 As Variant) As Variant
  If V1 > V2 Then
    Max = V1
  Else
    Max = V2
  End If
End Function

Public Function Min(ByVal V1 As Variant, ByVal V2 As Variant) As Variant
  If V1 < V2 Then
    Min = V1
  Else
    Min = V2
  End If
End Function

'V02.00.00 Light red used to alert of status in UI
'          Used at this version for the type text textbox on the canvas
'          and on the charmap background to indicate unsync with canvas font
Public Function GetAlertBackgroundColor() As Long
  GetAlertBackgroundColor = RGB(250, 160, 160)
End Function

Public Function GetCanvasFormName() As String
  GetCanvasFormName = "Canvas"
End Function

Public Function GetCharMapFormName() As String
  GetCharMapFormName = "CharacterMap"
End Function

Public Function GetPaletteFormName() As String
  GetPaletteFormName = "ColorPalette"
End Function

Public Function GetGenCodeFormName() As String
  GetGenCodeFormName = "GenerateCode"
End Function

Public Function GetLibrarianFormName() As String
  GetLibrarianFormName = "Librarian"
End Function

Public Function LOCK_EVENTS() As Boolean
  If mfEventsLocked Then Exit Function
  mfEventsLocked = True
  LOCK_EVENTS = True
End Function

Public Sub UNLOCK_EVENTS()
  mfEventsLocked = False
End Sub

Public Function EVENTS_LOCKED() As Boolean
  EVENTS_LOCKED = mfEventsLocked
End Function


