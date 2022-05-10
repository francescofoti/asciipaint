VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Canvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements ICsMouseEventSink
Implements IProgressIndicator
Implements ICsWmPaintEventSink
Implements IMessageReceiver

Private moCanvas        As CCanvas

'The console objects
' moConCanvas    Displays the drawing "canvas" console
Private moConCanvas     As CConsoul

Private mfTransparent         As Boolean
Private msTransImageFilename  As String
Private mfShowGrid            As Boolean
Private mfCursorFollowMouse   As Boolean
Private mlLastTypedCharCode   As Long

Private Const DEFAULT_CANVASROWS  As Integer = 40
Private Const DEFAULT_CANVASCOLS  As Integer = 60
Private Const DEFAULT_FONTNAME As String = "Courier New"
Private Const DEFAULT_FONTSIZE As Integer = 12

'Helper class to manipulate a console like a grid of characters
Private WithEvents mgrdCanvas      As CConsoleGrid
Attribute mgrdCanvas.VB_VarHelpID = -1
Private mgrdClipboard   As CConsoleGrid

'toggles for displaying colors as hex/dec or rgb in status bar
Private miCurBkColDisp  As Integer
Private miCurFgColDisp  As Integer

Private mrcClient   As RECT 'compute on Form_Resize and used in RepositionConsouls

Private mbDispCharCodeToggle  As Byte 'status bar character code

'V02.00.02 Consoul window as selection area indicator
Private moConSel  As CConsoul
Private miConSelLeft    As Integer
Private miConSelTop     As Integer
Private miConSelWidth   As Integer
Private miConSelHeight  As Integer

Private miFontFamily    As Integer

' Selection anchors
Private Enum eSelAnchorEdge
  eSelBegin
  eSelEnd
End Enum
Private Const HWND_TOP = 0
Private Const HWND_BOTTOM = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_SHOWWINDOW = &H40

'V02.00.01 New progress indicator, via a consoul progress bar code control
Private moProgressBar   As New CCProgressBar  'automatic is easier here
Private moConProgress   As CConsoul
Private mfProgressRunning  As Boolean
Private Const MAX_PRGB_CAPTION_LEN As Integer = 30

Private Const SUBDIR_IMAGES As String = "images"

Private moRibbon As CRibbon
Attribute moRibbon.VB_VarHelpID = -1

Private Const RIBBONTAB_FILE      As String = "File"
Private Const RIBBONTAB_DISPLAY   As String = "Display"
Private Const RIBBONTAB_TEXT      As String = "Text"
Private Const RIBBONTAB_CLIPBOARD As String = "Clipboard"
Private Const RIBBONTAB_LINECOL   As String = "LineCol"
Private Const RIBBONTAB_DEVELOPER As String = "Developer"

Private mfNoKeyPreview As Boolean

Private mconDebug         As CConsoul
Private mfConDebugVisible As Boolean

Private miLastDbgRow      As Integer
Private miDbgCharsPerLine As Integer
Private miDbgWindowLineCt As Integer

Private Sub DisableFormKeyPreview()
  mfNoKeyPreview = True
End Sub

Private Sub EnableFormKeyPreview()
  mfNoKeyPreview = False
End Sub

'**********************************************************************
'
' Progress indicator in top/right corner of canvas window methods
'
'**********************************************************************

Private Sub RefreshProgressBar()
  moProgressBar.EnsureVisible = True
  moConProgress.ShowWindow True
  moProgressBar.Render moConProgress
  moConProgress.RefreshWindow
  DoEvents
End Sub

Private Sub ShowProgressBar(ByVal pfShow As Boolean)
  On Error Resume Next
  moConProgress.ShowWindow pfShow
  If pfShow Then
    RepositionProgressBar
    RefreshProgressBar
  End If
End Sub

Private Sub cboFontFamily_Click()
  miFontFamily = Me.cboFontFamily
  ReloadFonts
  AppIniFile.Section = INISECTION_CANVAS
  AppIniFile.SetString INIKEY_FONTFAMILY, CStr(miFontFamily)
End Sub

Private Sub cboFontFamily_Enter()
  DisableFormKeyPreview
End Sub

Private Sub cboFontFamily_Exit(Cancel As Integer)
  EnableFormKeyPreview
End Sub

Private Sub cboFontName_Enter()
  DisableFormKeyPreview
End Sub

Private Sub cboFontName_Exit(Cancel As Integer)
  EnableFormKeyPreview
End Sub

Private Sub cmdBColor_Click()
  If Not LockUI() Then Exit Sub
  SetBackColor Me.cmdBColor.BackColor
  CreateView
  UpdateView
  UpdateCanvasColorsIndics
  UnlockUI
End Sub

Private Sub cmdFColor_Click()
  If Not LockUI() Then Exit Sub
  SetForeColor Me.cmdFColor.BackColor
  CreateView
  UpdateView
  UpdateCanvasColorsIndics
  UnlockUI
End Sub

Private Sub UpdateFreehandIndicator()
  If mfCursorFollowMouse Then
    Me.lblFreehand.Caption = "Freehand ON"
    Me.lblFreehand.ForeColor = RGB(0, 128, 0)
  Else
    Me.lblFreehand.Caption = "Freehand OFF"
    Me.lblFreehand.ForeColor = vbGrayText
  End If
End Sub

Private Sub cmdFreeHand_Click()
  mfCursorFollowMouse = Not mfCursorFollowMouse
  AppIniFile.SetString INIKEY_CURSORFOLLOWMOUSE, Abs(CInt(mfCursorFollowMouse)) & ""
  UpdateFreehandIndicator
End Sub

Private Sub cmdLibrarian_Click()
  On Error Resume Next
  DoCmd.OpenForm GetLibrarianFormName(), acNormal
End Sub

Private Sub cmdReplaceColor_Click()
  If Not LockUI() Then Exit Sub
  ReplaceColorAction
  UnlockUI
End Sub

Private Sub cmdResetLineAttr_Click()
  If Not LockUI() Then Exit Sub
  DoHourglass True
  Me.txtLinePaddingTop = "0"
  Me.txtLinePaddingBottom = "0"
  Me.txtLineSpacingTop = "0"
  Me.txtLineSpacingBottom = "0"
  Me.chkAutoAdjustWidth = False
  Me.chkFragmentText = False
  moConCanvas.AutoAdjustWidth = False
  mgrdCanvas.SetFragmentText False, 0
  moConCanvas.LineSpacing(elsTop) = 0
  moConCanvas.LineSpacing(elsBottom) = 0
  moConCanvas.LinePadding(elsTop) = 0
  moConCanvas.LinePadding(elsBottom) = 0
  RefreshView True
  moConCanvas.RefreshWindow
  Focus2Canvas
  DoHourglass False
  UnlockUI
End Sub

Private Sub cmdShowHideGrid_Click()
  ShowConsoleGrid Not mfShowGrid
End Sub

Private Sub cmdSwitchColors_Click()
  If Not LockUI() Then Exit Sub
  SwitchColorsAction
  UnlockUI
End Sub

Private Sub cmdTextBackcolor_Click()
  If Not LockUI() Then Exit Sub
  Dim iRow      As Integer
  Dim iCol      As Integer
  
  moConCanvas.GetCaretPos iRow, iCol
  BackColorAction iRow, iCol
  UnlockUI
End Sub

Private Sub cmdTextForecolor_Click()
  If Not LockUI() Then Exit Sub
  Dim iRow      As Integer
  Dim iCol      As Integer
  
  moConCanvas.GetCaretPos iRow, iCol
  ForeColorAction iRow, iCol
  UnlockUI
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  If mfNoKeyPreview Then
    Exit Sub
  End If
  If IsUILocked() Then
    Exit Sub
  End If
  If Shift And acAltMask Then
    Exit Sub
  End If
  
  Dim fProcessed As Boolean
  
  fProcessed = True
  Select Case KeyCode
  Case vbKeyRight
    OnCursorKeyRight
  Case vbKeyLeft
    OnCursorKeyLeft
  Case vbKeyDown
    OnCursorKeyDown
  Case vbKeyUp
    OnCursorKeyUp
  Case vbKeyHome
    OnKeyHome Shift
  Case vbKeyEnd
    OnKeyEnd Shift
  Case vbKeyPageUp
    OnKeyPageUp
  Case vbKeyPageDown
    OnKeyPageDown
  Case vbKeyF2  'get text into edit
    OnKeyF2 Shift
  Case vbKeyF3  'find next
    OnKeyF3 Shift
  Case vbKeyF   'forecolor / search char
    OnKeyF Shift
  Case vbKeyH   'replace char
    OnKeyH Shift
  Case vbKeyT   'type text
    OnKeyT Shift
  Case vbKeyB, vbKeyI, vbKeyU, vbKeyK, vbKeyE, vbKeyR 'attributes, B also for backcolor
    OnKeyAttributes KeyCode, Shift
  Case vbKeyS
    cmdSaveFileAs_Click
  Case vbKeyN
    cmdNewFile_Click
  Case vbKeyA   'Select all
    OnKeyA Shift
  Case vbKeyD   'Deselect all
    OnKeyD Shift
  Case vbKeyC   'Copy
    OnKeyC Shift
  Case vbKeyV   'Paste
    OnKeyV Shift
  Case vbKeyM   'Mark selection
    OnKeyM Shift
  Case vbKeyDelete
    OnKeyDelete Shift
  Case vbKeySpace
    OnKeySpace Shift
  Case vbKeyG
    OnKeyG Shift
  Case vbKeyF12
    OnKeyF12 Shift
  'Case vbKeyZ
    'Debug.Print "Current canvas memory used: " & mgrdCanvas.GetMemoryFootprint()
  Case Else
    fProcessed = False
  End Select
  
  If fProcessed Then
    KeyCode = 0
  End If
End Sub

Private Sub Form_MouseWheel(ByVal Page As Boolean, ByVal Count As Long)
  On Error Resume Next
  Dim hWnd                  As LongPtr
  Dim ptMouse               As POINTAPI
  Dim i                     As Long
  Dim hWndDebug             As LongPtr
  
  If moConCanvas Is Nothing Then Exit Sub
  If Not mconDebug Is Nothing Then
    hWndDebug = mconDebug.hWnd
  End If
  
  apiGetCursorPos ptMouse
  hWnd = apiWindowFromPoint(ptMouse.x, ptMouse.y)
  If hWnd = moConCanvas.hWnd Then
    For i = 1 To Abs(Count)
      If Count < 0 Then
        moConCanvas.ScrollUp
      Else
        moConCanvas.ScrollDown
      End If
    Next
  ElseIf hWnd = hWndDebug Then
    For i = 1 To Abs(Count)
      If Count < 0 Then
        mconDebug.ScrollUp
      Else
        mconDebug.ScrollDown
      End If
    Next
  ElseIf hWnd = moRibbon.TabConsole.hWnd Then
    moRibbon.OnMousWheel Count
  End If
End Sub

'**********************************************************************
'
' IProgressIndicator implementation
'
'**********************************************************************

Private Sub IProgressIndicator_BeginProgress(ByVal psMessage As String)
  mfProgressRunning = True
  'moConProgress.Clear
  moProgressBar.Max = 100
  moProgressBar.Value = 0
  moProgressBar.Caption = psMessage
  moConProgress.ShowWindow True
  RepositionProgressBar
  RefreshProgressBar
  RepositionProgressBar
  Me.TimerInterval = 2000  'Every 2 seconds
End Sub

Private Property Get IProgressIndicator_Console() As CConsoul
  IProgressIndicator_Console = Nothing
End Property

Private Sub IProgressIndicator_EndProgress()
  If Not mfProgressRunning Then 'if we called w/o BeginProgress
    Me.TimerInterval = 2000  'Every 2 seconds
  End If
  mfProgressRunning = False 'will be hidden by the form timer routine
  moProgressBar.Value = moProgressBar.Max
  RefreshProgressBar
End Sub

Public Property Get IIProgressIndicator() As IProgressIndicator
  Set IIProgressIndicator = Me
End Property

Private Sub IProgressIndicator_SetCaption(ByVal psCaption As String)
  On Error Resume Next
  If Len(psCaption) <= MAX_PRGB_CAPTION_LEN Then
    psCaption = Space$(MAX_PRGB_CAPTION_LEN - Len(psCaption)) & psCaption
  Else
    psCaption = left$(psCaption, MAX_PRGB_CAPTION_LEN - 3) & "..."
  End If
  moProgressBar.Caption = psCaption
  RefreshProgressBar
End Sub

Private Sub IProgressIndicator_SetMax(ByVal plMax As Long)
  On Error Resume Next
  moProgressBar.Max = plMax
  moProgressBar.Value = 0
  RefreshProgressBar
End Sub

Private Sub IProgressIndicator_SetText(ByVal psText As String)
  On Error Resume Next
  moProgressBar.Text = psText
  RefreshProgressBar
End Sub

Private Sub IProgressIndicator_SetValue(ByVal plValue As Long)
  On Error Resume Next
  moProgressBar.Value = plValue
  RefreshProgressBar
End Sub

Private Sub IProgressIndicator_ShowProgressIndicator(ByVal pfShow As Boolean)
  ShowProgressBar pfShow
End Sub

'**********************************************************************
'
' Support private methods
'
'**********************************************************************

Private Sub DoHourglass(ByVal pfShow As Boolean)
  On Error Resume Next
  DoCmd.Hourglass pfShow
  IProgressIndicator_EndProgress
End Sub

Private Function InputChar(ByVal psMsg As String, ByVal psTitle As String, ByVal psDefault As String) As String
  Dim sInput      As String
  Dim iLen        As Integer
  Dim sCode       As String
  
  sInput = InputBox$(psMsg, psTitle, psDefault)
  iLen = Len(sInput)
  If iLen > 0 Then
    InputChar = sInput
  End If
End Function

Private Function ParseChar(ByVal psCharOrHexedCode As String) As String
  If StrComp(left$(psCharOrHexedCode, 2), "&H", vbTextCompare) = 0 Then
    psCharOrHexedCode = Right$(psCharOrHexedCode, Len(psCharOrHexedCode) - 2)
    ParseChar = ChrW$(Val("&H" & psCharOrHexedCode))
  Else
    ParseChar = left$(psCharOrHexedCode, 1)
  End If
End Function

'**********************************************************************
'
' Onxxx keybord commands
'
'**********************************************************************

'Caller MUST adjust column with Canvas MarginWidth
Private Sub OnFindChar(ByVal piRow As Integer, ByVal piCol As Integer)
  Dim sChar     As String
  Dim sMsg      As String
  Dim iRow      As Integer
  Dim iCol      As Integer
  
  Static sLastInput As String
  
  sMsg = "Find character:" & vbCrLf & vbCrLf & "(Use prefix '&H' to input and hexadecimal unicode character code, like &H20 for space)"
  sChar = InputChar(sMsg, "Search for...", sLastInput)
  If Len(sChar) > 0 Then
    sLastInput = sChar
    'search
    moConCanvas.GetCaretPos iRow, iCol
    If mgrdCanvas.FindFirstChar(sChar, iRow, iCol) Then
      moConCanvas.SetCaretPos mgrdCanvas.LastFindRow, mgrdCanvas.LastFindCol
    Else
      Beep
    End If
  End If
End Sub

Private Sub SwitchColorsAction()
  Static sLastFindColor     As String
  Static sLastReplaceColor  As String
  Static sLastCanvasScope   As String
  Static sLastColorScope    As String
  
  Dim iRow          As Integer
  Dim iCol          As Integer
  Dim iEndRow       As Integer
  Dim iEndCol       As Integer
  Dim sMsg          As String
  Dim sTip          As String
  Dim iChoice       As Integer
  Dim lReplCt       As Long
  Dim fOK           As Boolean
  Const DIALOG_TITLE As String = "Switch colors scope"
  
  moConCanvas.GetCaretPos iRow, iCol
  
  'Switch colors in which (canvas) scope ?
  sMsg = ""
  sMsg = sMsg & "Switch Forecolor and Backcolor" & vbCrLf & vbCrLf
  sMsg = sMsg & "Where ?" & vbCrLf & vbCrLf
  sMsg = sMsg & "1 - In the current line (#" & iRow & ") starting at column #" & iRow & vbCrLf
  sMsg = sMsg & "2 - In the whole current line (#" & iRow & ") & vbcrlf"
  sMsg = sMsg & "3 - To the end of the canvas from current position (row " & iRow & ", col " & iCol & vbCrLf
  sMsg = sMsg & "4 - In the whole canvas"
  If HasSelection() Then
    sMsg = sMsg & vbCrLf & "5 - In the current selection ("
    iChoice = IntChooseBox(sMsg, DIALOG_TITLE, sLastCanvasScope, 5)
  Else
    iChoice = IntChooseBox(sMsg, DIALOG_TITLE, sLastCanvasScope, 4)
  End If
  If iChoice = 0 Then Exit Sub
  
  'we've got all we need to do the replace
  Me.IIProgressIndicator.BeginProgress "Switching "
  sLastCanvasScope = iChoice & ""
  Select Case iChoice
  Case 1
    iEndRow = iRow
    iEndCol = mgrdCanvas.Cols
  Case 2
    iCol = 1
    iEndRow = iRow
    iEndCol = mgrdCanvas.Cols
  Case 3
    iEndRow = mgrdCanvas.Rows
    iEndCol = mgrdCanvas.Cols
  Case 4
    iRow = 1
    iCol = 1
    iEndRow = mgrdCanvas.Rows
    iEndCol = mgrdCanvas.Cols
  Case 5
    If Not HasSelection() Then
      MsgBox "The selection is empty", vbInformation
      Exit Sub
    End If
    iRow = moCanvas.SelStartRow
    iCol = moCanvas.SelStartCol
    iEndRow = moCanvas.SelEndRow
    iEndCol = moCanvas.SelEndCol
  End Select
  DoHourglass True
  fOK = mgrdCanvas.SwitchColors(iRow, iCol, iEndRow, iEndCol, moConCanvas, lReplCt)
  DoHourglass False
  If fOK Then
    DoHourglass True
    RepositionProgressBar
    mgrdCanvas.LoadConsole moConCanvas, Me
    RefreshView
    DoHourglass False
    MsgBox lReplCt & " switches done", vbInformation
  Else
    ShowUFError "An error occured while switching colors", mgrdCanvas.LastErrDesc
  End If
  Me.IIProgressIndicator.EndProgress
End Sub

Private Sub ReplaceColorAction()
  Static sLastFindColor     As String
  Static sLastReplaceColor  As String
  Static sLastCanvasScope   As String
  Static sLastColorScope    As String
  
  Dim iRow          As Integer
  Dim iCol          As Integer
  Dim iEndRow       As Integer
  Dim iEndCol       As Integer
  Dim lFindColor    As Long
  Dim lReplaceColor As Long
  Dim sMsg          As String
  Dim sTip          As String
  Dim eReplaceScope As eReplaceCharScope
  Dim eColorScope   As eReplaceColorScope
  Dim iChoice       As Integer
  Dim lReplCt       As Long
  Dim fOK           As Boolean
  Const DIALOG_TITLE As String = "Switch colors scope"
  
  moConCanvas.GetCaretPos iRow, iCol
  
  sTip = "(Use prefix '&H' to input an hexadecimal color value (example: &H336699)"
  sMsg = "Color value to find and replace:" & vbCrLf & vbCrLf & sTip
  If Not InputColor(lFindColor, sMsg, "Search/Replace color", sLastFindColor) Then Exit Sub
  sLastFindColor = "&H" & ColorToHex(lFindColor)
  sMsg = "Replace color (" & sLastFindColor & ") with color:" & vbCrLf & vbCrLf & sTip
  If Not InputColor(lReplaceColor, sMsg, "Search/Replace color", sLastReplaceColor) Then Exit Sub
  sLastReplaceColor = "&H" & ColorToHex(lReplaceColor)
  'replace what (color scope)
  sMsg = ""
  sMsg = sMsg & "Replace " & sLastFindColor & " with " & sLastReplaceColor & vbCrLf & vbCrLf
  sMsg = sMsg & "For ?" & vbCrLf & vbCrLf
  sMsg = sMsg & "1 - Forecolor" & vbCrLf
  sMsg = sMsg & "2 - Backcolor" & vbCrLf
  sMsg = sMsg & "3 - Both" & vbCrLf
  iChoice = IntChooseBox(sMsg, "Replace what colors", sLastColorScope, 3)
  If iChoice = 0 Then Exit Sub
  sLastColorScope = CStr(iChoice)
  Select Case iChoice
  Case 1
    eColorScope = eForeColor
  Case 2
    eColorScope = eBackColor
  Case 3
    eColorScope = eBoth
  End Select
  
  'replace in which (canvas) scope ?
  sMsg = ""
  sMsg = sMsg & "Replace " & sLastFindColor & " with " & sLastReplaceColor & vbCrLf & vbCrLf
  sMsg = sMsg & "Where ?" & vbCrLf & vbCrLf
  sMsg = sMsg & "1 - In the current line (#" & iRow & ") starting at column #" & iRow & vbCrLf
  sMsg = sMsg & "2 - In the whole current line (#" & iRow & ") & vbcrlf"
  sMsg = sMsg & "3 - To the end of the canvas from current position (row " & iRow & ", col " & iCol & vbCrLf
  sMsg = sMsg & "4 - In the whole canvas"
  If HasSelection() Then
    sMsg = sMsg & vbCrLf & "5 - In the current selection ("
    iChoice = IntChooseBox(sMsg, DIALOG_TITLE, sLastCanvasScope, 5)
  Else
    iChoice = IntChooseBox(sMsg, DIALOG_TITLE, sLastCanvasScope, 4)
  End If
  If iChoice = 0 Then Exit Sub
  
  'we've got all we need to do the replace
  Me.IIProgressIndicator.BeginProgress "Replacing "
  sLastCanvasScope = iChoice & ""
  Select Case iChoice
  Case 1
    eReplaceScope = eToEndOfLine
    iEndRow = iRow
    iEndCol = mgrdCanvas.Cols
  Case 2
    eReplaceScope = eFullLine
    iCol = 1
    iEndRow = iRow
    iEndCol = mgrdCanvas.Cols
  Case 3
    eReplaceScope = eToEndOfText
    iEndRow = mgrdCanvas.Rows
    iEndCol = mgrdCanvas.Cols
  Case 4
    eReplaceScope = eFullText
    iRow = 1
    iCol = 1
    iEndRow = mgrdCanvas.Rows
    iEndCol = mgrdCanvas.Cols
  Case 5
    eReplaceScope = eSelection
    If Not HasSelection() Then
      MsgBox "The selection is empty", vbInformation
      Exit Sub
    End If
    iRow = moCanvas.SelStartRow
    iCol = moCanvas.SelStartCol
    iEndRow = moCanvas.SelEndRow
    iEndCol = moCanvas.SelEndCol
  End Select
  DoHourglass True
  fOK = mgrdCanvas.ReplaceColors(lFindColor, lReplaceColor, iRow, iCol, iEndRow, iEndCol, eColorScope, lReplCt)
  DoHourglass False
  If fOK Then
    DoHourglass True
    RepositionProgressBar
    mgrdCanvas.LoadConsole moConCanvas, Me
    RefreshView
    DoHourglass False
    MsgBox lReplCt & " colors were replaced", vbInformation
  Else
    ShowUFError "An error occured while replacing colors", mgrdCanvas.LastErrDesc
  End If
  Me.IIProgressIndicator.EndProgress
End Sub

'Caller MUST adjust column with Canvas MarginWidth
Private Sub OnReplaceChar(ByVal piRow As Integer, ByVal piCol As Integer)
  Dim sChar     As String
  Dim sRepl     As String
  Dim sMsg      As String
  Dim sTip      As String
  Dim iChoice   As Integer
  Dim eScope    As eReplaceCharScope
  Dim fOK       As Boolean
  Dim lReplCt   As Long
  Dim iRow      As Integer
  Dim iCol      As Integer
  
  Static sLastInputSearch   As String
  Static sLastInputReplace  As String
  Static sLastInputScope    As String
  
  moConCanvas.GetCaretPos iRow, iCol
  
  sTip = "(Use prefix '&H' to input an hexadecimal unicode character code, like &H20 for space)"
  sMsg = "Find and replace character:" & vbCrLf & vbCrLf & sTip
  sChar = InputChar(sMsg, "Search for...", sLastInputSearch)
  If Len(sChar) > 0 Then
    sLastInputSearch = sChar
    'replace
    sMsg = "Replace '" & ParseChar(sChar) & "' (&H" & Hex$(AscW(sChar)) & ") with :" & vbCrLf & vbCrLf & sTip
    sRepl = InputChar(sMsg, "Replace with...", sLastInputReplace)
    If Len(sRepl) > 0 Then
      sLastInputReplace = sRepl
      If Len(sLastInputScope) = 0 Then sLastInputScope = "1"
      'replace in which scope ?
      sMsg = ""
      sMsg = sMsg & "Replace '" & ParseChar(sChar) & "' (&H" & Hex$(AscW(sChar)) & ") with '"
      sMsg = sMsg & ParseChar(sRepl) & "' (&H" & Hex$(AscW(ParseChar(sRepl))) & ")" & vbCrLf & vbCrLf
      sMsg = sMsg & "In which scope ?" & vbCrLf & vbCrLf
      sMsg = sMsg & "1 - In the current line (#" & iRow & ") starting at column #" & iRow & vbCrLf
      sMsg = sMsg & "2 - In the whole current line (#" & iRow & ") & vbcrlf"
      sMsg = sMsg & "3 - To the end of the canvas from current position (row " & iRow & ", col " & iCol & vbCrLf
      sMsg = sMsg & "4 - In the whole canvas"
      'do the replace
      iChoice = IntChooseBox(sMsg, "Character replacement scope", sLastInputScope, 4)
      If iChoice > 0 Then
        Me.IIProgressIndicator.BeginProgress "Replacing "
        sLastInputScope = iChoice & ""
        Select Case iChoice
        Case 1
          eScope = eToEndOfLine
        Case 2
          eScope = eFullLine
        Case 3
          eScope = eToEndOfText
        Case 4
          eScope = eFullText
        End Select
        DoHourglass True
        fOK = mgrdCanvas.ReplaceChar(ParseChar(sChar), ParseChar(sRepl), iRow, iCol, eScope, lReplCt)
        DoHourglass False
        If fOK Then
          DoHourglass True
          RepositionProgressBar
          mgrdCanvas.LoadConsole moConCanvas, Me
          RefreshView
          DoHourglass False
          MsgBox lReplCt & " characters were replaced", vbInformation
        Else
          ShowUFError "An error occured while replacing", mgrdCanvas.LastErrDesc
        End If
        Me.IIProgressIndicator.EndProgress
      End If
    End If
  End If

End Sub

'**********************************************************************
'
' Handle cursor movement with arrow key
'
'**********************************************************************

Private Sub OnCursorKeyLeft(Optional ByVal Shift As Integer = 0)
  Dim iRow    As Integer
  Dim iCol    As Integer
  moConCanvas.GetCaretPos iRow, iCol
  If iCol > 1 Then
    iCol = iCol - 1
    moConCanvas.SetCaretPos iRow, iCol
    OnCaretPosChange
  End If
End Sub

Private Sub OnCursorKeyRight(Optional ByVal Shift As Integer = 0)
  Dim iRow    As Integer
  Dim iCol    As Integer
  moConCanvas.GetCaretPos iRow, iCol
  If iCol < mgrdCanvas.Cols Then
    iCol = iCol + 1
    moConCanvas.SetCaretPos iRow, iCol
    OnCaretPosChange
  End If
End Sub

Private Sub OnCursorKeyUp(Optional ByVal Shift As Integer = 0)
  Dim iRow    As Integer
  Dim iCol    As Integer
  moConCanvas.GetCaretPos iRow, iCol
  If iRow > 1 Then
    iRow = iRow - 1
    If iRow < moConCanvas.TopLine Then
      moConCanvas.ScrollPageUp
    End If
    moConCanvas.SetCaretPos iRow, iCol
    OnCaretPosChange
  End If
End Sub

Private Sub OnCursorKeyDown(Optional ByVal Shift As Integer = 0)
  Dim iRow    As Integer
  Dim iCol    As Integer
  moConCanvas.GetCaretPos iRow, iCol
  If iRow < mgrdCanvas.Rows Then
    iRow = iRow + 1
    If iRow > (moConCanvas.TopLine + moConCanvas.MaxVisibleRows - 1) Then
      moConCanvas.ScrollPageDown
    End If
    moConCanvas.SetCaretPos iRow, iCol
    OnCaretPosChange
  End If
End Sub

Private Sub OnKeyHome(Optional ByVal Shift As Integer = 0)
  Dim iRow    As Integer
  Dim iCol    As Integer
  moConCanvas.GetCaretPos iRow, iCol
  
  If Shift = 0 Then
    moConCanvas.SetCaretPos iRow, 1
  ElseIf Shift = acCtrlMask Then 'ctrl
    moConCanvas.SetCaretPos 1, 1
    moConCanvas.ScrollTop
  End If
  
  OnCaretPosChange
End Sub

Private Sub OnKeyEnd(Optional ByVal Shift As Integer = 0)
  Dim iRow    As Integer
  Dim iCol    As Integer
  moConCanvas.GetCaretPos iRow, iCol

  If Shift = 0 Then
    moConCanvas.SetCaretPos iRow, mgrdCanvas.Cols
  ElseIf Shift = acCtrlMask Then
    moConCanvas.SetCaretPos mgrdCanvas.Rows, mgrdCanvas.Cols
    moConCanvas.ScrollBottom
  End If
  
  OnCaretPosChange
End Sub

Private Sub OnKeyPageUp(Optional ByVal Shift As Integer = 0)
  Dim iRow    As Integer
  Dim iCol    As Integer
  moConCanvas.GetCaretPos iRow, iCol

  If Shift = 0 Then
    moConCanvas.ScrollPageUp
    moConCanvas.SetCaretPos moConCanvas.TopLine, iCol
  End If
  
  OnCaretPosChange
End Sub

Private Sub OnKeyPageDown(Optional ByVal Shift As Integer = 0)
  Dim iRow    As Integer
  Dim iCol    As Integer
  moConCanvas.GetCaretPos iRow, iCol

  If Shift = 0 Then
    moConCanvas.ScrollPageDown
    moConCanvas.SetCaretPos moConCanvas.TopLine, iCol
  End If
  
  OnCaretPosChange
End Sub

'Get text from position into edit textbox (txtTypeText)
Private Sub OnKeyF2(Optional ByVal Shift As Integer = 0)
  Dim iRow    As Integer
  Dim iCol    As Integer
  Dim sText   As String
  
  On Error Resume Next
  
  moConCanvas.GetCaretPos iRow, iCol
  
  If Shift = 0 Then
    sText = Trim$(mgrdCanvas.TextAt(iRow, iCol))
    If Len(sText) > 0 Then
      Me.txtTypeText = sText
      Me.lblTypeTextLen.Visible = False
      Focus2Canvas
    End If
  End If
End Sub

Private Sub OnKeyF3(Optional ByVal Shift As Integer = 0)
  If mgrdCanvas.FindNextChar() Then
    moConCanvas.SetCaretPos mgrdCanvas.LastFindRow, mgrdCanvas.LastFindCol
  End If
End Sub

'[F] set forecolor action
Private Sub OnKeyF(Optional ByVal Shift As Integer = 0)
  Dim iRow    As Integer
  Dim iCol    As Integer
  moConCanvas.GetCaretPos iRow, iCol
  
  If Shift = 0 Then
    ForeColorAction iRow, iCol
    OnCaretPosChange
  ElseIf Shift = acCtrlMask Then  'ctrl + f = find char
    OnFindChar iRow, iCol
  End If
End Sub

'ctrl+h = replace char
Private Sub OnKeyH(Optional ByVal Shift As Integer = 0)
  Dim iRow    As Integer
  Dim iCol    As Integer
  moConCanvas.GetCaretPos iRow, iCol
  
  If Shift = acCtrlMask Then  'ctrl + h = replace char
    OnReplaceChar iRow, iCol
    OnCaretPosChange
    Focus2Canvas
  End If
End Sub

'[B] set backcolor action
Private Sub OnKeyB(Optional ByVal Shift As Integer = 0)
  Dim iRow    As Integer
  Dim iCol    As Integer
  moConCanvas.GetCaretPos iRow, iCol
  
  If Shift = 0 Then
    BackColorAction iRow, iCol
  End If
  
  OnCaretPosChange
End Sub

'[T] Type text
Private Sub OnKeyT(Optional ByVal Shift As Integer = 0)
  Dim iRow    As Integer
  Dim iCol    As Integer
  moConCanvas.GetCaretPos iRow, iCol
  
  If Shift = 0 Then
    cmdTypeText_Click
  End If
  
  RenderLine iRow
  OnCaretPosChange
End Sub

'[ctrl]+[a] select all
Private Sub OnKeyA(Optional ByVal Shift As Integer = 0)
  If Shift = acCtrlMask Then
    moCanvas.SelStartRow = 1
    moCanvas.SelEndRow = mgrdCanvas.Rows
    moCanvas.SelStartCol = 1
    moCanvas.SelEndCol = mgrdCanvas.Cols
    ShowSelection
  End If
End Sub

'[ctrl]+[d] deselect all
Private Sub OnKeyD(Optional ByVal Shift As Integer = 0)
  If Shift = acCtrlMask Then
    moCanvas.SelStartRow = 0
    moCanvas.SelEndRow = 0
    moCanvas.SelStartCol = 0
    moCanvas.SelEndCol = 0
    ShowSelection
  End If
End Sub

'[ctrl]+[c] copy
Private Sub OnKeyC(Optional ByVal Shift As Integer = 0)
  If Not Me.cmdCopy.Enabled Then Exit Sub
  If Shift = acCtrlMask Then cmdCopy_Click
End Sub

'[ctrl]+[v] pase
Private Sub OnKeyV(Optional ByVal Shift As Integer = 0)
  If Not Me.cmdPaste.Enabled Then Exit Sub
  If Shift = acCtrlMask Then cmdPaste_Click
End Sub

'[M] = set selection start
'[Shift]+[M] = set selection end
Private Sub OnKeyM(Optional ByVal Shift As Integer = 0)
  If Shift = acShiftMask Then
    cmdSelEnd_Click
  ElseIf Shift = 0 Then
    cmdSelStart_Click
  End If
End Sub

Private Sub TypeCharAt(ByVal plCharCode As Long, ByVal piRow As Integer, ByVal piCol As Integer)
  Dim sChar As String
  mlLastTypedCharCode = plCharCode
  sChar = ChrW$(plCharCode)
  mgrdCanvas.CharAt(piRow, piCol) = sChar
  RenderLine piRow
  Focus2Canvas
End Sub

Private Sub TypeChar(ByVal plCharCode As Long, Optional ByVal pfNoMove As Boolean = False)
  Dim sChar As String
  Dim iRow  As Integer
  Dim iCol  As Integer
  
  'Insert the character at current caret position
  If plCharCode <> 0 Then
    moConCanvas.GetCaretPos iRow, iCol
    TypeCharAt plCharCode, iRow, iCol
    If Not pfNoMove Then OnCursorKeyRight
  End If
End Sub

Private Sub ReTypeChar()
  If mlLastTypedCharCode <> 0& Then
    TypeChar mlLastTypedCharCode
  End If
End Sub

'[Space] = Type character selected in character palette
Private Sub OnKeySpace(Optional ByVal Shift As Integer = 0)
  ReTypeChar
End Sub

'[G] = Show/Hide grid
Private Sub OnKeyG(Optional ByVal Shift As Integer = 0)
  ShowConsoleGrid Not mfShowGrid
End Sub

'[F12] = Show/Hide debug console
Private Sub OnKeyF12(Optional ByVal Shift As Integer = 0)
  cmdDebugConsole_Click
End Sub

'[Ctrl]+[Delete] = Clear button
'[Delete] = delete char at cursor position
'[Shift]+[Delete] = Delete char and attribs at cursor position
Private Sub OnKeyDelete(Optional ByVal Shift As Integer = 0)
  Dim iRow      As Integer
  Dim iCol      As Integer
  
  moConCanvas.GetCaretPos iRow, iCol
  
  If Shift = acCtrlMask Then
    cmdClear_Click
  ElseIf Shift = acShiftMask Then
    mgrdCanvas.ClearChar iRow, iCol
    RenderLine iRow
  Else
    mgrdCanvas.ClearChar iRow, iCol, False, False
    RenderLine iRow
  End If
  OnCaretPosChange
End Sub

'[Ctrl][b]/[i]/[u]/[s]/[n]/[r] set attribute
'[Shift][b]/[i]/[u]/[s]/[n]/[r] unset attribute
Private Sub OnKeyAttributes(ByVal piKeyCode As Integer, Optional ByVal Shift As Integer = 0)
  Dim iRow      As Integer
  Dim iCol      As Integer
  Dim iAttribs  As Integer
  
  moConCanvas.GetCaretPos iRow, iCol
  
  If Shift = acCtrlMask Then 'ctrl = toggle on
    iAttribs = mgrdCanvas.CharAttribs(iRow, iCol)
    Select Case piKeyCode
    Case vbKeyB
      If iAttribs And ATTRIB_BOLDON Then
        iAttribs = iAttribs And Not ATTRIB_BOLDON
      Else
        iAttribs = iAttribs Or ATTRIB_BOLDON
      End If
    Case vbKeyI
      If iAttribs And ATTRIB_ITALICON Then
        iAttribs = iAttribs And Not ATTRIB_ITALICON
      Else
        iAttribs = iAttribs Or ATTRIB_ITALICON
      End If
    Case vbKeyU
      If iAttribs And ATTRIB_UNDLON Then
        iAttribs = iAttribs And Not ATTRIB_UNDLON
      Else
        iAttribs = iAttribs Or ATTRIB_UNDLON
      End If
    Case vbKeyK
      If iAttribs And ATTRIB_STRIKEON Then
        iAttribs = iAttribs And Not ATTRIB_STRIKEON
      Else
        iAttribs = iAttribs Or ATTRIB_STRIKEON
      End If
    Case vbKeyE
      If iAttribs And ATTRIB_INVERSEON Then
        iAttribs = iAttribs And Not ATTRIB_INVERSEON
      Else
        iAttribs = iAttribs Or ATTRIB_INVERSEON
      End If
    Case vbKeyR
      If iAttribs And ATTRIB_RESET Then
        iAttribs = iAttribs And Not ATTRIB_RESET
      Else
        iAttribs = iAttribs Or ATTRIB_RESET
      End If
    End Select
    mgrdCanvas.CharAttribs(iRow, iCol) = iAttribs
    RenderLine iRow
  ElseIf Shift = acShiftMask Then ' shift = toggle off
    iAttribs = mgrdCanvas.CharAttribs(iRow, iCol)
    Select Case piKeyCode
    Case vbKeyB
      If iAttribs And ATTRIB_BOLDOFF Then
        iAttribs = iAttribs And Not ATTRIB_BOLDOFF
      Else
        iAttribs = iAttribs Or ATTRIB_BOLDOFF
      End If
    Case vbKeyI
      If iAttribs And ATTRIB_ITALICOFF Then
        iAttribs = iAttribs And Not ATTRIB_ITALICOFF
      Else
        iAttribs = iAttribs Or ATTRIB_ITALICOFF
      End If
    Case vbKeyU
      If iAttribs And ATTRIB_UNDLOFF Then
        iAttribs = iAttribs And Not ATTRIB_UNDLOFF
      Else
        iAttribs = iAttribs Or ATTRIB_UNDLOFF
      End If
    Case vbKeyK
      If iAttribs And ATTRIB_STRIKEOFF Then
        iAttribs = iAttribs And Not ATTRIB_STRIKEOFF
      Else
        iAttribs = iAttribs Or ATTRIB_STRIKEOFF
      End If
    Case vbKeyE
      If iAttribs And ATTRIB_INVERSEOFF Then
        iAttribs = iAttribs And Not ATTRIB_INVERSEOFF
      Else
        iAttribs = iAttribs Or ATTRIB_INVERSEOFF
      End If
    End Select
    mgrdCanvas.CharAttribs(iRow, iCol) = iAttribs
    RenderLine iRow
  Else
    If piKeyCode = vbKeyB Then
      OnKeyB
    End If
  End If
  
  OnCaretPosChange
End Sub

Private Sub cboFontName_AfterUpdate()
  Dim sFontName   As String
  On Error Resume Next
  If Not LockUI() Then Exit Sub
  
  sFontName = Trim$(cboFontName & "")
  If Len(sFontName) = 0 Then
    UnlockUI
    Exit Sub
  End If
  FontSetSelectedFont sFontName
  moCanvas.FontName = sFontName
  Me.txtTypeText.FontName = sFontName
  CreateView
  UpdateView
  On Error Resume Next
  Focus2Canvas
  MessageManager.Broadcast Me.Name, MSGTOPIC_FONTNAMECHANGED, sFontName
  
  UnlockUI
End Sub

'**********************************************************************
'
' controls events
'
'**********************************************************************

Private Sub chkTransparency_Click()
  On Error Resume Next
  If Not LockUI() Then Exit Sub
  'Debug.Print "Call SetCanvasTransparency(" & CBool(Me.chkTransparency) & "," & meTransMode & ")"
  SetCanvasTransparency CBool(Me.chkTransparency), moCanvas.TransparencyMode
  UnlockUI
End Sub

Private Sub cmdBrowseBkgndImage_Click()
  Dim iChoice     As Integer
  Dim fLoaded     As Boolean
  Dim sFilename   As String
  Dim sInitialDir As String
  
  On Error GoTo cmdBrowseBkgndImage_Click_Err
  If Not LockUI() Then Exit Sub
  
  AppIniFile.GetOption INIOPT_LASTBKIMAGEPATH, sInitialDir, ""
  sInitialDir = Trim$(sInitialDir)
  If Len(sInitialDir) = 0 Then
    sInitialDir = CombinePath _
      ( _
        CombinePath _
        ( _
          GetSpecialFolder(Application.hWndAccessApp, CSIDL_PERSONAL), _
          APP_NAME _
        ), _
        SUBDIR_IMAGES _
      )
  End If
  If Not ExistDir(sInitialDir) Then
    If Not CreatePath(sInitialDir) Then
      sInitialDir = GetSpecialFolder(Application.hWndAccessApp, CSIDL_PERSONAL)
    End If
  End If
  
  With Application.FileDialog(msoFileDialogFilePicker)
    .Title = "Load compatible file"
    .InitialFileName = NormalizePath(sInitialDir)
    '"AsciiPaint (*.ascp)|*.ascp|Text files (*.txt, *.asc, *.vt100)|*.txt;*.asc;*.vt100"
    .Filters.Clear
    .Filters.Add "Image files", "*.jpg;*.jpeg;*.jpe;*.gif;*.bmp;*.png;*.rle;*.rib;*.exif;*.tiff;*.tif;*.icon;*.wmf;*.emf"
    .FilterIndex = 1
    iChoice = .Show()
    If iChoice <> 0 Then
      sFilename = .SelectedItems(1)
      AppIniFile.SetOption INIOPT_LASTBKIMAGEPATH, (StripFileName(sFilename))
      DoHourglass True
      SetCanvasBackgroundImageFromFile sFilename
    End If
  End With

cmdBrowseBkgndImage_Click_Exit:
  UnlockUI
  DoHourglass False
  Focus2Canvas
  Exit Sub

cmdBrowseBkgndImage_Click_Err:
  ShowUFError "An error occured while loading the picture", Err.Description
  Resume cmdBrowseBkgndImage_Click_Exit
End Sub

Private Sub cmdCanvasSize_Click()
  Dim oDlg      As New CCanvasSizeDialog
  
  On Error Resume Next
  If Not LockUI() Then Exit Sub
  
  oDlg.Rows = mgrdCanvas.Rows
  oDlg.Cols = mgrdCanvas.Cols
  If oDlg.IIDialog.ShowDialog(True) Then
    If Not oDlg.IIDialog.Cancelled Then
      If mgrdCanvas.Resize(oDlg.Rows, oDlg.Cols) Then
        CreateView
        UpdateView
      Else
        ShowUFError "Resizing the canvas failed", mgrdCanvas.LastErrDesc
      End If
    End If
  End If
  Focus2Canvas
  
  UnlockUI
End Sub

Private Function GetSelectedColor() As Variant
  Dim rowParams     As New CRow
  rowParams.AddCol "color", Null, 0, 0
  MessageManager.Broadcast Me.Name, MSGTOPIC_GETSELCOLOR, rowParams, GetPaletteFormName()
  GetSelectedColor = rowParams("color")
End Function

Private Sub cmdCharBackColor_Click()
  If Not LockUI() Then Exit Sub
  
  Dim vColor As Variant
  vColor = GetSelectedColor()
  If Not IsNull(vColor) Then
    SetPenBkColor vColor
  End If
  Focus2Canvas
  
  UnlockUI
End Sub

Private Sub cmdCharForeColor_Click()
  If Not LockUI() Then Exit Sub
  
  Dim vColor As Variant
  vColor = GetSelectedColor()
  If Not IsNull(vColor) Then
    SetPenForeColor vColor
  End If
  Focus2Canvas
  
  UnlockUI
End Sub

Private Sub cmdClear_Click()
  Dim iRow        As Integer
  Dim iCol        As Integer
  Dim fClrText    As Boolean
  Dim fClrColors  As Boolean
  Dim fClrAttribs As Boolean
  Dim i           As Long
  
  On Error Resume Next
  If Not LockUI() Then Exit Sub
  
  moConCanvas.GetCaretPos iRow, iCol
  
  i = Me.cboDelWhat.ListIndex
  Select Case i
  Case 0
    fClrText = True
  Case 1
    fClrColors = True
  Case 2
    fClrAttribs = True
  Case 3
    fClrText = True
    fClrColors = True
    fClrAttribs = True
  End Select
  
  DoHourglass True
  'Line or column
  If Me.optLineCol = 0 Then 'line
    'Clear from/to cursor position
    If Me.OptClearDir = 0 Then
      'Clear to cursor
      mgrdCanvas.ClearLine iRow, iCol, eToCursorPosition, fClrText, fClrColors, fClrAttribs
    Else
      'Clear from cursor to end
      mgrdCanvas.ClearLine iRow, iCol, eFromCursorPosition, fClrText, fClrColors, fClrAttribs
    End If
    RenderLine iRow
  Else
    'Clear from/to cursor position
    If Me.OptClearDir = 0 Then
      'Clear to cursor
      mgrdCanvas.ClearCol iRow, iCol, eToCursorPosition, fClrText, fClrColors, fClrAttribs
      For i = 1 To iRow
        RenderLine i
      Next i
    Else
      'Clear from cursor to end
      mgrdCanvas.ClearCol iRow, iCol, eFromCursorPosition, fClrText, fClrColors, fClrAttribs
      For i = iRow To mgrdCanvas.Rows
        RenderLine i
      Next i
    End If
  End If
  UpdateStbCharCode
  DoHourglass False
  Focus2Canvas
  
  UnlockUI
End Sub

Private Sub cmdClearBkgndImage_Click()
  On Error Resume Next
  If Not LockUI() Then Exit Sub
  
  DoHourglass True
  SetCanvasBackgroundImageFromFile ""
  DoHourglass False
  
  UnlockUI
End Sub

Private Sub cmdClearSel_Click()
  If Not LockUI() Then Exit Sub
  
  OnKeyD 2  '"send" ctrl+d
  Focus2Canvas
  
  UnlockUI
End Sub

Private Sub cmdConBackColor_Click()
  If Not LockUI() Then Exit Sub
  
  Dim vColor As Variant
  vColor = GetSelectedColor()
  If Not IsNull(vColor) Then
    Me.cmdBColor.BackColor = vColor
    Me.cmdBColor.HoverColor = vColor
    Me.cmdBColor.PressedColor = vColor
  End If
  
  UnlockUI
End Sub

Private Sub cmdConForeColor_Click()
  If Not LockUI() Then Exit Sub
  
  Dim vColor As Variant
  vColor = GetSelectedColor()
  If Not IsNull(vColor) Then
    Me.cmdFColor.BackColor = vColor
    Me.cmdFColor.HoverColor = vColor
    Me.cmdFColor.PressedColor = vColor
  End If
  
  UnlockUI
End Sub

Private Sub cmdDeleteLineCol_Click()
  On Error Resume Next
  Dim iRow      As Integer
  Dim iCol      As Integer
  
  Dim fClrText    As Boolean
  Dim fClrColors  As Boolean
  Dim fClrAttribs As Boolean
  Dim i           As Long
  
  If Not LockUI() Then Exit Sub
  
  i = Me.cboDelWhat.ListIndex
  Select Case i
  Case 0
    fClrText = True
  Case 1
    fClrColors = True
  Case 2
    fClrAttribs = True
  Case 3
    fClrText = True
    fClrColors = True
    fClrAttribs = True
  End Select
  
  DoHourglass True
  moConCanvas.GetCaretPos iRow, iCol
  If Me.optLineCol = 0 Then 'line
    mgrdCanvas.ShiftUp fClrText, fClrColors, fClrAttribs, iRow
  Else
    mgrdCanvas.ShiftLeft fClrText, fClrColors, fClrAttribs, iCol
  End If
  RepositionProgressBar
  mgrdCanvas.LoadConsole moConCanvas, Me
  RefreshView
  DoHourglass False
  Focus2Canvas
  
  UnlockUI
End Sub

Private Sub cmdGenCode_Click()
  If Not LockUI() Then Exit Sub
  
  Dim oDlg      As New CGenerateCodeDialog
  
  'oDlg... =
  'mgrdCanvas
  Set oDlg.ConsoleGrid = mgrdCanvas
  If oDlg.IIDialog.ShowDialog(True) Then
    If Not oDlg.IIDialog.Cancelled Then
      'do something after the modal display
    End If
  End If
  Set oDlg.ConsoleGrid = Nothing
  Focus2Canvas
  
  UnlockUI
End Sub

Private Sub cmdInsertLineCol_Click()
  On Error Resume Next
  Dim iRow      As Integer
  Dim iCol      As Integer
  
  Dim fClrText    As Boolean
  Dim fClrColors  As Boolean
  Dim fClrAttribs As Boolean
  Dim i           As Long
  
  If Not LockUI() Then Exit Sub
  
  i = Me.cboDelWhat.ListIndex
  Select Case i
  Case 0
    fClrText = True
  Case 1
    fClrColors = True
  Case 2
    fClrAttribs = True
  Case 3
    fClrText = True
    fClrColors = True
    fClrAttribs = True
  End Select
  
  DoHourglass True
  moConCanvas.GetCaretPos iRow, iCol
  If Me.optLineCol = 0 Then 'line
    mgrdCanvas.ShiftDown fClrText, fClrColors, fClrAttribs, iRow
  Else
    mgrdCanvas.ShiftRight fClrText, fClrColors, fClrAttribs, iCol
  End If
  RepositionProgressBar
  mgrdCanvas.LoadConsole moConCanvas, Me
  RefreshView
  DoHourglass False
  Focus2Canvas
  
  UnlockUI
End Sub

Private Sub cmdShiftDown_Click()
  Dim fClrText    As Boolean
  Dim fClrColors  As Boolean
  Dim fClrAttribs As Boolean
  Dim i           As Long
  
  If Not LockUI() Then Exit Sub
  
  i = Me.cboDelWhat.ListIndex
  Select Case i
  Case 0
    fClrText = True
  Case 1
    fClrColors = True
  Case 2
    fClrAttribs = True
  Case 3
    fClrText = True
    fClrColors = True
    fClrAttribs = True
  End Select
  
  On Error Resume Next
  DoHourglass True
  mgrdCanvas.ShiftDown fClrText, fClrColors, fClrAttribs
  RepositionProgressBar
  mgrdCanvas.LoadConsole moConCanvas, Me
  RefreshView
  DoHourglass False
  Focus2Canvas
  
  UnlockUI
End Sub

Private Sub cmdShiftLeft_Click()
  Dim fClrText    As Boolean
  Dim fClrColors  As Boolean
  Dim fClrAttribs As Boolean
  Dim i           As Long
  
  If Not LockUI() Then Exit Sub
  
  i = Me.cboDelWhat.ListIndex
  Select Case i
  Case 0
    fClrText = True
  Case 1
    fClrColors = True
  Case 2
    fClrAttribs = True
  Case 3
    fClrText = True
    fClrColors = True
    fClrAttribs = True
  End Select
  
  On Error Resume Next
  DoHourglass True
  mgrdCanvas.ShiftLeft fClrText, fClrColors, fClrAttribs
  RepositionProgressBar
  mgrdCanvas.LoadConsole moConCanvas, Me
  RefreshView
  DoHourglass False
  Focus2Canvas
  
  UnlockUI
End Sub

Private Sub cmdShiftRight_Click()
  Dim fClrText    As Boolean
  Dim fClrColors  As Boolean
  Dim fClrAttribs As Boolean
  Dim i           As Long
  
  If Not LockUI() Then Exit Sub
  
  i = Me.cboDelWhat.ListIndex
  Select Case i
  Case 0
    fClrText = True
  Case 1
    fClrColors = True
  Case 2
    fClrAttribs = True
  Case 3
    fClrText = True
    fClrColors = True
    fClrAttribs = True
  End Select
  
  On Error Resume Next
  DoHourglass True
  mgrdCanvas.ShiftRight fClrText, fClrColors, fClrAttribs
  RepositionProgressBar
  mgrdCanvas.LoadConsole moConCanvas, Me
  RefreshView
  DoHourglass False
  Focus2Canvas

  UnlockUI
End Sub

Private Sub cmdShiftUp_Click()
  Dim fClrText    As Boolean
  Dim fClrColors  As Boolean
  Dim fClrAttribs As Boolean
  Dim i           As Long
  
  If Not LockUI() Then Exit Sub
  
  i = Me.cboDelWhat.ListIndex
  Select Case i
  Case 0
    fClrText = True
  Case 1
    fClrColors = True
  Case 2
    fClrAttribs = True
  Case 3
    fClrText = True
    fClrColors = True
    fClrAttribs = True
  End Select
  
  On Error Resume Next
  DoHourglass True
  mgrdCanvas.ShiftUp fClrText, fClrColors, fClrAttribs
  RepositionProgressBar
  mgrdCanvas.LoadConsole moConCanvas, Me
  RefreshView
  DoHourglass False

  UnlockUI
End Sub

Private Sub cmdTransColor_Click()
  If Not LockUI() Then Exit Sub
  
  Dim vColor As Variant
  vColor = GetSelectedColor()
  If Not IsNull(vColor) Then
    DoHourglass True
    SetCanvasTransparentColor vColor
    SetCanvasTransparency mfTransparent, moCanvas.TransparencyMode  'just redo
    SetTransColor moCanvas.TransparentColor
    DoHourglass False
  End If
  
  UnlockUI
End Sub

Private Sub cmdTypeText_Click()
  Dim sText       As String
  Dim iRow        As Integer
  Dim iCol        As Integer
  
  If Not LockUI() Then Exit Sub
  On Error Resume Next
  
  sText = Me.txtTypeText & ""
  If Len(sText) > 0 Then
    moConCanvas.GetCaretPos iRow, iCol
    mgrdCanvas.TextAt(iRow, iCol) = sText
    RenderLine iRow
    UpdateToolbar
    UpdateStbCharCode
  End If
  Focus2Canvas

  UnlockUI
End Sub

Private Sub RenderLine(ByVal piRow As Integer)
  moConCanvas.SetLine piRow, mgrdCanvas.GetLineVT100(piRow)
  moConCanvas.RedrawLine piRow
End Sub

Private Sub Form_Load()
  On Error Resume Next
  
  'We use a timer to wait for the form to be visible and have its final size,
  'then we initialize it (See Form_Timer).
  'We could do the initialisation here, but we can close the form in the resize
  'event, if something goes bad.
  Me.TimerInterval = 200
  
  Set moCanvas = New CCanvas
  Set mgrdCanvas = moCanvas.ConsoleGrid
End Sub

Private Sub Form_Timer()
  If Not Me.Visible Then Exit Sub 'wait for next timer event
  
  On Error Resume Next
  Static fDoneOnce As Boolean
  
  If Not fDoneOnce Then
    InitNewCanvas moCanvas
    If InitForm() Then
      fDoneOnce = True
      Form_Resize
      OnCaretPosChange
      UpdateToolbar
      Focus2Canvas
      Me.TimerInterval = 2000  'Every 2 seconds
    Else
      Me.TimerInterval = 0
      DoCmd.Close acForm, Me.Name
      Exit Sub
    End If
  End If
  
  If Not mfProgressRunning Then
    ShowProgressBar False
    Me.TimerInterval = 0
  End If
End Sub

Private Sub lblCurBkCol_Click()
  If Not LockUI() Then Exit Sub
  miCurBkColDisp = miCurBkColDisp + 1
  If miCurBkColDisp > 2 Then  '3 modes (rgb, hex, dec)
    miCurBkColDisp = 0
  End If
  OnCaretPosChange
  Focus2Canvas
  UnlockUI
End Sub

Private Sub lblCurFgCol_Click()
  If Not LockUI() Then Exit Sub
  miCurFgColDisp = miCurFgColDisp + 1
  If miCurFgColDisp > 2 Then  '3 modes
    miCurFgColDisp = 0
  End If
  OnCaretPosChange
  Focus2Canvas
  UnlockUI
End Sub

Private Sub lblStbSelection_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Not LockUI() Then Exit Sub
  On Error Resume Next
  
  If mfTransparent Then
    MsgBox "Sorry, can't show selection while in transparency mode", vbCritical
    UnlockUI
    Exit Sub
  End If
  
  If moCanvas.SelStartRow > 0 Then
    Dim iLeft         As Long
    Dim iTop          As Long
    Dim iRight        As Long
    Dim iBottom       As Long
    Dim iCharWidth    As Integer
    Dim iCharHeight   As Integer
    Dim fVisible      As Boolean
    Dim rcWindow      As RECT
    
    DoHourglass True
    
    fVisible = True
    'GetWindowRect moConCanvas.hwnd, rcWindow
    
    iCharHeight = moConCanvas.CharHeight
    iCharWidth = moConCanvas.CharWidth
    rcWindow.left = (moCanvas.SelStartCol - 1) * iCharWidth
    rcWindow.Right = moCanvas.SelEndCol * iCharWidth
    If moCanvas.SelStartRow >= moConCanvas.TopLine Then
      rcWindow.Top = (moCanvas.SelStartRow - moConCanvas.TopLine) * iCharHeight
      rcWindow.Bottom = (moCanvas.SelEndRow - moConCanvas.TopLine + 1) * iCharHeight
    Else
      If moCanvas.SelEndRow >= moConCanvas.TopLine Then
        rcWindow.Top = 0
        rcWindow.Bottom = (moCanvas.SelEndRow - moConCanvas.TopLine + 1) * iCharHeight
      Else
        fVisible = False
      End If
    End If
    'Debug.Print "Before MapPoints, left=" & rcWindow.Left & ", top=" & rcWindow.Top & ", right=" & rcWindow.Right & ", bottom=" & rcWindow.Bottom
    MapWindowPoints moConCanvas.hWnd, Me.hWnd, rcWindow
    'Debug.Print "After MapPoints, left=" & rcWindow.Left & ", top=" & rcWindow.Top & ", right=" & rcWindow.Right & ", bottom=" & rcWindow.Bottom
    
    If fVisible Then
      SetSelectionWindowCoords rcWindow.left, rcWindow.Top, rcWindow.Right - rcWindow.left, rcWindow.Bottom - rcWindow.Top
    End If
    
    ShowSelectionWindow fVisible
    
    DoHourglass False
  End If
  
  UnlockUI
End Sub

Private Sub lblStbSelection_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
  If mfTransparent Then
    Exit Sub
  End If
  
  DoHourglass True
  ShowSelectionWindow False
  DoEvents
  moConCanvas.SetAlphaTransparency 0
  DoHourglass False
End Sub

Private Sub lblSelEnd_Click()
  If Not LockUI() Then Exit Sub
  If moCanvas.SelEndRow > 0 Then
    moConCanvas.SetCaretPos moCanvas.SelEndRow, moCanvas.SelEndCol
    If Not moConCanvas.IsRowVisible(moCanvas.SelEndRow) Then
      moConCanvas.ScrollTo moCanvas.SelEndRow
      moConCanvas.RefreshWindow
    End If
  End If
  UnlockUI
End Sub

Private Sub lblSelStart_Click()
  If Not LockUI() Then Exit Sub
  If moCanvas.SelStartRow > 0 Then
    moConCanvas.SetCaretPos moCanvas.SelStartRow, moCanvas.SelStartCol
    If Not moConCanvas.IsRowVisible(moCanvas.SelStartRow) Then
      moConCanvas.ScrollTo moCanvas.SelStartRow
      moConCanvas.RefreshWindow
    End If
  End If
  UnlockUI
End Sub

Private Sub lblStbCharCode_Click()
  If (GetKeyState(VK_CONTROL) < 0) Then
    Dim sCharCode   As String
    Dim iCharCode   As Integer
    Dim iRow        As Integer
    Dim iCol        As Integer
    
    moConCanvas.GetCaretPos iRow, iCol
    If (iRow <= 0) Or (iCol <= 0) Then
      Exit Sub
    End If
    
    sCharCode = mgrdCanvas.CharAt(iRow, iCol)
    If Len(sCharCode) = 0 Then Exit Sub
    iCharCode = AscW(sCharCode)
    MessageManager.Broadcast Me.Name, MSGTOPIC_FINDCHAR, iCharCode, GetCharMapFormName()
  Else
    mbDispCharCodeToggle = mbDispCharCodeToggle + 1
    If mbDispCharCodeToggle > 1 Then mbDispCharCodeToggle = 0
    UpdateStbCharCode
  End If
End Sub

Private Sub UpdateDialogTitle(ByVal pfDirty As Boolean)
  On Error Resume Next
  Dim sCaption    As String
  Dim lPixelWidth As Long
  Dim rcClient    As RECT
  
  sCaption = "Canvas ("
  If Len(moCanvas.Filename) > 0 Then
    'max 2/3 of dialog client width
    Call moConCanvas.GetClientRect(rcClient)
    lPixelWidth = rcClient.Right - rcClient.left
    lPixelWidth = (lPixelWidth * 2) / 3
    'Debug.Print "lPixelWidth="; lPixelWidth
    sCaption = sCaption & CompactPath(Me.hWnd, moCanvas.Filename, lPixelWidth) & ")"
  Else
    sCaption = sCaption & "new)"
  End If
  If pfDirty Then sCaption = sCaption & "*"
  Me.Caption = sCaption
End Sub

Private Sub mgrdCanvas_OnDirtyChange(ByVal pfDirty As Boolean)
  UpdateDialogTitle pfDirty
End Sub

Private Sub OptClearDir_AfterUpdate()
  Focus2Canvas
End Sub

Private Sub optDelWhat_AfterUpdate()
  Focus2Canvas
End Sub

Private Sub optLineCol_AfterUpdate()
  Focus2Canvas
End Sub

Private Sub optTransparencyMode_AfterUpdate()
  If Not LockUI() Then Exit Sub
  On Error Resume Next
  moCanvas.TransparencyMode = optTransparencyMode
  If optTransparencyMode = eTransparencyMode.eColorTransparency Then
    SetCanvasTransparentColor moCanvas.TransparentColor
  Else
    SetCanvasAlphaTransparencyPct moCanvas.TransparentAlphaPct
  End If
  SetCanvasTransparency mfTransparent, moCanvas.TransparencyMode
  SetCanvasTransparency CBool(Me.chkTransparency), moCanvas.TransparencyMode
  UnlockUI
End Sub

Private Sub AddAndSelectColor(ByVal plColor As Long)
  Dim rowParams     As New CRow
  rowParams.AddCol "color", plColor, 0, 0
  MessageManager.Broadcast Me.Name, MSGTOPIC_ADDNSELCOLOR, rowParams, GetPaletteFormName()
End Sub

Private Sub rectCurBkCol_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Not LockUI() Then Exit Sub
  If (Button = acLeftButton) And (Shift And acCtrlMask) Then
    Dim iRow    As Integer
    Dim iCol    As Integer
    Dim lColor  As Long
    
    On Error Resume Next
    moConCanvas.GetCaretPos iRow, iCol
    lColor = mgrdCanvas.CharBackCol(iRow, iCol)
    AddAndSelectColor lColor
  End If
  UnlockUI
End Sub

Private Sub rectCurFgCol_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Not LockUI() Then Exit Sub
  If (Button = acLeftButton) And (Shift And acCtrlMask) Then
    Dim iRow    As Integer
    Dim iCol    As Integer
    Dim lColor  As Long
    
    On Error Resume Next
    moConCanvas.GetCaretPos iRow, iCol
    lColor = mgrdCanvas.CharForeCol(iRow, iCol)
    AddAndSelectColor lColor
  End If
  UnlockUI
End Sub

Private Sub rectTransColor_Click()
  If Not LockUI() Then Exit Sub
  Dim lColor    As Long
  If PickColor(Me.hWnd, lColor, moCanvas.ForeColor) Then
    On Error Resume Next
    SetCanvasTransparentColor lColor
    SetCanvasTransparency mfTransparent, moCanvas.TransparencyMode  'just redo
  End If
  UnlockUI
End Sub

Private Sub txtBkgndBitmap_Enter()
  DisableFormKeyPreview
End Sub

Private Sub txtBkgndBitmap_Exit(Cancel As Integer)
  EnableFormKeyPreview
End Sub

Private Sub txtFontSize_AfterUpdate()
  If Not LockUI() Then Exit Sub
  On Error Resume Next
  If Val(Me.txtFontSize) > 1 Then
    moCanvas.FontSize = Val(Me.txtFontSize)
    DoHourglass True
    CreateView
    UpdateView
    DoHourglass False
    Focus2Canvas
  End If
  UnlockUI
End Sub

Private Sub cmdDecSize_Click()
  If Not LockUI() Then Exit Sub
  On Error Resume Next
  Form_Resize
  If Val(Me.txtFontSize) > 1 Then
    moCanvas.FontSize = Val(Me.txtFontSize) - 1
    Me.txtFontSize = moCanvas.FontSize
    DoHourglass True
    CreateView
    UpdateView
    Focus2Canvas
    DoHourglass False
  End If
  UnlockUI
End Sub

Private Sub cmdIncSize_Click()
  If Not LockUI() Then Exit Sub
  On Error Resume Next
  moCanvas.FontSize = Val(Me.txtFontSize) + 1
  Me.txtFontSize = moCanvas.FontSize
  DoHourglass True
  CreateView
  UpdateView
  Focus2Canvas
  DoHourglass False
  UnlockUI
End Sub

Private Sub SetBackColor(ByVal plBackColor As Long)
  moCanvas.BackColor = plBackColor
  Me.cmdBColor.BackColor = plBackColor
End Sub

Private Sub SetForeColor(ByVal plForeColor As Long)
  moCanvas.ForeColor = plForeColor
  Me.cmdFColor.BackColor = plForeColor
End Sub

Private Sub SetPenBkColor(ByVal pvColor As Variant)
  moCanvas.PenBackColor = pvColor
  If Not IsNull(pvColor) Then
    Me.cmdTextBackcolor.Caption = ""
    Me.cmdTextBackcolor.BackColor = pvColor
  Else
    Me.cmdTextBackcolor.Caption = "x"
    Me.cmdTextBackcolor.BackColor = Me.cmdClearBkgndImage.BackColor
  End If
End Sub

Private Sub SetPenForeColor(ByVal pvColor As Variant)
  moCanvas.PenForeColor = pvColor
  If Not IsNull(pvColor) Then
    Me.cmdTextForecolor.Caption = ""
    Me.cmdTextForecolor.BackColor = pvColor
  Else
    Me.cmdTextForecolor.Caption = "x"
    Me.cmdTextForecolor.BackColor = Me.cmdClearBkgndImage.BackColor
  End If
End Sub

Private Sub cmdClearPenColor_Click()
  SetPenForeColor Null
End Sub

Private Sub cmClearBkColor_Click()
  SetPenBkColor Null
End Sub

Private Function InitNewCanvas(poCanvas As CCanvas) As Boolean
  Dim lBackColor  As Long
  Dim lForeColor  As Long
  Dim iRows       As Integer
  Dim iCols       As Integer
  
  AppIniFile.Section = INISECTION_CANVAS
  
  poCanvas.FontName = Trim$(AppIniFile.GetString(INIKEY_DEFFONTNAME))
  If Len(moCanvas.FontName) = 0 Then
    poCanvas.FontName = DEFAULT_FONTNAME
  End If
  poCanvas.FontSize = AppIniFile.GetInt(INIKEY_DEFFONTSIZE, DEFAULT_FONTSIZE)

  lForeColor = AppIniFile.GetLong(INIKEY_DEFFORECOLOR, QBColor(QBCOLOR_WHITE))
  lBackColor = AppIniFile.GetLong(INIKEY_DEFBACKCOLOR, vbBlack)
  
  'Initialize the grid object that will hold, size and render the buffer
  iRows = AppIniFile.GetInt(INIKEY_DEFCANVASROWS, DEFAULT_CANVASROWS)
  iCols = AppIniFile.GetInt(INIKEY_DEFCANVASCOLS, DEFAULT_CANVASCOLS)
  If Not poCanvas.ConsoleGrid.Resize(iRows, iCols) Then  'default values
    MsgBox "Failed to resize buffers to [ " & iRows & "  x " & iCols & "]:" & vbCrLf & vbCrLf & poCanvas.ConsoleGrid.LastErrDesc, vbCritical
    InitNewCanvas = False
    Exit Function
  End If
  
  'Set up the default drawing environment
  SetForeColor lForeColor
  SetBackColor lBackColor
  SetPenForeColor Null
  SetPenBkColor vbWhite
  
  poCanvas.TransparencyMode = AppIniFile.GetInt(INIKEY_CANVASTRANSMODE, eTransparencyMode.eColorTransparency)
  poCanvas.TransparentColor = AppIniFile.GetLong(INIKEY_CANVASTRANSCOLOR, vbWhite)
  poCanvas.TransparentAlphaPct = AppIniFile.GetInt(INIKEY_CANVASALPHAPCT, 0)
  
  InitNewCanvas = True
End Function

Private Sub LoadFontFamilies()
  Dim sSource     As String
  Dim iIniValue   As Integer
  Dim fOK         As Boolean
  
  AppIniFile.Section = INISECTION_CANVAS
  iIniValue = AppIniFile.GetInt(INIKEY_FONTFAMILY, FF_MODERN)
  
  'sSource = "0;FF_DONTCARE;16;FF_ROMAN;32;FF_SWISS;48;FF_MODERN;64;FF_SCRIPT;80;FF_DECORATIVE"
  sSource = "0;(All);16;Roman;32;Swiss;48;Modern;64;Script;80;Decorative"
  Me.cboFontFamily.RowSource = sSource
  Me.cboFontFamily = iIniValue
  miFontFamily = iIniValue
  fOK = MFonts.LoadFontNames(miFontFamily)
  If Not fOK Then
    ShowUFError "Failed to load font families", FontLastErrDesc()
  End If
End Sub

Private Sub ReloadFonts()
  Dim i               As Integer
  Dim sSource         As String
  Dim sDefaultFont    As String
  
  Me.cboFontName.RowSource = GetFontsComboSource(miFontFamily)
  FontSetSelectedFont moCanvas.FontName
  Me.cboFontName = moCanvas.FontName
  Me.txtTypeText.FontName = moCanvas.FontName
  Me.txtFontSize = moCanvas.FontSize
End Sub

Private Function InitForm() As Boolean
  Dim lWidth      As Long
  Dim lHeight     As Long
  Dim fResize     As Boolean
  Dim rcWindow    As RECT
  
  'V02.00.00 Apply transparency settings
  On Error Resume Next
  
  CreateRibbon
  
  AppIniFile.Section = INISECTION_CANVAS
  
  Call GetWindowRect(Me.hWnd, rcWindow)
  lWidth = AppIniFile.GetInt(INIKEY_CANVASWINDOWWIDTH, 0)
  lHeight = AppIniFile.GetInt(INIKEY_CANVASWINDOWHEIGHT, 0)
  If (lWidth <> 0) Or (lHeight <> 0) Then
    If lHeight = 0 Then
      lHeight = rcWindow.Bottom - rcWindow.Top
    Else
      If lWidth = 0 Then
        lWidth = rcWindow.Right - rcWindow.left
      End If
    End If
    fResize = True
  End If
  
  LoadFontFamilies
  ReloadFonts
    
  'Initialize controls on the form
  
  'V02.00.00 Get back transparency settings
'    mfTransparent = AppIniFile.GetFlag(INIKEY_CANVASTRANSPARENT, False)
  msTransImageFilename = Trim$(AppIniFile.GetString(INIKEY_CANVASBKGNDIMAGE))
  mfShowGrid = CBool(AppIniFile.GetInt(INIKEY_SHOWGRID, 0))
  mfCursorFollowMouse = CBool(AppIniFile.GetInt(INIKEY_CURSORFOLLOWMOUSE, 1))
  
  'V02.00.00 Apply transparency settings
  SetCanvasTransparentColor moCanvas.TransparentColor
  SetCanvasBackgroundImageFromFile msTransImageFilename
  Me.optTransparencyMode = moCanvas.TransparencyMode
  Me.txtTransAlphaPct = moCanvas.TransparentAlphaPct & ""
'    Me.chkTransparency = mfTransparent
  SetTransColor moCanvas.TransparentColor
  
  ShowCharMapDialog
  ShowPaletteDialog
  
  CreateView
  CreateProgressBar
  
  PositionRibbonControls
  
  If fResize Then
    Call MoveWindow(Me.hWnd, 0, 0, lWidth, lHeight, 1&)
  End If
  
  UpdateDialogTitle False
  UpdateFreehandIndicator
  
  mlLastTypedCharCode = AscW("A")
  InitDebugConsole
  
  MessageManager.SubscribeMulti Me, Array(MSGTOPIC_LOCKUI, MSGTOPIC_UNLOCKUI, MSGTOPIC_CANUNLOAD, MSGTOPIC_CHARSELECTED)
  
  InitForm = True
End Function

Private Sub ReapplyTransparencySettings()
  On Error Resume Next
  
  Dim fSaveTrans As Boolean
  If mfTransparent Then
    fSaveTrans = mfTransparent
    SetCanvasTransparency False, moCanvas.TransparencyMode
    DoEvents
    moConCanvas.RefreshWindow
    mfTransparent = fSaveTrans
  End If
  SetCanvasTransparency mfTransparent, moCanvas.TransparencyMode
  Me.chkTransparency = mfTransparent
  moConCanvas.RefreshWindow
  DoEvents
End Sub

Public Sub RepositionConsouls()
  On Error Resume Next
  
  'Adjust to full form client area of the detail section of the form.
  Dim iHalfWidth  As Integer
  Dim iWidth      As Integer
  Dim iHeight     As Integer
  Dim iStHeight   As Integer
  Dim iHdHeight   As Integer
  Dim iDbgHeight  As Integer
  
  GetClientRect Me.hWnd, mrcClient
  
  iHdHeight = TwipsToPixelsY(Me.Section(AcSection.acHeader).Height)  'Height of the header in pixels
  iHdHeight = iHdHeight + TwipsToPixelsY(Me.rectToolbar.Top + Me.rectToolbar.Height)
  iStHeight = TwipsToPixelsY(Me.Section(AcSection.acFooter).Height)  'Height of the footer in pixels
  
  iWidth = mrcClient.Right - mrcClient.left
  iHeight = mrcClient.Bottom - mrcClient.Top - iHdHeight - iStHeight
  
  If mfConDebugVisible Then
    iDbgHeight = mconDebug.LineHeight * miDbgWindowLineCt
    mconDebug.MoveWindow 0, iHdHeight + iHeight - iDbgHeight, iWidth, iDbgHeight
    iHeight = iHeight - iDbgHeight
  End If
  
  'canvas
  moConCanvas.MoveWindow 0, iHdHeight, iWidth, iHeight
  
  'background image
  Me.imgBackground.Move 0, PixelsToTwipsY(iHdHeight), PixelsToTwipsX(iWidth), PixelsToTwipsY(iHeight)
  
  'progressbar
  RepositionProgressBar
  
  'ribbon
  moRibbon.TabConsole.MoveWindow 0, 0, mrcClient.Right - mrcClient.left, TwipsToPixelsY(Me.rectRibbon.Height)
  Me.rectToolbar.Width = Me.Width
  Me.linToolbar.left = 0
  Me.linToolbar.Top = Me.rectToolbar.Top + Me.rectToolbar.Height - PixelsToTwipsY(1)
  Me.linToolbar.Width = Me.rectToolbar.Width
  
  moRibbon.OnResize
End Sub

Private Sub RepositionProgressBar()
  'The progressbar goes on top right
  Dim iPGBWidth As Integer
  On Error Resume Next
  If (moConProgress.CharWidth() = 0) Or (moConProgress.CharHeight() = 0) Then
    moConProgress.OutputLn ""
  End If
  iPGBWidth = moConProgress.CharWidth() * (MAX_PRGB_CAPTION_LEN + moProgressBar.CharWidth) + 10
  'moConProgress.MoveWindow mrcClient.Right - mrcClient.left - iPGBWidth - 8, TwipsToPixelsY(Me.cmdLoadFile.Top), iPGBWidth, moConProgress.LineHeight + 3
  moConProgress.MoveWindow mrcClient.Right - mrcClient.left - iPGBWidth - 8, 4&, iPGBWidth, moConProgress.LineHeight + 3
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  Dim rowParams   As New CRow
  
  If Not moConCanvas Is Nothing Then
    GetClientRect Me.hWnd, mrcClient
    RepositionConsouls
    RepositionProgressBar
  End If
  
  DoEvents  'Absolutely necessary to get the window size after the event
  rowParams.AddCol "WindowLeft", Me.WindowLeft, 0, 0
  rowParams.AddCol "WindowTop", Me.WindowTop, 0, 0
  rowParams.AddCol "WindowWidth", Me.WindowWidth, 0, 0
  rowParams.AddCol "WindowHeight", Me.WindowHeight, 0, 0
  MessageManager.Broadcast Me.Name, MSGTOPIC_CANVASRESIZED, rowParams
End Sub

' Managing console windows (one on this form, the canvas)
Private Function CreateConsouls() As Boolean
  Dim hwndParent  As LongPtr
  
  On Error GoTo CreateConsouls_Err
  
  hwndParent = Me.hWnd
  
  'We repeatedly call this function to create/destroy the console windows
  If Not moConCanvas Is Nothing Then
    moConCanvas.SetWmPaintCallback WMPAINTCBK_AFTER, 0
    ConsoulEventDispatcher.UnregisterEventSink moConCanvas.hWnd
    Set moConCanvas = Nothing
  End If
  
  'The console window displaying the character set
  Set moConCanvas = New CConsoul
  moConCanvas.FontName = moCanvas.FontName
  moConCanvas.FontSize = moCanvas.FontSize
  moConCanvas.MaxCapacity = mgrdCanvas.Rows + 2 '1
  moConCanvas.ForeColor = moCanvas.ForeColor
  moConCanvas.BackColor = moCanvas.BackColor
  'Create the console window and tell the library that we want click feedback
  If Not moConCanvas.Attach(hwndParent, 0, 0, 0, 0, AddressOf MSupport.OnConsoulMouseButton, piCreateAttributes:=LW_RENDERMODEBYLINE Or LW_TRACK_ZONES) Then
    MsgBox "Failed to create canvas window", vbCritical
    GoTo CreateConsouls_Exit
  End If
  'These properties can be set only after attaching (creating) the console window
  moConCanvas.AutoAdjustWidth = Me.chkAutoAdjustWidth
  moConCanvas.LinePadding(elsTop) = Val(Nz(Me.txtLinePaddingTop, 0))
  moConCanvas.LinePadding(elsBottom) = Val(Nz(Me.txtLinePaddingBottom, 0))
  moConCanvas.LineSpacing(elsTop) = Val(Nz(Me.txtLineSpacingTop, 0))
  moConCanvas.LineSpacing(elsBottom) = Val(Nz(Me.txtLineSpacingBottom, 0))
  
  moConCanvas.SetWmPaintCallback WMPAINTCBK_AFTER, AddressOf MSupport.OnConsoulWmPaint
  
  'Let the system know that click for our canvas console should arrive here
  ConsoulEventDispatcher.RegisterEventSink moConCanvas.hWnd, Me, eCsMouseEvent
  ConsoulEventDispatcher.RegisterEventSink moConCanvas.hWnd, Me, eCsWmPaint
  'Show the console window
  moConCanvas.ShowWindow True
  
  On Error Resume Next
  'V02.00.00 apply transparency setting
  DoEvents
  SetCanvasTransparency mfTransparent, moCanvas.TransparencyMode
  
  CreateConsouls = True
  
CreateConsouls_Exit:
  Exit Function

CreateConsouls_Err:
  MsgBox "Failed to create consoul's output windows"
End Function

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
    moRibbon.TabControl.AddTab RIBBONTAB_DISPLAY, "Display"
    moRibbon.TabControl.AddTab RIBBONTAB_TEXT, "Text"
    moRibbon.TabControl.AddTab RIBBONTAB_CLIPBOARD, "Clipboard"
    moRibbon.TabControl.AddTab RIBBONTAB_LINECOL, "Line & column"
    moRibbon.TabControl.AddTab RIBBONTAB_DEVELOPER, "Developer"
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

Private Function CreateProgressBar() As Boolean
  Dim hwndParent  As LongPtr
  
  On Error GoTo CreateProgressBar_Err
  
  hwndParent = Me.hWnd
  
  'V02.00.01 Progress control
  Set moConProgress = New CConsoul
  moConProgress.FontName = "Consolas"
  moConProgress.FontSize = 8
  moConProgress.MaxCapacity = 1
  moConProgress.ForeColor = vbBlack
  moConProgress.BackColor = moRibbon.TabConsole.BackColor  'Me.rectToolbar.BackColor 'Me.Section(AcSection.acHeader).BackColor
  moConProgress.LineSpacing(elsTop) = 4
  moConProgress.LineSpacing(elsBottom) = 4
  'Create the console window and tell the library that we want click feedback
  If Not moConProgress.Attach(hwndParent, 0, 0, 0, 0, piCreateAttributes:=LW_RENDERMODEBYLINE) Then
    MsgBox "Failed to create progress window", vbCritical
    GoTo CreateProgressBar_Exit
  End If
  moConProgress.ShowWindow True
  moConProgress.OutputLn ""
  moConProgress.Clear
  moConProgress.ShowWindow False
  
  CreateProgressBar = True
  
CreateProgressBar_Exit:
  Exit Function

CreateProgressBar_Err:
  MsgBox "Failed to create consoul's output windows"
End Function

Private Function AskEverybodyIfWeCanUnload(ByVal psMsgTitle As String) As Boolean
  Dim lRet As Long
  lRet = MessageManager.Broadcast(Me.Name, MSGTOPIC_CANUNLOAD, psMsgTitle)
  AskEverybodyIfWeCanUnload = CBool(lRet = 0&)
End Function

Private Sub Form_Unload(Cancel As Integer)
  Dim rcWindow    As RECT
  
  On Error Resume Next
  
  'Ask to save changes if dirty
  If Not AskEverybodyIfWeCanUnload("Close") Then
    Cancel = 1
    Exit Sub
  End If
  'Send order to close and unload to everybody (except us, we don't listen)
  MessageManager.Broadcast Me.Name, MSGTOPIC_UNLOADNOW, Nothing
  
  DestroyDebugConsole
  
  Call GetWindowRect(Me.hWnd, rcWindow)
  AppIniFile.Section = INISECTION_CANVAS
  AppIniFile.SetString INIKEY_CANVASWINDOWWIDTH, (rcWindow.Right - rcWindow.left) & ""
  AppIniFile.SetString INIKEY_CANVASWINDOWHEIGHT, (rcWindow.Bottom - rcWindow.Top) & ""
  AppIniFile.SetString INIKEY_DEFFONTNAME, Me.cboFontName & ""
  AppIniFile.SetString INIKEY_DEFFONTSIZE, Me.txtFontSize & ""
  AppIniFile.SetString INIKEY_DEFFORECOLOR, moCanvas.ForeColor & ""
  AppIniFile.SetString INIKEY_DEFBACKCOLOR, moCanvas.BackColor & ""
  AppIniFile.SetString INIKEY_CANVASTRANSPARENT, Abs(CInt(mfTransparent)) & ""
  AppIniFile.SetString INIKEY_CANVASTRANSMODE, CInt(moCanvas.TransparencyMode) & ""
  AppIniFile.SetString INIKEY_CANVASTRANSCOLOR, moCanvas.TransparentColor & ""
  AppIniFile.SetString INIKEY_CANVASALPHAPCT, moCanvas.TransparentAlphaPct & ""
  AppIniFile.SetString INIKEY_CANVASBKGNDIMAGE, msTransImageFilename
  AppIniFile.SetString INIKEY_SHOWGRID, Abs(CInt(mfShowGrid)) & ""
  AppIniFile.SetString INIKEY_CURSORFOLLOWMOUSE, Abs(CInt(mfCursorFollowMouse)) & ""
  
  'kill the clipboard
  Set mgrdClipboard = Nothing
  
  'Tell the system to not forward anymore mouse message to us
  If Not moConCanvas Is Nothing Then
    moConCanvas.SetWmPaintCallback WMPAINTCBK_AFTER, 0
    ConsoulEventDispatcher.UnregisterEventSink moConCanvas.hWnd
  End If
  
  MessageManager.Unsubscribe Me.Name, ""
  
  Set moConProgress = Nothing
  Set moConCanvas = Nothing
  Set moConSel = Nothing
  Set mgrdCanvas = Nothing
  Set moRibbon = Nothing
  Set moCanvas = Nothing
End Sub

Public Sub CreateView()
  If Len(moCanvas.FontName) = 0 Then
    moCanvas.FontName = DEFAULT_FONTNAME
  End If
  If moCanvas.FontSize <= 0 Then
    moCanvas.FontSize = DEFAULT_FONTSIZE
  End If
  CreateConsouls
  'Load the grid in the console
  On Error Resume Next
  DoHourglass True
  RepositionProgressBar
  mgrdCanvas.LoadConsole moConCanvas, Me
  moConCanvas.CaretBlinkMs = 500
  ShowCaret True
  moConCanvas.SetCaretPos 1, 1
  moConCanvas.ScrollTop
  DoHourglass False
End Sub

Public Sub UpdateView()
  On Error Resume Next
  RepositionConsouls
  RepositionProgressBar
  moConCanvas.RefreshWindow
  OnCaretPosChange
  UpdateDebugConsole True
End Sub

Public Sub RefreshView(Optional ByVal pfReload As Boolean = False)
  'refresh visible lines
  Dim iRow      As Integer
  For iRow = moConCanvas.TopLine To (moConCanvas.TopLine + moConCanvas.MaxVisibleRows - 1)
    If Not pfReload Then
      moConCanvas.RedrawLine iRow
    Else
      RenderLine iRow
    End If
  Next iRow
  OnCaretPosChange
  UpdateDebugConsole True
End Sub

Private Function ICsMouseEventSink_OnMouseButton( _
  ByVal phWnd As LongPtr, _
  ByVal piEvtCode As Integer, _
  ByVal pwParam As Integer, _
  ByVal piZoneID As Integer, _
  ByVal piRow As Integer, _
  ByVal piCol As Integer, _
  ByVal piPosX As Integer, _
  ByVal piPosY As Integer) As Integer
  
  Dim lCharCode As Long
  Dim fUpdatePosDisplay As Boolean
  Dim sChar As String
  
  Static iLastRow As Integer, iLastCol As Integer
  
  If phWnd = moConCanvas.hWnd Then
    Focus2Canvas
    fUpdatePosDisplay = CBool(mfCursorFollowMouse And (pwParam And MK_SHIFT))
    
    If (piEvtCode = eWmMouseButton.WM_LBUTTONUP) Or (mfCursorFollowMouse And (pwParam And MK_SHIFT)) Then
      If (piRow > 0) And (piCol > 0) And (piRow <= mgrdCanvas.Rows) And (piCol <= mgrdCanvas.Cols) Then
        fUpdatePosDisplay = True
      End If
    End If
    
    If mfCursorFollowMouse And (pwParam And MK_LBUTTON) Then
      If (piRow > 0) And (piCol > 0) And (piRow <= mgrdCanvas.Rows) And (piCol <= mgrdCanvas.Cols) Then
        If (piRow <> iLastRow) Or (piCol <> iLastCol) Then
          TypeCharAt mlLastTypedCharCode, piRow, piCol
          iLastRow = piRow
          iLastCol = piCol
        End If
      End If
    End If
    
    If fUpdatePosDisplay Then
      'piCol = piCol - mgrdCanvas.MarginWidth
      moConCanvas.SetCaretPos piRow, piCol
      OnCaretPosChange
    End If
  ElseIf phWnd = moRibbon.TabConsole.hWnd Then
    ICsMouseEventSink_OnMouseButton = moRibbon.OnTabsMouseButton(phWnd, piEvtCode, pwParam, piZoneID, piRow, piCol, piPosX, piPosY)
  End If
End Function

Private Sub txtFontSize_Enter()
  DisableFormKeyPreview
End Sub

Private Sub txtFontSize_Exit(Cancel As Integer)
  EnableFormKeyPreview
End Sub

Private Sub txtFontSize_KeyDown(KeyCode As Integer, Shift As Integer)
  If Not IsIntegerEditKeyCodeAllowed(KeyCode) Then
    KeyCode = 0
  End If
End Sub

Private Sub ShowCaret(ByVal pfShow As Boolean)
  If pfShow Then
    If Not moConCanvas.IsCaretVisible Then
      moConCanvas.ShowCaret True
    End If
  Else
    If moConCanvas.IsCaretVisible Then
      moConCanvas.ShowCaret False
    End If
  End If
End Sub

'Canvas Margin MUST be adjusted before call
Private Sub ForeColorAction(ByVal piRow As Integer, ByVal piCol As Integer)
  If IsNull(moCanvas.PenForeColor) Then
    mgrdCanvas.HasForeCol(piRow, piCol) = False
  Else
    mgrdCanvas.CharForeCol(piRow, piCol) = moCanvas.PenForeColor
  End If
  RenderLine piRow
  'position the caret where it was clicked
  moConCanvas.SetCaretPos piRow, piCol
  OnCaretPosChange
  UpdateToolbar
End Sub

'Canvas Margin MUST be adjusted before call
Private Sub BackColorAction(ByVal piRow As Integer, ByVal piCol As Integer)
  If IsNull(moCanvas.PenBackColor) Then
    mgrdCanvas.HasBackCol(piRow, piCol) = False
  Else
    mgrdCanvas.CharBackCol(piRow, piCol) = moCanvas.PenBackColor
  End If
  RenderLine piRow
  'position the caret where it was clicked
  moConCanvas.SetCaretPos piRow, piCol
  OnCaretPosChange
  UpdateToolbar
End Sub

Private Sub OnCaretPosChange()
  Dim iRow      As Integer
  Dim iCol      As Integer
  Dim lColor    As Long
  
  moConCanvas.GetCaretPos iRow, iCol
  If (iRow > 0) And (iCol > 0) Then
    Me.lblCursorPos.Caption = iRow & "," & iCol
  End If
  If (iRow <= 0) Or (iCol <= 0) Then
    Exit Sub
  End If
  
  'Get font attributes at this position and set checkboxes accordingly
  SetChecksFromAttribs
  
  'margin already accounted for from here
  If mgrdCanvas.HasBackCol(iRow, iCol) Then
    lColor = mgrdCanvas.CharBackCol(iRow, iCol)
    Me.rectCurBkCol.BackStyle = 1
    Me.rectCurBkCol.BackColor = lColor
    Me.lblCurBkCol.Caption = GetColorDispString(lColor, miCurBkColDisp)
  Else
    Me.rectCurBkCol.BackStyle = 0
    Me.rectCurBkCol.BackColor = Me.Section(AcSection.acFooter).BackColor
    Me.lblCurBkCol.Caption = "(not set)"
  End If
  
  If mgrdCanvas.HasForeCol(iRow, iCol) Then
    lColor = mgrdCanvas.CharForeCol(iRow, iCol)
    Me.rectCurFgCol.BackStyle = 1
    Me.rectCurFgCol.BackColor = lColor
    Me.lblCurFgCol.Caption = GetColorDispString(lColor, miCurFgColDisp)
  Else
    Me.rectCurFgCol.BackStyle = 0
    Me.rectCurFgCol.BackColor = Me.Section(AcSection.acFooter).BackColor
    Me.lblCurFgCol.Caption = "(not set)"
  End If
  
  UpdateStbCharCode
  UpdateToolbar
  UpdateDebugConsole False
End Sub

Private Sub ShowCharMapDialog()
  On Error Resume Next
  Dim oDialog As CCharMapDialog
  'Create and show the modeless char map dialog
  Set oDialog = New CCharMapDialog
  oDialog.IIDialog.ShowDialog False
  Set oDialog = Nothing
End Sub

Private Sub ShowPaletteDialog()
  On Error Resume Next
  Dim oDialog As CPaletteDialog
  'Create and show the modeless char map dialog
  Set oDialog = New CPaletteDialog
  oDialog.IIDialog.ShowDialog False
  Set oDialog = Nothing
End Sub

Private Sub SetAttribsFromChecks()
  Dim i As Integer
  Dim iAttribs    As Integer
  Dim iRow        As Integer
  Dim iCol        As Integer
  
  For i = 0 To 12
    'Debug.Print "Attrib #"; i; "="; Abs(Me.Controls("chkAttr" & 2 ^ i).Value); ", "; ((2 ^ i) * Abs(Me.Controls("chkAttr" & 2 ^ i).Value))
    iAttribs = iAttribs Or ((2 ^ i) * Abs(CBool(Me.Controls("chkAttr" & 2 ^ i).Value)))
  Next i
  'Debug.Print "Attribs="; Hex$(iAttribs); " ("; iAttribs; ")"
  moConCanvas.GetCaretPos iRow, iCol
  mgrdCanvas.CharAttribs(iRow, iCol) = iAttribs
  RenderLine iRow
  Focus2Canvas
End Sub

'Called by OnCaretPosChange
Private Sub SetChecksFromAttribs()
  Dim i As Integer
  Dim iAttribs    As Integer
  Dim iRow        As Integer
  Dim iCol        As Integer
  
  moConCanvas.GetCaretPos iRow, iCol
  iAttribs = mgrdCanvas.CharAttribs(iRow, iCol)
  
  For i = 0 To 12
    Me.Controls("chkAttr" & 2 ^ i) = Abs(CBool(iAttribs And (2 ^ i)))
  Next i
  'Debug.Print "Attribs="; Hex$(iAttribs); " ("; iAttribs; ")"
End Sub

Private Sub SetButtonEnabled(ByRef poButton As Control, ByVal pfEnabled As Boolean)
  If poButton.Enabled <> pfEnabled Then
    poButton.Enabled = pfEnabled
  End If
End Sub

Private Sub UpdateToolbar()
  Dim fEnabled      As Boolean
  
  fEnabled = True 'CBool(meTool = dtText)
  
  SetButtonEnabled Me.cmdNewFile, True
  SetButtonEnabled Me.cmdLoadFile, True
  SetButtonEnabled Me.cmdSaveFileAs, True
  
  SetButtonEnabled Me.cmdInsertLineCol, fEnabled
  SetButtonEnabled Me.cmdDeleteLineCol, fEnabled
  
  SetButtonEnabled Me.cmdClear, fEnabled
  
  ShowSelection
  UpdateCanvasSizeIndics
  
  UpdateCanvasColorsIndics
End Sub

Private Sub UpdateCanvasColorsIndics()
  Me.cmdCurBackColor.BackColor = moCanvas.BackColor
  Me.cmdCurBackColor.HoverColor = moCanvas.BackColor
  Me.cmdCurBackColor.PressedForeColor = moCanvas.BackColor
  
  Me.cmdCurForeColor.BackColor = moCanvas.ForeColor
  Me.cmdCurForeColor.HoverColor = moCanvas.ForeColor
  Me.cmdCurForeColor.PressedForeColor = moCanvas.ForeColor
End Sub

Private Sub UpdateStbCharCode()
  Dim iCharCode   As Integer
  Dim sCharCode   As String
  Dim sHexCode    As String
  Dim iRow        As Integer
  Dim iCol        As Integer
  
  moConCanvas.GetCaretPos iRow, iCol
  If (iRow <= 0) Or (iCol <= 0) Then
    Me.lblStbCharCode.Caption = "#n/a"
    Exit Sub
  End If
  
  sCharCode = mgrdCanvas.CharAt(iRow, iCol)
  If Len(sCharCode) = 0 Then Exit Sub
  iCharCode = AscW(sCharCode)
  sHexCode = Hex$(iCharCode)
  If Len(sHexCode) < 4 Then
    sHexCode = String$(4 - Len(sHexCode), "0") & sHexCode
  End If
  
  If mbDispCharCodeToggle = 0 Then
    Me.lblStbCharCode.Caption = "Hex: " & sHexCode
  Else
    Me.lblStbCharCode.Caption = "Dec: " & iCharCode
  End If
End Sub

Private Sub chkAttr1_Click()
  SetAttribsFromChecks
End Sub

Private Sub chkAttr2_Click()
  SetAttribsFromChecks
End Sub

Private Sub chkAttr4_Click()
  SetAttribsFromChecks
End Sub

Private Sub chkAttr8_Click()
  SetAttribsFromChecks
End Sub

Private Sub chkAttr16_Click()
  SetAttribsFromChecks
End Sub

Private Sub chkAttr32_Click()
  SetAttribsFromChecks
End Sub

Private Sub chkAttr64_Click()
  SetAttribsFromChecks
End Sub

Private Sub chkAttr128_Click()
  SetAttribsFromChecks
End Sub

Private Sub chkAttr256_Click()
  SetAttribsFromChecks
End Sub

Private Sub chkAttr512_Click()
  SetAttribsFromChecks
End Sub

Private Sub chkAttr1024_Click()
  SetAttribsFromChecks
End Sub

Private Sub chkAttr2048_Click()
  SetAttribsFromChecks
End Sub

Private Sub chkAttr4096_Click()
  SetAttribsFromChecks
End Sub

Private Function CanUnload(ByVal psMsgTitle As String) As Boolean
  Dim sMsg      As String
  Dim iRet      As VbMsgBoxResult
  
  If mgrdCanvas.Dirty Then
    sMsg = "The canvas has been modified, do you want to save changes ?"
    
    iRet = MsgBox(sMsg, vbYesNo + vbQuestion + vbDefaultButton1, psMsgTitle)
    If iRet = vbYes Then
      If SaveToFile() Then
        CanUnload = True
      End If
    Else
      CanUnload = True
    End If
  Else
    CanUnload = True
  End If
End Function

Private Sub cmdNewFile_Click()
  Dim oDlg      As New CCanvasSizeDialog
  
  On Error Resume Next
  If Not LockUI() Then Exit Sub
  
  If CanUnload("New file") Then
    AppIniFile.Section = INISECTION_CANVAS
    oDlg.Rows = AppIniFile.GetInt(INIKEY_DEFCANVASROWS, mgrdCanvas.Rows)
    oDlg.Cols = AppIniFile.GetInt(INIKEY_DEFCANVASCOLS, mgrdCanvas.Cols)
    If oDlg.IIDialog.ShowDialog(True) Then
      If Not oDlg.IIDialog.Cancelled Then
        If oDlg.SaveAsDefaults Then
          AppIniFile.SetString INIKEY_DEFCANVASROWS, oDlg.Rows & ""
          AppIniFile.SetString INIKEY_DEFCANVASCOLS, oDlg.Cols & ""
        End If
        mgrdCanvas.Clear
        moCanvas.Filename = ""
        If mgrdCanvas.Resize(oDlg.Rows, oDlg.Cols) Then
          InitNewCanvas moCanvas
          CreateView
          UpdateView
          UpdateToolbar
        Else
          ShowUFError "Resizing the canvas failed", mgrdCanvas.LastErrDesc
        End If
        UpdateDialogTitle mgrdCanvas.Dirty
      End If
    End If
  End If
  Focus2Canvas
  
  UnlockUI
End Sub

Private Sub cmdSaveFileAs_Click()
  If Not LockUI() Then Exit Sub
  Call SaveToFile
  UnlockUI
End Sub

'Save As bitmap
Private Sub SaveAsBitmap(ByVal psFilename As String)
  'convert canvas to bitmap and show bitmap in imgCanvasBitmap
  Dim oBitmap As CBitmap
  Dim lWidth  As Long
  
  On Error GoTo SaveAsBitmap_Err
  
  lWidth = moConCanvas.CharWidth * mgrdCanvas.Cols
  Set oBitmap = New CBitmap

  DoHourglass True
  mgrdCanvas.ShowBoundaries = False
  RepositionProgressBar
  Me.IIProgressIndicator.BeginProgress "Rebuild w/o boundary markers"
  mgrdCanvas.LoadConsole moConCanvas, Me
  
  If Not oBitmap.SaveConsoleAsBitmap(moConCanvas, psFilename, 1, mgrdCanvas.Rows) Then
    DoHourglass False
    ShowUFError "Failed to save grabbed bitmap file", oBitmap.LastErrDesc
    DoHourglass True
  End If
  
  mgrdCanvas.ShowBoundaries = True
  RepositionProgressBar
  Me.IIProgressIndicator.BeginProgress "Rebuild with boundary markers"
  mgrdCanvas.LoadConsole moConCanvas, Me
  RefreshView
  DoHourglass False

SaveAsBitmap_Exit:
  Set oBitmap = Nothing
  Exit Sub

SaveAsBitmap_Err:
  ShowUFError "Failed to save canvas as bitmap", Err.Description
  Resume SaveAsBitmap_Exit
End Sub

Private Function SaveToFile() As Boolean
  Static iDumbCounter As Integer
  
  On Error Resume Next
  If iDumbCounter = 0 Then iDumbCounter = 1
  
  Dim fOK           As Boolean
  Dim fSaved        As Boolean
  Dim sFilename     As String
  Dim sExt          As String
  Dim sInitialDir   As String
  
  AppIniFile.GetOption INIOPT_LASTSAVEPATH, sInitialDir, ""
  sInitialDir = Trim$(sInitialDir)
  If Len(sInitialDir) = 0 Then
    sInitialDir = CombinePath(GetSpecialFolder(Application.hWndAccessApp, CSIDL_PERSONAL), APP_NAME)
  End If
  If Not ExistDir(sInitialDir) Then
    If Not CreatePath(sInitialDir) Then
      sInitialDir = GetSpecialFolder(Application.hWndAccessApp, CSIDL_PERSONAL)
    End If
  End If

  If Len(moCanvas.Filename) > 0 Then
    sFilename = StripFilePath(StripFileExt(moCanvas.Filename))
  End If

  fOK = VBGetSaveFileName( _
    sFilename, _
    "", _
    True, _
    "AsciiPaint (*.ascp)|*.ascp|Text files (*.txt, *.asc, *.vt100)|*.txt;*.asc;*.vt100|Bitmap (*.bmp)|*.bmp", _
    1, _
    sInitialDir, _
    "Save canvas as...", _
    "ascp", _
    Me.hWnd)
  If fOK Then
    sExt = LCase$(GetFileExt(sFilename))
    Me.IIProgressIndicator.BeginProgress "Save " & sExt & " file"
    Select Case sExt
    Case "ascp" 'native
      DoHourglass True
      fSaved = moCanvas.SaveNative(sFilename, Me) 'If fOK Then
      If Not fSaved Then
        ShowUFError "Failed to save console to file [" & sFilename & "]", moCanvas.LastErrDesc
      End If
      DoHourglass False
    Case "txt", "asc", "vt100"
      DoHourglass True
      fSaved = mgrdCanvas.SaveVT100(sFilename)
      DoHourglass False
      If Not fSaved Then
        ShowUFError "Failed to save console to file [" & sFilename & "]", mgrdCanvas.LastErrDesc
      End If
    Case "htm", "html"
      MsgBox APP_NAME & " coming soon... (" & sExt & ")", vbCritical
      GoTo cmdSaveFileAs_Click_exit
    Case "bmp"
      SaveAsBitmap sFilename
    Case Else
      MsgBox APP_NAME & " doesn't know how to save in this format (" & sExt & ")", vbCritical
      GoTo cmdSaveFileAs_Click_exit
    End Select
    moCanvas.Filename = sFilename
  End If
  
  If fSaved Then
    iDumbCounter = iDumbCounter + 1
    SaveToFile = True
    AppIniFile.SetOption INIOPT_LASTSAVEPATH, (StripFileName(moCanvas.Filename))
    UpdateDialogTitle mgrdCanvas.Dirty
  End If
cmdSaveFileAs_Click_exit:
  Focus2Canvas
  Me.IIProgressIndicator.EndProgress
End Function

Private Sub cmdLoadFile_Click()
  If Not LockUI() Then Exit Sub
  Call LoadFile
  UnlockUI
End Sub

Private Sub cmdInsertFile_Click()
  If Not LockUI() Then Exit Sub
  Call InsertFile
  UnlockUI
End Sub

Private Sub cmdLoadClipboard_Click()
  If Not LockUI() Then Exit Sub
  Call LoadClipboard
  UnlockUI
End Sub

Private Sub LoadFile()
  Dim fLoaded     As Boolean
  Dim sFilename   As String
  Dim iMaxRows    As Integer
  Dim iMaxCols    As Integer
  Dim sMsg        As String
  Dim sExt        As String
  
  Dim lstFilters  As CList
  Dim fOK         As Boolean
  
  On Error Resume Next
  If Not CanUnload("Load file") Then
    Exit Sub
  End If
  On Error GoTo LoadFile_Err
  
  Set lstFilters = NewSelectFileFilterList()
  lstFilters.AddValues "AsciiPaint", "*.ascp"
  lstFilters.AddValues "Text files", "*.txt;*.asc;*.vt100"
  fOK = SelectLoadFile(INIOPT_LASTLOADPATH, "Load File", lstFilters, sFilename)
  If Not fOK Then
    GoTo LoadFile_Exit
  End If
      
  sExt = LCase$(GetFileExt(sFilename))
  Me.IIProgressIndicator.BeginProgress "Load " & sExt & " file"
  
  Select Case sExt
  Case "ascp" 'native
    DoHourglass True
    fLoaded = moCanvas.LoadNative(sFilename, Me)
    If Not fLoaded Then
      ShowUFError "Failed to load native file [" & sFilename & "]", moCanvas.LastErrDesc
    End If
  Case "txt", "asc", "vt100"
    DoHourglass True
    fLoaded = mgrdCanvas.LoadVT100(sFilename, iMaxRows, iMaxCols)
    If fLoaded Then
      If (iMaxRows > mgrdCanvas.Rows) Or (iMaxCols > mgrdCanvas.Cols) Then
        sMsg = "The current canvas size (" & mgrdCanvas.Rows & " rows x " & mgrdCanvas.Cols & " cols)"
        sMsg = sMsg & " is smaller than the loaded file (" & iMaxRows & " rows x " & iMaxCols & " cols)"
        sMsg = sMsg & vbCrLf & vbCrLf & "The loaded file has been trimmed to the canvas size."
        MsgBox sMsg, vbExclamation, "Drawing size overflow"
      End If
    Else
      ShowUFError "Failed to load text file [" & sFilename & "]", mgrdCanvas.LastErrDesc
    End If
  Case Else
    MsgBox APP_NAME & " doesn't know how to load this file format (" & sExt & ")", vbCritical
    GoTo LoadFile_Exit
  End Select
  
  AppIniFile.SetOption INIOPT_LASTLOADPATH, (StripFileName(sFilename))
  
  RepositionProgressBar
  DoHourglass False

  If fLoaded Then
    moCanvas.Filename = sFilename
    UpdateRibbonWith moCanvas
    CreateView
    UpdateDialogTitle mgrdCanvas.Dirty
    UpdateToolbar
  Else
    mgrdCanvas.LoadConsole moConCanvas, Me
  End If
  
  UpdateView

LoadFile_Exit:
  On Error Resume Next
  Set lstFilters = Nothing
  Me.IIProgressIndicator.EndProgress
  DoHourglass False
  Exit Sub

LoadFile_Err:
  ShowUFError "An error occured while loading the file", Err.Description
  Resume LoadFile_Exit
End Sub

Private Sub LoadClipboard()
  Dim fLoaded     As Boolean
  Dim sFilename   As String
  Dim sMsg        As String
  Dim sExt        As String
  
  Dim lstFilters  As CList
  Dim fOK         As Boolean
  
  Dim oLoadCanvas As CCanvas
  
  On Error GoTo LoadClipboard_Err
  
  Set lstFilters = NewSelectFileFilterList()
  lstFilters.AddValues "AsciiPaint", "*.ascp"
  fOK = SelectLoadFile(INIOPT_CLIPLASTLOADPATH, "Load clipboard", lstFilters, sFilename)
  If Not fOK Then
    GoTo LoadClipboard_Exit
  End If
      
  sExt = LCase$(GetFileExt(sFilename))
  
  DoHourglass True
  Me.IIProgressIndicator.BeginProgress "Load " & sExt & " file"
  
  Set oLoadCanvas = New CCanvas
  
  Select Case sExt
  Case "ascp" 'native
    DoHourglass True
    fLoaded = oLoadCanvas.LoadNative(sFilename, Me)
    If Not fLoaded Then
      ShowUFError "Failed to load native file [" & sFilename & "]", oLoadCanvas.LastErrDesc
    End If
  Case Else
    MsgBox APP_NAME & " doesn't know how to load this file format (" & sExt & ")", vbCritical
    GoTo LoadClipboard_Exit
  End Select
  
  AppIniFile.SetOption INIOPT_LASTLOADPATH, (StripFileName(sFilename))
  
  If fLoaded Then
    Me.IIProgressIndicator.SetCaption "Merging"
    RefreshProgressBar
    
    If Not mgrdClipboard Is Nothing Then
      Set mgrdClipboard = Nothing
    End If
    Set mgrdClipboard = New CConsoleGrid
    
    fOK = mgrdClipboard.CreateFrom( _
            oLoadCanvas.ConsoleGrid, _
            1, _
            oLoadCanvas.ConsoleGrid.Rows, _
            1, _
            oLoadCanvas.ConsoleGrid.Cols _
          )
    DoHourglass False
    If fOK Then
      UpdateToolbar
    Else
      ShowUFError "Copy failed", mgrdClipboard.IIClassError.LastErrDesc
      Set mgrdClipboard = Nothing
    End If
  Else
    DoHourglass False
  End If
  
  RepositionProgressBar

LoadClipboard_Exit:
  On Error Resume Next
  Set oLoadCanvas = Nothing
  Set lstFilters = Nothing
  Me.IIProgressIndicator.EndProgress
  DoHourglass False
  Exit Sub

LoadClipboard_Err:
  ShowUFError "An error occured while loading the clipboard", Err.Description
  Resume LoadClipboard_Exit
End Sub

Private Function InsertNativeFile( _
    ByVal psFilename As String, _
    ByVal piRow As Integer, _
    ByVal piCol As Integer _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "InsertNativeFile"
  Dim grdTemp As CConsoleGrid
  Dim fOK     As Boolean
  Dim fTransp As Boolean
  
  On Error GoTo InsertNativeFile_Err
  
  Set grdTemp = New CConsoleGrid
  fTransp = CBool(GetKeyState(VK_SHIFT) < 0)
  fOK = grdTemp.LoadNative(psFilename, pfAutoResize:=False)
  If Not fOK Then
    ShowUFError "Failed to insert native file [" & psFilename & "]", grdTemp.LastErrDesc
    GoTo InsertNativeFile_Exit
  End If
  
  fOK = mgrdCanvas.Paste(grdTemp, piRow, piCol, fTransp)
  If Not fOK Then
    ShowUFError "Failed to paste native file [" & psFilename & "]", grdTemp.LastErrDesc
    GoTo InsertNativeFile_Exit
  End If
  
  InsertNativeFile = fOK
  
InsertNativeFile_Exit:
  Set grdTemp = Nothing
  Exit Function
InsertNativeFile_Err:
  ShowUFError "Failed to insert native file [" & psFilename & "]", Err.Description
  Resume InsertNativeFile_Exit
End Function

Private Function InsertVT100File( _
    ByVal psFilename As String, _
    ByVal piRow As Integer, _
    ByVal piCol As Integer _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "InsertVT100File"
  Dim fOK     As Boolean
  Dim fTransp As Boolean
  Dim iMaxRows    As Integer
  Dim iMaxCols    As Integer
  
  On Error GoTo InsertVT100File_Err
  
  fTransp = CBool(GetKeyState(VK_SHIFT) < 0)
  fOK = mgrdCanvas.LoadVT100(psFilename, iMaxRows, iMaxCols, piRow, piCol, False, fTransp)
  If Not fOK Then
    ShowUFError "Failed to insert text file [" & psFilename & "]", mgrdCanvas.LastErrDesc
  End If
  
  InsertVT100File = fOK
  
InsertVT100File_Exit:
  Exit Function
InsertVT100File_Err:
  ShowUFError "Failed to insert native file [" & psFilename & "]", Err.Description
  Resume InsertVT100File_Exit
End Function
  
Private Sub InsertFile()
  Dim fLoaded     As Boolean
  Dim sFilename   As String
  Dim iCurRow     As Integer
  Dim iCurCol     As Integer
  Dim sExt        As String
  
  Dim lstFilters  As CList
  Dim fOK         As Boolean
  
  On Error GoTo InsertFile_Err
  
  Call moConCanvas.GetCaretPos(iCurRow, iCurCol)
  
  Set lstFilters = NewSelectFileFilterList()
  lstFilters.AddValues "AsciiPaint", "*.ascp"
  lstFilters.AddValues "Text files", "*.txt;*.asc;*.vt100"
  fOK = SelectLoadFile(INIOPT_LASTINSERTPATH, "Insert File", lstFilters, sFilename)
  If Not fOK Then
    GoTo InsertFile_Exit
  End If
      
  sExt = LCase$(GetFileExt(sFilename))
  Me.IIProgressIndicator.BeginProgress "Load " & sExt & " file"
  
  Select Case sExt
  Case "ascp" 'native
    DoHourglass True
    fLoaded = InsertNativeFile(sFilename, iCurRow, iCurCol)
  Case "txt", "asc", "vt100"
    DoHourglass True
    fLoaded = InsertVT100File(sFilename, iCurRow, iCurCol)
  Case Else
    MsgBox APP_NAME & " doesn't know how to load this file format (" & sExt & ")", vbCritical
    GoTo InsertFile_Exit
  End Select
  
  AppIniFile.SetOption INIOPT_LASTINSERTPATH, (StripFileName(sFilename))
  
  RepositionProgressBar
  DoHourglass False

  If fLoaded Then
    UpdateRibbonWith moCanvas
    CreateView
    UpdateDialogTitle mgrdCanvas.Dirty
    UpdateToolbar
  Else
    mgrdCanvas.LoadConsole moConCanvas, Me
  End If
  
  UpdateView

InsertFile_Exit:
  On Error Resume Next
  Set lstFilters = Nothing
  Me.IIProgressIndicator.EndProgress
  DoHourglass False
  Exit Sub

InsertFile_Err:
  ShowUFError "An error occured while inserting the file", Err.Description
  Resume InsertFile_Exit
End Sub

Private Sub UpdateRibbonWith(poCanvas As CCanvas)
  On Error Resume Next
  Me.txtFontSize = moCanvas.FontSize
  Me.cboFontName = moCanvas.FontName
  Me.txtTypeText.FontName = moCanvas.FontName
  Me.optTransparencyMode = moCanvas.TransparencyMode
  Me.txtTransAlphaPct = moCanvas.TransparentAlphaPct
      
  Me.txtLinePaddingTop = CStr(moCanvas.LinePaddingTop)
  Me.txtLinePaddingBottom = CStr(moCanvas.LinePaddingBottom)
  Me.txtLineSpacingTop = CStr(moCanvas.LineSpacingTop)
  Me.txtLineSpacingBottom = CStr(moCanvas.LineSpacingBottom)
  Me.chkAutoAdjustWidth = moCanvas.AutoAdjustOnCharWidth
  
  SetForeColor moCanvas.ForeColor
  SetBackColor moCanvas.BackColor
  SetCanvasTransparentColor moCanvas.TransparentColor
  SetCanvasAlphaTransparencyPct moCanvas.TransparentAlphaPct
  
  On Error Resume Next
  Forms(GetCharMapFormName()).WarnOnCanvasUnsync 'V02.00.00
End Sub

Public Sub ColorsToPalette(ByRef poPalette As CColorPalette)
  On Error Resume Next
  mgrdCanvas.ColorsToPalette poPalette
End Sub

'
' "clipboard"
'
Private Sub cmdCut_Click()
  On Error GoTo cmdCut_Click_Err
  Dim fOK       As Boolean
  Dim i         As Integer
  
  On Error Resume Next
  
  If Not LockUI() Then Exit Sub
  
  If Not HasSelection() Then
    Exit Sub
  End If
  If Not mgrdClipboard Is Nothing Then
    Set mgrdClipboard = Nothing
  End If
  Set mgrdClipboard = New CConsoleGrid
  
  DoHourglass True
  fOK = mgrdClipboard.CreateFrom(mgrdCanvas, moCanvas.SelStartRow, moCanvas.SelEndRow, moCanvas.SelStartCol, moCanvas.SelEndCol)
  If fOK Then
    For i = 0 To moCanvas.SelEndRow - moCanvas.SelStartRow
      fOK = mgrdCanvas.ClearLine(moCanvas.SelStartRow + i, moCanvas.SelStartCol, eFromCursorPosition, True, True, True, moCanvas.SelEndCol)
    Next i
    RepositionProgressBar
    mgrdCanvas.LoadConsole moConCanvas, Me
    RefreshView
    UpdateToolbar
    OnCaretPosChange
  Else
    ShowUFError "Copy failed", mgrdClipboard.IIClassError.LastErrDesc
    Set mgrdClipboard = Nothing
  End If
  
cmdCut_Click_Exit:
  Focus2Canvas
  DoHourglass False
  UpdateToolbar
  UnlockUI
  Exit Sub

cmdCut_Click_Err:
  ShowUFError "Cut failed", Err.Description
  Resume cmdCut_Click_Exit
End Sub

Private Sub cmdCopy_Click()
  On Error GoTo cmdCopy_Click_Err
  Dim fOK       As Boolean
  
  If Not LockUI() Then Exit Sub
  
  If Not HasSelection() Then
    UnlockUI
    Exit Sub
  End If
  
  If Not mgrdClipboard Is Nothing Then
    Set mgrdClipboard = Nothing
  End If
  Set mgrdClipboard = New CConsoleGrid
  
  DoHourglass True
  fOK = mgrdClipboard.CreateFrom(mgrdCanvas, moCanvas.SelStartRow, moCanvas.SelEndRow, moCanvas.SelStartCol, moCanvas.SelEndCol)
  If Not fOK Then
    ShowUFError "Copy failed", mgrdClipboard.IIClassError.LastErrDesc
    Set mgrdClipboard = Nothing
  End If
  
cmdCopy_Click_Exit:
  Focus2Canvas
  DoHourglass False
  UpdateToolbar
  UnlockUI
  Exit Sub

cmdCopy_Click_Err:
  ShowUFError "Copy failed", Err.Description
  Resume cmdCopy_Click_Exit
End Sub

Private Sub cmdPaste_Click()
  On Error GoTo cmdPaste_Click_Err
  Dim fTransparent      As Boolean
  Dim fOK               As Boolean
  Dim iRow              As Integer
  Dim iCol              As Integer
  
  If mgrdClipboard Is Nothing Then
    Beep
    Exit Sub
  End If
  
  If Not LockUI() Then Exit Sub
  
  moConCanvas.GetCaretPos iRow, iCol
  If (iRow <= 0) Or (iCol <= 0) Then
    Beep
    GoTo cmdPaste_Click_Exit
  End If
  
  DoHourglass True
  fTransparent = CBool(GetKeyState(VK_SHIFT))
  fOK = mgrdCanvas.Paste(mgrdClipboard, iRow, iCol, fTransparent)
  If fOK Then
    RepositionProgressBar
    mgrdCanvas.LoadConsole moConCanvas, Me
    RefreshView
    UpdateToolbar
    OnCaretPosChange
  Else
    ShowUFError "Paste failed", mgrdCanvas.IIClassError.LastErrDesc
  End If
  
cmdPaste_Click_Exit:
  Focus2Canvas
  DoHourglass False
  UnlockUI
  Exit Sub

cmdPaste_Click_Err:
  ShowUFError "Paste failed", Err.Description
  Resume cmdPaste_Click_Exit
End Sub

'Set top,left or bottom,right selection anchor
Private Function SetSelectionAnchor(ByVal peEdge As eSelAnchorEdge) As Boolean
  Dim iRow              As Integer
  Dim iCol              As Integer
  
  moConCanvas.GetCaretPos iRow, iCol
  If (iRow <= 0) Or (iCol <= 0) Then
    Beep
    Exit Function
  End If
  
  If peEdge = eSelBegin Then
    moCanvas.SelStartRow = iRow
    moCanvas.SelStartCol = iCol
    If moCanvas.SelEndRow = 0 Then
      moCanvas.SelEndRow = iRow
      moCanvas.SelEndCol = iCol
    End If
  Else
    moCanvas.SelEndRow = iRow
    moCanvas.SelEndCol = iCol
    If moCanvas.SelStartRow = 0 Then
      moCanvas.SelStartRow = iRow
      moCanvas.SelStartCol = iCol
    End If
  End If
  
  If (moCanvas.SelEndRow < moCanvas.SelStartRow) And (moCanvas.SelEndCol < moCanvas.SelStartCol) Then
    'swap row and column
    iRow = moCanvas.SelEndRow
    iCol = moCanvas.SelEndCol
    moCanvas.SelEndRow = moCanvas.SelStartRow
    moCanvas.SelEndCol = moCanvas.SelStartCol
    moCanvas.SelStartRow = iRow
    moCanvas.SelStartCol = iCol
  ElseIf moCanvas.SelEndRow < moCanvas.SelStartRow Then
    'swap rows
    iRow = moCanvas.SelEndRow
    moCanvas.SelEndRow = moCanvas.SelStartRow
    moCanvas.SelStartRow = iRow
  ElseIf moCanvas.SelEndCol < moCanvas.SelStartCol Then
    'swap columns
    iCol = moCanvas.SelEndCol
    moCanvas.SelEndCol = moCanvas.SelStartCol
    moCanvas.SelStartCol = iCol
  End If
  
  SetSelectionAnchor = True
End Function

Private Function HasSelection() As Boolean
  'HasSelection = Not CBool((moCanvas.SelStartRow <= 0) Or (moCanvas.SelStartCol <= 0) Or (moCanvas.SelEndRow <= 0) Or (moCanvas.SelEndCol <= 0))
  HasSelection = CBool((moCanvas.SelStartRow > 0) And (moCanvas.SelStartCol > 0) And (moCanvas.SelEndRow > 0) Or (moCanvas.SelEndCol > 0))
End Function

Private Function MemoryUsedAsText(ByVal plBytes As Long) As String
  Dim sUnit As String
  Dim sRet  As String
  
  If plBytes > 1024& Then
    sRet = Format$(plBytes / 1024, "#.00")
    sUnit = "Kb"
  Else
    sRet = CStr(plBytes)
    sUnit = "bytes"
  End If
  sRet = sRet & " " & sUnit
  MemoryUsedAsText = sRet
End Function

Private Sub UpdateCanvasSizeIndics()
  On Error Resume Next
  Me.lblCanvasSize.Caption = mgrdCanvas.Rows & " x " & mgrdCanvas.Cols
  Me.lblCanvasRam.Caption = MemoryUsedAsText(mgrdCanvas.GetMemoryFootprint())
End Sub

Private Sub UpdateSelectionIndicatorText()
  On Error Resume Next
  Dim sCaption As String
  Dim sMemory As String
  Dim lMemUsed As Long
  
  If Not mgrdClipboard Is Nothing Then
    sCaption = mgrdClipboard.Rows & " x " & mgrdClipboard.Cols
    lMemUsed = mgrdClipboard.GetMemoryFootprint()
  Else
    sCaption = "(empty)"
  End If
  
  sMemory = MemoryUsedAsText(lMemUsed)
  sCaption = sCaption & " (" & sMemory & ")"
  Me.lblClipboardSize.Caption = "Clipboard " & sCaption
End Sub

Public Sub ShowSelection()
  On Error Resume Next
  
  Dim fEnabled As Boolean
  
  'show selection here
  Me.lblSelStart.Caption = moCanvas.SelStartRow & "," & moCanvas.SelStartCol
  Me.lblSelEnd.Caption = moCanvas.SelEndRow & "," & moCanvas.SelEndCol
  
  fEnabled = HasSelection()
  SetButtonEnabled Me.cmdCopy, fEnabled
  SetButtonEnabled Me.cmdCut, fEnabled
  fEnabled = Not mgrdClipboard Is Nothing
  SetButtonEnabled Me.cmdPaste, fEnabled
  UpdateSelectionIndicatorText
End Sub

Private Sub cmdSelStart_Click()
  If SetSelectionAnchor(eSelBegin) Then
    ShowSelection
  End If
  Focus2Canvas
End Sub

Private Sub cmdSelEnd_Click()
  If SetSelectionAnchor(eSelEnd) Then
    ShowSelection
  End If
  Focus2Canvas
End Sub

Private Sub OnTxtTypeTextChange()
  On Error Resume Next
  Dim iLen    As Integer
  iLen = Len(Me.txtTypeText.Text)
  If iLen > mgrdCanvas.Cols Then
    Me.lblTypeTextLen.ForeColor = vbRed
  Else
    Me.lblTypeTextLen.ForeColor = Me.lblCanvasSize.ForeColor
  End If
  Me.lblTypeTextLen.Caption = CStr(iLen)
  Me.lblTypeTextLen.Visible = CBool(iLen > 0)
End Sub

Private Sub txtLinePaddingBottom_AfterUpdate()
  DoHourglass True
  moConCanvas.LinePadding(elsBottom) = Val(Nz(Me.txtLinePaddingBottom, 0))
  RefreshView
  moConCanvas.RefreshWindow
  Focus2Canvas
  DoHourglass False
End Sub

Private Sub txtLinePaddingBottom_Enter()
  DisableFormKeyPreview
End Sub

Private Sub txtLinePaddingBottom_Exit(Cancel As Integer)
  EnableFormKeyPreview
End Sub

Private Sub txtLinePaddingBottom_KeyDown(KeyCode As Integer, Shift As Integer)
  If Not IsIntegerEditKeyCodeAllowed(KeyCode) And (KeyCode <> vbKeySubtract) Then
    KeyCode = 0
  End If
End Sub

Private Sub txtLinePaddingTop_AfterUpdate()
  DoHourglass True
  moConCanvas.LinePadding(elsTop) = Val(Nz(Me.txtLinePaddingTop, 0))
  RefreshView
  moConCanvas.RefreshWindow
  Focus2Canvas
  DoHourglass False
End Sub

Private Sub txtLinePaddingTop_Enter()
  DisableFormKeyPreview
End Sub

Private Sub txtLinePaddingTop_Exit(Cancel As Integer)
  EnableFormKeyPreview
End Sub

Private Sub txtLinePaddingTop_KeyDown(KeyCode As Integer, Shift As Integer)
  If Not IsIntegerEditKeyCodeAllowed(KeyCode) And (KeyCode <> vbKeySubtract) Then
    KeyCode = 0
  End If
End Sub

Private Sub txtLineSpacingBottom_AfterUpdate()
  DoHourglass True
  moConCanvas.LineSpacing(elsBottom) = Val(Nz(Me.txtLineSpacingBottom, 0))
  RefreshView
  moConCanvas.RefreshWindow
  Focus2Canvas
  DoHourglass False
End Sub

Private Sub txtLineSpacingBottom_Enter()
  DisableFormKeyPreview
End Sub

Private Sub txtLineSpacingBottom_Exit(Cancel As Integer)
  EnableFormKeyPreview
End Sub

Private Sub txtLineSpacingBottom_KeyDown(KeyCode As Integer, Shift As Integer)
  If Not IsIntegerEditKeyCodeAllowed(KeyCode) And (KeyCode <> vbKeySubtract) Then
    KeyCode = 0
  End If
End Sub

Private Sub txtLineSpacingTop_AfterUpdate()
  DoHourglass True
  moConCanvas.LineSpacing(elsTop) = Val(Nz(Me.txtLineSpacingTop, 0))
  RefreshView
  moConCanvas.RefreshWindow
  Focus2Canvas
  DoHourglass False
End Sub

Private Sub txtLineSpacingTop_Enter()
  DisableFormKeyPreview
End Sub

Private Sub txtLineSpacingTop_Exit(Cancel As Integer)
  EnableFormKeyPreview
End Sub

Private Sub txtLineSpacingTop_KeyDown(KeyCode As Integer, Shift As Integer)
  If Not IsIntegerEditKeyCodeAllowed(KeyCode) And (KeyCode <> vbKeySubtract) Then
    KeyCode = 0
  End If
End Sub

Private Sub chkAutoAdjustWidth_AfterUpdate()
  moConCanvas.AutoAdjustWidth = Me.chkAutoAdjustWidth
  RefreshView
  moConCanvas.RefreshWindow
  Focus2Canvas
End Sub

Private Sub chkFragmentText_Click()
  If Not LockUI() Then Exit Sub
  DoHourglass True
  mgrdCanvas.SetFragmentText Me.chkFragmentText, moConCanvas.CharWidth
  
  CreateView
  UpdateView
  Focus2Canvas
  UnlockUI
  DoHourglass False
End Sub

Private Sub txtTransAlphaPct_AfterUpdate()
  moCanvas.TransparentAlphaPct = Val(Me.txtTransAlphaPct)
  SetCanvasTransparency mfTransparent, moCanvas.TransparencyMode
End Sub

Private Sub txtTransAlphaPct_Enter()
  DisableFormKeyPreview
End Sub

Private Sub txtTransAlphaPct_Exit(Cancel As Integer)
  EnableFormKeyPreview
End Sub

Private Sub txtTypeText_Change()
  OnTxtTypeTextChange
End Sub

'V02.00.00 Canvas Transparency methods

Private Sub SetCanvasTransparentColor(ByVal plTransColor As Long)
  moCanvas.TransparentColor = plTransColor
  'Now apply COLOR transparency to canvas if we're in transparent mode
  If mfTransparent Then
    If moCanvas.TransparencyMode = eColorTransparency Then
      'moConCanvas.SetAlphaTransparency 0
      moConCanvas.SetColorTransparency moCanvas.TransparentAlphaPct, moCanvas.TransparentColor, True
    End If
  End If
End Sub

Private Sub SetCanvasAlphaTransparencyPct(ByVal piTransPct As Integer)
  moCanvas.TransparentAlphaPct = piTransPct
  'Now apply ALPHA transparency to canvas if we're in transparent mode
  If mfTransparent Then
    If moCanvas.TransparencyMode = eAlphaTransparency Then
      'moConCanvas.SetColorTransparency miTransAlphaPct, mlTransColor, False
      moConCanvas.SetAlphaTransparency piTransPct
    End If
  End If
End Sub

Private Sub SetCanvasTransparency(ByVal pfTransparent As Boolean, ByVal peTransparencyMode As eTransparencyMode)
  On Error Resume Next
  'Debug.Print "SetCanvasTransparency(" & pfTransparent & ", " & peTransparencyMode & ")"
  If pfTransparent Then
    If peTransparencyMode = eColorTransparency Then
      'Debug.Print "SetColorTransparency(" & miTransAlphaPct & ", " & Hex$(mlTransColor) & ", -1)"
      moConCanvas.SetColorTransparency moCanvas.TransparentAlphaPct, moCanvas.TransparentColor, True
    Else
      'Debug.Print "SetAlphaTransparency(" & miTransAlphaPct & ")"
      moConCanvas.SetAlphaTransparency moCanvas.TransparentAlphaPct
    End If
  Else
    'turn off transparency
    moConCanvas.SetAlphaTransparency 0
  End If
  mfTransparent = pfTransparent
End Sub

Private Sub SetCanvasBackgroundImageFromFile(ByVal psFilename As String)
  On Error Resume Next
  Dim fVisible As Boolean
  psFilename = Trim$(psFilename)
  If Len(psFilename) > 0 Then
    Me.imgBackground.Picture = psFilename
    If Err.Number = 0 Then
      msTransImageFilename = psFilename
      Me.txtBkgndBitmap = psFilename
      Me.txtBkgndBitmap.BackColor = Me.txtTypeText.BackColor
      fVisible = True
    Else
      Me.txtBkgndBitmap.BackColor = GetAlertBackgroundColor()
      'ShowUFError "Failed to load background image", Err.Description
    End If
  Else
    msTransImageFilename = ""
    Me.txtBkgndBitmap = ""
    Me.imgBackground.Picture = ""
    fVisible = False
  End If
  Me.txtBkgndBitmap = msTransImageFilename
  Me.imgBackground.Visible = fVisible
  moConCanvas.BringWindowToTop
End Sub

Private Sub SetTransColor(ByVal plTransColor As Long)
  moCanvas.TransparentColor = plTransColor
  Me.rectTransColor.BackStyle = 1
  Me.rectTransColor.BackColor = plTransColor
End Sub

'V02.00.02 Selection indicator with transparent console overlay

Private Function CreateSelectionWindow() As Boolean
  Const LOCAL_ERR_CTX As String = "CreateSelectionWindow"
  On Error GoTo CreateSelectionWindow_Err
  
  Set moConSel = New CConsoul
  moConSel.BackColor = Nz(moCanvas.BackColor, 0&)
  'Create the console window and tell the library that we want click feedback
  If Not moConSel.Attach(Me.hWnd, 0, 0, 0, 0, piCreateAttributes:=LW_RENDERMODEBYLINE) Then
    MsgBox "Failed to create selection window", vbCritical
    GoTo CreateSelectionWindow_Exit
  End If
  CreateSelectionWindow = True
  
CreateSelectionWindow_Exit:
  Exit Function

CreateSelectionWindow_Err:
  ShowUFError "Failed to create selection window", Err.Description
  Resume CreateSelectionWindow_Exit
End Function

Private Function SetSelectionWindowCoords(ByVal piLeft As Integer, ByVal piTop As Integer, ByVal piWidth As Integer, ByVal piHeight As Integer)
  'Debug.Print "SetSelectionWindowCoords"; piLeft, piTop, piWidth, piHeight
  miConSelLeft = piLeft
  miConSelTop = piTop
  miConSelWidth = piWidth
  miConSelHeight = piHeight
End Function

Private Function ShowSelectionWindow(ByVal pfVisible As Boolean)
  If moConSel Is Nothing Then
    If Not CreateSelectionWindow() Then
      Exit Function
    End If
  End If
  
  Me.chkTransparency.Value = False
  
  If pfVisible Then
    moConSel.Clear
    moConSel.OutputLn "(Selection)"
    If Not moConCanvas Is Nothing Then
      moConSel.SetWindowPos moConCanvas.hWnd, miConSelLeft, miConSelTop, miConSelWidth, miConSelHeight, HWND_TOP
      moConSel.SetAlphaTransparency 50
    End If
    moConSel.ShowWindow True
  Else
    moConSel.ShowWindow False
  End If
End Function

Private Function ICsWmPaintEventSink_OnConsolePaint(ByVal phWnd As Long, ByVal pwCbkMode As Integer, ByVal phDC As LongPtr, ByVal lprcLinePos As Long, ByVal lprcLineRect As Long, ByVal lprcPaint As Long) As Integer
  If phWnd = moConCanvas.hWnd Then
    If mfShowGrid Then
      If pwCbkMode = WMPAINTCBK_AFTER Then
        ICsWmPaintEventSink_OnConsolePaint = OnPaintConsoleGrid(moConCanvas, phWnd, phDC, lprcLinePos, lprcLineRect, lprcPaint)
      End If
    End If
  End If
End Function

Private Function ShowConsoleGrid(ByVal pfShow As Boolean)
  On Error Resume Next
  mfShowGrid = pfShow
  moConCanvas.RefreshWindow
  AppIniFile.SetString INIKEY_SHOWGRID, Abs(CInt(mfShowGrid)) & ""
End Function

Private Sub PositionRibbonControls()
  Dim iTop        As Integer
  iTop = Me.rectRibRef1.Top
  moRibbon.PositionBandControls iTop
End Sub

Private Sub txtTypeText_Enter()
  DisableFormKeyPreview
End Sub

Private Sub txtTypeText_Exit(Cancel As Integer)
  EnableFormKeyPreview
End Sub

Private Sub Focus2Canvas()
  On Error Resume Next
  Me.cmdDummy.SetFocus
End Sub

Private Property Get IMessageReceiver_ClientID() As String
  IMessageReceiver_ClientID = Me.Name
End Property

Private Function IMessageReceiver_OnMessageReceived(ByVal psSenderID As String, ByVal psTopic As String, pvData As Variant) As Long
  Dim fOK     As Boolean
  Select Case psTopic
  Case MSGTOPIC_LOCKUI
    Focus2Canvas
    fOK = MForms.FormSetAllowEdits(Me, False, "", "", 0)
    Focus2Canvas
  Case MSGTOPIC_UNLOCKUI
    fOK = MForms.FormSetAllowEdits(Me, True, "", "", 0)
  Case MSGTOPIC_CANUNLOAD
    fOK = CanUnload(pvData)
    If fOK Then
      IMessageReceiver_OnMessageReceived = 0&
    Else
      IMessageReceiver_OnMessageReceived = 1& 'breaks the broadcast chain
    End If
  Case MSGTOPIC_CHARSELECTED
    TypeChar CLng(pvData) 'pvData is the char code as a Long
  End Select
End Function

'
' Debug console
'

Private Sub InitDebugConsole()
  AppIniFile.Section = INISECTION_DEBUGCONSOLE
  miDbgWindowLineCt = AppIniFile.GetInt(INIKEY_DISPLAYLINECT, 8)
  miDbgCharsPerLine = AppIniFile.GetInt(INIKEY_CHARSPERLINE, 18)
End Sub

Private Function CreateDebugConsole() As Boolean
  On Error GoTo CreateDebugConsole_Err
  
  Set mconDebug = New CConsoul
  mconDebug.FontName = "Lucida Console"
  mconDebug.FontSize = 12
  mconDebug.MaxCapacity = 5000
  mconDebug.ForeColor = RGB(200, 200, 240)
  mconDebug.BackColor = vbBlack
  'Create the console window and tell the library that we want click feedback
  If Not mconDebug.Attach(Me.hWnd, 0, 0, 0, 0, 0, piCreateAttributes:=LW_RENDERMODEBYLINE) Then
    MsgBox "Failed to create debug console window", vbCritical
    GoTo CreateDebugConsole_Exit
  End If
  mconDebug.LineSpacing(elsTop) = 2
  CreateDebugConsole = True
  
CreateDebugConsole_Exit:
  Exit Function

CreateDebugConsole_Err:
  ShowUFError "Failed to create the debug console", Err.Description
  Resume CreateDebugConsole_Exit
End Function

Private Sub DestroyDebugConsole()
  If mconDebug Is Nothing Then Exit Sub
  mconDebug.Detach
  Set mconDebug = Nothing
End Sub

Private Sub ShowDebugConsole()
  If mconDebug Is Nothing Then
    If Not CreateDebugConsole() Then Exit Sub
  End If
  mconDebug.ShowWindow True
  mfConDebugVisible = True
  Form_Resize
  DoEvents
  RepositionConsouls
End Sub

Private Sub HideDebugConsole()
  If mconDebug Is Nothing Then Exit Sub
  mconDebug.ShowWindow False
  mfConDebugVisible = False
  RepositionConsouls
End Sub

Private Sub cmdDebugConsole_Click()
  If Not mfConDebugVisible Then
    ShowDebugConsole
    UpdateDebugConsole True
  Else
    HideDebugConsole
  End If
End Sub

Private Sub UpdateDebugConsole(ByVal pfForceUpdate As Boolean)
  If Not mfConDebugVisible Then Exit Sub
  
  On Error GoTo UpdateDebugConsole_Err
  
  Dim iCol            As Integer
  Dim iRow            As Integer
  Dim sVT100          As String
  Dim iLenVT100       As Integer
  Dim iDbgLineCt      As Integer
  Dim iLineCt         As Integer
  Dim sHex            As String
  Dim sLine           As String
  Dim i               As Integer
  Dim sChar           As String
  Dim k               As Integer
  Dim iAscW           As Integer
  
  moConCanvas.GetCaretPos iRow, iCol
  If (iRow <= 0) Or (iCol <= 0) Then
    Exit Sub
  End If
  If (Not pfForceUpdate) And (iRow = miLastDbgRow) Then
    Exit Sub
  End If
  miLastDbgRow = iRow
  
  mconDebug.Clear
  
  'dump current consolegrid contents
  sVT100 = mgrdCanvas.GetLineVT100(iRow, False)
  iLenVT100 = Len(sVT100)

  iDbgLineCt = iLenVT100 \ miDbgCharsPerLine
  If (iLenVT100 Mod miDbgCharsPerLine) > 0 Then
    iDbgLineCt = iDbgLineCt + 1
  End If

  k = 1
  i = 1
  iLineCt = 1
  Do
    If k <= iLenVT100 Then
      sChar = Mid$(sVT100, k, 1)
      iAscW = AscW(sChar)
      If (iAscW < 32) Or (iAscW > 128) Then
        sLine = sLine & VT_FCOLOR(vbWhite) & "." & VT_RESET()
        sHex = sHex & VT_FCOLOR(vbWhite) & HexInt(iAscW) & VT_RESET() & " "
      Else
        sLine = sLine & sChar
        sHex = sHex & HexInt(iAscW) & " "
      End If
    Else
      sLine = sLine & " "
      sHex = sHex & "     "
    End If
    
    k = k + 1
    i = i + 1
    If i > miDbgCharsPerLine Then
      mconDebug.OutputLn Format$(iLineCt, "000") & ": " & sHex & sLine
      sHex = ""
      sLine = ""
      iLineCt = iLineCt + 1
      i = 1
    Else
      sLine = sLine & " "
    End If
  Loop Until iLineCt > iDbgLineCt
  mconDebug.ScrollTop
  
UpdateDebugConsole_Exit:
  Exit Sub
UpdateDebugConsole_Err:
  Debug.Print Err.Description
  Resume UpdateDebugConsole_Exit
End Sub
