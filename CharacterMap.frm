VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CharacterMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements ICsMouseEventSink
Implements IMessageReceiver

'The dialog data class that we'll link on open
Private moDialog As CCharMapDialog

'The console objects
' moConCharset    Displays the character set page and the selection
Private moConCharset    As CConsoul
'ROWS and COLS for the character map display grid
Private Const Rows  As Long = 6&
Private Const Cols  As Long = 10&
'Each character cell is CELL_SIZE widht and height
Private Const CELL_SIZE As Long = 3&

'Display character set
Private msFontName  As String
Private miFontSize  As Integer
Private mlStart     As Long   'The character code index of the first displayed character (top/left)
Private malCodes()  As Long
Private mlCodesCt   As Long
'Selected item
Private miSelCol    As Integer
Private miSelRow    As Integer
Private miSelZone   As Integer

Private mrcClient   As RECT 'compute on Form_Resize and used in RepositionConsouls

Private mbDispCharCodeToggle  As Byte

Private miFixedCharWidth      As Integer

Private mlComboBackColor      As Long

Private Sub OnPageUp(Optional ByVal Shift As Integer = 0)
  mlStart = mlStart - Rows * Cols
  If mlStart < 1& Then mlStart = 1&
  DisplayPage True
End Sub

Private Sub OnPageDown(Optional ByVal Shift As Integer = 0)
  'If (mlStart + Rows * Cols) <= mlCodesCt Then
  If (mlCodesCt - (mlStart + Rows * Cols)) > 0 Then
    mlStart = mlStart + Rows * Cols
    DisplayPage True
  End If
End Sub

Private Sub OnCursorKeyLeft(Optional ByVal Shift As Integer = 0)
  SelectChar False
  If miSelRow > 1 Then
    If miSelCol > 1 Then
      miSelCol = miSelCol - 1
    Else
      miSelCol = Cols
      miSelRow = miSelRow - 1
    End If
  Else
    If miSelCol > 1 Then
      miSelCol = miSelCol - 1
    End If
  End If
  miSelZone = miSelRow * 256 + miSelCol '16 bits word
  SelectChar True
End Sub

Private Sub OnCursorKeyRight(Optional ByVal Shift As Integer = 0)
  Dim lIndex As Long
  SelectChar False
  If miSelCol < Cols Then
    lIndex = CharIndexOf(miSelRow, miSelCol + 1)
    If (lIndex > 0) And (lIndex <= mlCodesCt) Then
      miSelCol = miSelCol + 1
    End If
  Else
    If miSelRow < Rows Then
      lIndex = CharIndexOf(miSelRow + 1, 1)
      If (lIndex > 0) And (lIndex <= mlCodesCt) Then
        miSelCol = 1
        miSelRow = miSelRow + 1
      End If
    End If
  End If
  miSelZone = miSelRow * 256 + miSelCol '16 bits word
  SelectChar True
End Sub

Private Sub OnCursorKeyUp(Optional ByVal Shift As Integer = 0)
  SelectChar False
  If miSelRow > 1 Then
    miSelRow = miSelRow - 1
  End If
  miSelZone = miSelRow * 256 + miSelCol '16 bits word
  SelectChar True
End Sub

Private Sub OnCursorKeyDown(Optional ByVal Shift As Integer = 0)
  Dim lIndex As Long
  SelectChar False
  If miSelRow < Rows Then
    lIndex = CharIndexOf(miSelRow + 1, miSelCol)
    If (lIndex > 0) And (lIndex <= mlCodesCt) Then
      miSelRow = miSelRow + 1
    End If
  End If
  miSelZone = miSelRow * 256 + miSelCol '16 bits word
  SelectChar True
End Sub

Private Sub OnKeyHome(Optional ByVal Shift As Integer = 0)
  If Shift = 2 Then 'ctrl
    mlStart = 1
  End If
  SelectChar False
  miSelRow = 1
  miSelCol = 1
  miSelZone = miSelRow * 256 + miSelCol '16 bits word
  If Shift = 2 Then DisplayPage True
  SelectChar True
End Sub

Private Sub OnKeyEnd(Optional ByVal Shift As Integer = 0)
  Dim lRemain As Long
  Dim lLastPg As Long
  
  lRemain = mlCodesCt Mod (Rows * Cols)
  lLastPg = (mlCodesCt \ (Rows * Cols)) * (Rows * Cols)
  
  If Shift = 2 Then 'ctrl
    If lRemain > 0 Then
      mlStart = mlCodesCt - lRemain + 1&
    Else
      mlStart = mlCodesCt - (Rows * Cols) + 1
    End If
  End If
  
  SelectChar False
  If mlStart > lLastPg Then
    miSelRow = (lRemain \ Cols) + 1
    miSelCol = lRemain Mod Cols
  Else
    miSelRow = Rows
    miSelCol = Cols
  End If
  miSelZone = miSelRow * 256 + miSelCol '16 bits word
  If Shift = 2 Then DisplayPage True
  SelectChar True
End Sub

Private Sub cboFontName_AfterUpdate()
  Dim sFontName   As String
  sFontName = Trim$(cboFontName & "")
  If Len(sFontName) = 0 Then Exit Sub
  msFontName = sFontName
  RecreateConsole
  RepositionConsouls
  DisplayPage True
  'V02.00.00
  WarnOnCanvasUnsync
End Sub

Private Sub cmdDecSize_Click()
  On Error Resume Next
  If Val(Me.txtFontSize) > 1 Then
    Me.txtFontSize = Val(Me.txtFontSize) - 1
    RecreateConsole
    RepositionConsouls
    DisplayPage True
  End If
End Sub

Private Sub cmdIncSize_Click()
  On Error Resume Next
  Me.txtFontSize = Val(Me.txtFontSize) + 1
  miFontSize = Val(Me.txtFontSize)
  RecreateConsole
  RepositionConsouls
  DisplayPage True
End Sub

Private Sub InitForm()
  On Error Resume Next
  
  mlComboBackColor = Me.cboFontName.BackColor
  Me.cboFontName.RowSource = GetFontsComboSource(FontGetFamilyFontFilter())
  msFontName = FontGetSelectedFont()
  Me.cboFontName = msFontName
  
  RecreateConsole
  FindChar AscW("A")
  DisplayPage True
End Sub

Public Sub RepositionConsouls()
  On Error Resume Next
  
  'Adjust to full form client area of the detail section of the form.
  Dim iHalfWidth  As Integer
  Dim iWidth      As Integer
  Dim iHeight     As Integer
  Dim iStHeight   As Integer
  Dim iHdHeight   As Integer
  
  iHdHeight = TwipsToPixelsY(Me.Section(1).Height)  'Height of the page header in pixels
  iStHeight = TwipsToPixelsY(Me.Section(2).Height)  'Height of the page footer in pixels
  
  iWidth = mrcClient.Right - mrcClient.left
  iHeight = mrcClient.Bottom - mrcClient.Top - iHdHeight - iStHeight
  
  moConCharset.MoveWindow 0, iHdHeight, iWidth, iHeight
End Sub

' Managing console windows
Private Function CreateConsouls() As Boolean
  Dim hwndParent As LongPtr
  
  On Error GoTo CreateConsouls_Err
  
  hwndParent = Me.hWnd
  
  'We repeatedly call this function to create/destroy the console windows
  If Not moConCharset Is Nothing Then
    ConsoulEventDispatcher.UnregisterEventSink moConCharset.hWnd
    Set moConCharset = Nothing
  End If
  
  'The console window displaying the character set
  Set moConCharset = New CConsoul
  With moConCharset
    .FontName = msFontName
    .FontSize = miFontSize
    .MaxCapacity = 20
    .BackColor = Me.Section(1).BackColor
    .ForeColor = Me.cboFontName.ForeColor
    '.RefreshOnAutoRedraw = True
  End With
  If Not moConCharset.Attach( _
      hwndParent, 0, 0, 0, 0, _
      AddressOf MSupport.OnConsoulMouseButton, _
      piCreateAttributes:=LW_TRACK_ZONES Or LW_RENDERMODEBYLINE Or LW_BKCOLSPILL _
    ) Then
    MsgBox "Failed to create charset window", vbCritical
    GoTo CreateConsouls_Exit
  End If
  'moConCharset.SetWmPaintCallback WMPAINTCBK_BEFORE + WMPAINTCBK_AFTER, AddressOf MSupport.OnConsoulWmPaint
  'moConCharset.SetWmPaintCallback WMPAINTCBK_BEFORE, AddressOf MSupport.OnConsoulWmPaint
  
  UpdateStatusBar
  
  ConsoulEventDispatcher.RegisterEventSink moConCharset.hWnd, Me, eCsMouseEvent
  'ConsoulEventDispatcher.RegisterEventSink moConCharset.hWnd, Me, eCsWmPaint
  
  moConCharset.ShowWindow True
  
  mlCodesCt = moConCharset.GetUnicodeCharCodes(malCodes()) '1 based array returned
  
  'txtKeyTrap is a textbox hidden behind our console,
  'that we'll use to trap keys and other events for us.
  'We hide and shrink the control, but it will receive/lose focus for us too.
  txtKeyTrap.Width = 0
  txtKeyTrap.Height = 0
  txtKeyTrap.SetFocus 'which is sort of a setfocus to the console window
  
  CreateConsouls = True
  
CreateConsouls_Exit:
  Exit Function

CreateConsouls_Err:
  MsgBox "Failed to create consoul's output windows"
End Function

Private Sub cmdNextPage_Click()
  OnPageDown
End Sub

Private Sub cmdPrevPage_Click()
  OnPageUp
End Sub

Private Sub Form_Load()
  Me.TimerInterval = 200
End Sub

Private Sub Form_Open(Cancel As Integer)
  If Len(Me.OpenArgs) > 0 Then
    Set moDialog = GetDialogClass(Me.OpenArgs)
    msFontName = moDialog.FontName
    miFontSize = moDialog.FontSize
  End If
  Me.cboFontName = msFontName
  Me.txtFontSize = miFontSize
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  GetClientRect Me.hWnd, mrcClient
  RepositionConsouls
End Sub

Private Sub Form_Timer()
  If Not Me.Visible Then Exit Sub
  
  On Error Resume Next
  Static fDoneOnce As Boolean
  
  If Not fDoneOnce Then
    InitForm
    fDoneOnce = True
  End If
  
  GetClientRect Me.hWnd, mrcClient
  RepositionConsouls
  Me.TimerInterval = 0
  MessageManager.SubscribeMulti Me, Array( _
                                      MSGTOPIC_LOCKUI, MSGTOPIC_UNLOCKUI, _
                                      MSGTOPIC_CANUNLOAD, MSGTOPIC_CANVASRESIZED, _
                                      MSGTOPIC_UNLOADNOW, MSGTOPIC_FONTNAMECHANGED, _
                                      MSGTOPIC_FONTFAMLCHANGED, MSGTOPIC_FINDCHAR _
                                    )
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not moConCharset Is Nothing Then
    ConsoulEventDispatcher.UnregisterEventSink moConCharset.hWnd
  End If
  Set moConCharset = Nothing
  MessageManager.Unsubscribe Me.Name, ""
End Sub

Private Function IsDisplayAble(ByVal plCharCode As Long) As Boolean
  'IsDisplayAble = CBool(moConCharset.TextWidth(psChar) = moConCharset.CharWidth)
  'IsDisplayAble = CBool((AscW(psChar) > 13) And (AscW(psChar) <> 27))
  IsDisplayAble = CBool((plCharCode > 13&) And (plCharCode <> 27&))
End Function

Private Function SegmentedSpace(ByVal piSpaceCt As Integer, Optional ByVal piForcedWidth As Integer = 0) As String
  Dim i As Integer
  Dim sSpace As String
  Dim iSpaceWidth As Integer
  
  iSpaceWidth = moConCharset.TextWidth(" ")
  For i = 1 To piSpaceCt
    If piForcedWidth = 0 Then
      sSpace = sSpace & " " & VT_NOOP()
    Else
      'sSpace = VTX_NEXT_WIDTH(piForcedWidth) & " "
      sSpace = sSpace & " " & VTX_ADVANCE_REL(piForcedWidth - iSpaceWidth)
    End If
  Next i
  SegmentedSpace = sSpace
End Function

Private Sub SelectChar(ByVal pfSelected As Boolean)
  Dim iConRow     As Integer
  Dim lCharCode   As Long
  Dim asRow(1 To CELL_SIZE) As String
  Dim i           As Integer
  Dim sChar       As String
  Dim sBegColor   As String
  Dim sEndColor   As String
  Dim fRed        As Boolean
  Dim iCharWidth As Integer
  
  If (miSelRow = 0) Or (miSelCol = 0) Then
    miSelRow = 1
    miSelCol = 1
  End If
  
  iConRow = (miSelRow - 1) * CELL_SIZE + 1
  'Check that we land on a valid index as we can click "outside" a char cell
  If (mlStart + (miSelRow - 1) * Cols + (miSelCol - 1)) > mlCodesCt Then Exit Sub
  
  lCharCode = CharCodeOf(miSelRow, miSelCol) ' malCodes(mlStart + (miSelRow - 1) * COLS + (miSelCol - 1))
  sChar = ChrW$(lCharCode)
  If (Not IsDisplayAble(lCharCode)) Then
    fRed = True
    sChar = "?"
    If pfSelected Then
      sBegColor = VT_FCOLOR(vbRed) & VT_INV_ON()
      sEndColor = VT_INV_OFF() & VT_FCOLOR(vbHighlight) & VT_BCOLOR(moConCharset.BackColor)
    Else
      sBegColor = VT_FCOLOR(vbRed)
      sEndColor = VT_FCOLOR(vbHighlight) & VT_BCOLOR(moConCharset.BackColor)
    End If
  Else
    If pfSelected Then
      sBegColor = VTX_SPILL(1) & VT_FCOLOR(vbHighlight) & VT_INV_ON()
'      sEndColor = VTX_SPILL(0) & VT_INV_OFF() & VT_RESET()
'      sBegColor = VT_FCOLOR(vbHighlight) & VT_INV_ON()
      sEndColor = VT_INV_OFF()
    Else
      sBegColor = VTX_SPILL(1) & VT_FCOLOR(vbHighlight)
'      sEndColor = VTX_SPILL(0) & VT_RESET()
'      sBegColor = VT_FCOLOR(vbHighlight)
      sEndColor = VT_RESET()
    End If
  End If
  iCharWidth = moConCharset.TextWidth(sChar)
  asRow(1) = sBegColor & _
             SegmentedSpace(CELL_SIZE, miFixedCharWidth) & _
             sEndColor
  asRow(2) = sBegColor & _
             SegmentedSpace(1, miFixedCharWidth) & _
             sChar & _
             VTX_ADVANCE_REL(miFixedCharWidth - iCharWidth) & _
             SegmentedSpace(1, miFixedCharWidth) & _
             sEndColor
  asRow(3) = sBegColor & _
             SegmentedSpace(CELL_SIZE, miFixedCharWidth) & _
             sEndColor
  miSelZone = miSelRow * 256 + miSelCol '16 bits word
  For i = 1 To CELL_SIZE
    moConCharset.ReplaceZone iConRow + i - 1, miSelZone, asRow(i)
    moConCharset.RedrawLine iConRow + i - 1
  Next i
  
  'adjust vertical view, make sure selection is visible
  i = (miSelRow - 1) * CELL_SIZE + 1
  If i > (moConCharset.TopLine + moConCharset.MaxVisibleRows - 2) Then 'not -1, take "empty" line on top of cell
    moConCharset.ScrollPageDown
  ElseIf i < moConCharset.TopLine Then
    If (i >= 1) And (i <= CELL_SIZE) Then
      moConCharset.ScrollTop
    Else
      moConCharset.ScrollPageUp
    End If
  End If
  
  moConCharset.RefreshWindow
  UpdateStatusBar
End Sub

Private Function MakeCellLine1(ByVal piZone As Integer, ByVal piFixedWidth As Integer) As String
  MakeCellLine1 = _
             VTX_ZONE_BEGIN(piZone) & _
             VT_BCOLOR(moConCharset.BackColor) & _
             SegmentedSpace(CELL_SIZE, piFixedWidth) & _
             VTX_ZONE_END(piZone) '& VT_RESET() 'top of character cell
End Function

Private Function MakeCellLine2(ByVal piZone As Integer, ByVal psChar As String, ByVal piFixedWidth As Integer) As String
  Dim iCharWidth As Integer
  
  iCharWidth = moConCharset.TextWidth(psChar)
  MakeCellLine2 = _
                 VTX_ZONE_BEGIN(piZone) & _
                 VT_BCOLOR(moConCharset.BackColor) & _
                 SegmentedSpace(1, piFixedWidth) & _
                 psChar & _
                 VTX_ADVANCE_REL(piFixedWidth - iCharWidth) & _
                 SegmentedSpace(1, piFixedWidth) & VTX_ZONE_END(piZone)  '& VT_RESET()

End Function

Private Function MakeCellLine3(ByVal piZone As Integer, ByVal piFixedWidth As Integer) As String
  MakeCellLine3 = _
                 VTX_ZONE_BEGIN(piZone) & _
                 VT_BCOLOR(moConCharset.BackColor) & _
                 SegmentedSpace(CELL_SIZE, piFixedWidth) & _
                 VTX_ZONE_END(piZone) '& VT_RESET() 'bottom of character cell
End Function

Public Sub DisplayPage(ByVal pfSelectChar As Boolean)
  On Error Resume Next
  
  'We display a grid of 10 columns per 6 rows
  'Each row takes 3 console lines
  Dim i     As Long
  
  'moConCharset.AutoRedraw = False
  moConCharset.Clear
  'moConCharset.ForeColor = vbWhite
  
  'we can go from character 0 to character 65535
  'Let's start at "A"
  
  Dim lCharCode   As Long
  Dim lRow        As Long
  Dim lCol        As Long
  Dim sChar       As String
  Dim iZone       As Integer
  Dim asRow(1& To Rows * CELL_SIZE) As String
  Dim lRowCt      As Long
  Dim iCharWidth  As Integer
  
  On Error GoTo OutOfLoops
  
  miFixedCharWidth = 0
  lCharCode = mlStart
  For lRow = 1& To Rows
    For lCol = 1& To Cols
      sChar = ChrW$(malCodes(lCharCode))
      If (Not IsDisplayAble(malCodes(lCharCode))) Then
        sChar = "?" ' lCharCode & ""
      End If
      iCharWidth = moConCharset.TextWidth(sChar)
      If iCharWidth > miFixedCharWidth Then
        miFixedCharWidth = iCharWidth
      End If
      lCharCode = lCharCode + 1&
      If lCharCode > mlCodesCt Then
        lRowCt = lRowCt + 1
        GoTo OutOfLoops2
      End If
    Next lCol
  Next lRow
OutOfLoops2:
  If miFixedCharWidth = 0 Then
    miFixedCharWidth = moConCharset.CharWidth
  End If
  
  lCharCode = mlStart
  For lRow = 1& To Rows
    For lCol = 1& To Cols
      i = (lRow - 1) * CELL_SIZE + 1
      sChar = ChrW$(malCodes(lCharCode))
      If (Not IsDisplayAble(malCodes(lCharCode))) Then
        sChar = "?" ' lCharCode & ""
      End If
      iZone = lRow * 256 + lCol '16 bits word
      
      asRow(i) = asRow(i) & MakeCellLine1(iZone, miFixedCharWidth)
      asRow(i + 1) = asRow(i + 1) & MakeCellLine2(iZone, sChar, miFixedCharWidth)
      asRow(i + 2) = asRow(i + 2) & MakeCellLine3(iZone, miFixedCharWidth)
      
      lCharCode = lCharCode + 1&
      If lCharCode > mlCodesCt Then
        lRowCt = lRowCt + 1
        GoTo OutOfLoops
      End If
    Next lCol
    lRowCt = lRowCt + 1
  Next lRow
  
  GoTo skipnextlabel
OutOfLoops:
  
skipnextlabel:

  'moConCharset.AutoRedraw = True
  
  For lRow = 1& To lRowCt * CELL_SIZE
    If lRow > UBound(asRow) Then Exit For
    moConCharset.OutputLn asRow(lRow) & VT_BCOLOR(moConCharset.BackColor) & " " & VT_RESET()
  Next lRow
  
  If pfSelectChar Then
    SelectChar True
  Else
    UpdatePageIndicator 'V02.00.00
  End If
End Sub

Private Function CharIndexOf(ByVal piRow As Integer, ByVal piCol As Integer) As Long
  Dim lIndex As Long
  lIndex = mlStart + (piRow - 1) * Cols + piCol - 1
  If (lIndex <= 0) Or (lIndex > mlCodesCt) Then Exit Function
  CharIndexOf = lIndex
End Function

Private Function CharCodeOf(ByVal piRow As Integer, ByVal piCol As Integer) As Long
  Dim lIndex As Long
  lIndex = mlStart + (piRow - 1) * Cols + piCol - 1
  If (lIndex <= 0) Or (lIndex > mlCodesCt) Then Exit Function
  CharCodeOf = malCodes(lIndex)
End Function

Private Sub UpdatePageIndicator()
  Dim iPageNo     As Integer
  Dim iPageCt     As Integer
  Dim iPageSize   As Integer
  On Error Resume Next
  iPageSize = Rows * Cols
  iPageNo = mlStart \ iPageSize + 1
  iPageCt = mlCodesCt \ iPageSize
  If mlCodesCt Mod iPageSize > 0 Then
    iPageCt = iPageCt + 1
  End If
  Me.lblPagePos.Caption = iPageNo & "/" & iPageCt
End Sub

Private Sub UpdateStatusBar()
  'Page XX/XX | Code: ABCD
  Dim sText       As String
  Dim sHexCode    As String
  
  sHexCode = Hex$(CharCodeOf(miSelRow, miSelCol))
  If Len(sHexCode) < 4 Then
    sHexCode = String$(4 - Len(sHexCode), "0") & sHexCode
  End If
  
  UpdatePageIndicator
  If mbDispCharCodeToggle = 0 Then
    Me.lblCharCode.Caption = "Hex: " & sHexCode
  Else
    Me.lblCharCode.Caption = "Dec: " & CharCodeOf(miSelRow, miSelCol)
  End If
End Sub

Public Sub RecreateConsole()
  'mlStart = 1
  miFontSize = Val(Me.txtFontSize)
  CreateConsouls
  RepositionConsouls
  mlStart = 1
  DisplayPage True
  UpdateStatusBar
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
  
  If phWnd = moConCharset.hWnd Then
    If piZoneID > 0 Then
      If (piEvtCode = eWmMouseButton.WM_LBUTTONUP) Then
        If piRow <= Rows * CELL_SIZE Then
          
          If miSelZone Then SelectChar False
          miSelZone = piZoneID
          miSelCol = piZoneID And &HFF
          miSelRow = (piZoneID - miSelCol) \ 256
          SelectChar True
          
          'Ctrl+click test with win32 api, take action on click after selection
          'Action is triggered by trapped [Enter] key also.
          If (GetKeyState(VK_CONTROL) < 0) Then
            If Not moDialog Is Nothing Then
              lCharCode = CharCodeOf((piZoneID - miSelCol) \ 256, piZoneID And &HFF)
              moDialog.OnCharacterSelected lCharCode
            End If
          End If
        End If
      'Double-click or CTRL+Left click
      ElseIf (piEvtCode = eWmMouseButton.WM_LBUTTONDBLCLK) Then
        lCharCode = CharCodeOf((piZoneID - miSelCol) \ 256, piZoneID And &HFF)
        If lCharCode <> 0 Then
          miSelZone = piZoneID
          miSelCol = piZoneID And &HFF
          miSelRow = (piZoneID - miSelCol) \ 256
          If Not moDialog Is Nothing Then
            If IsDisplayAble(lCharCode) Then
              moDialog.OnCharacterSelected lCharCode
            End If
          End If
        End If
      End If
    End If
    txtKeyTrap.SetFocus
  End If
End Function

Private Sub OnKeyReturn(Optional ByVal Shift As Integer = 0)
  Dim lCharCode As Long
  
  If miSelZone <> 0 Then
    If Not moDialog Is Nothing Then
      lCharCode = CharCodeOf((miSelZone - miSelCol) \ 256, miSelZone And &HFF)
      If IsDisplayAble(lCharCode) Then
        moDialog.OnCharacterSelected lCharCode
      End If
    End If
  End If
End Sub

Private Sub lblCharCode_Click()
  mbDispCharCodeToggle = mbDispCharCodeToggle + 1
  If mbDispCharCodeToggle > 1 Then mbDispCharCodeToggle = 0
  UpdateStatusBar
End Sub

Private Sub FindChar(ByVal plCharCode As Long)
  Dim i           As Long
  Dim iFound      As Long
  Dim iPageSize   As Integer
  Dim iPage       As Integer
  Dim iOffset     As Integer
  
  For i = 1& To mlCodesCt
    If malCodes(i) = plCharCode Then
      iFound = i
      Exit For
    End If
  Next i
  
  If iFound = 0 Then Exit Sub
  
  SelectChar False
  iPageSize = Rows * Cols
  mlStart = (iFound \ iPageSize) * iPageSize + 1
  iOffset = iFound - mlStart
  miSelRow = iOffset \ Cols + 1
  miSelCol = iOffset - ((miSelRow - 1) * Cols) + 1
  DisplayPage True
  UpdateStatusBar
End Sub

Private Sub lblPagePos_Click()
  Dim iPageNo     As Integer
  Dim iPageCt     As Integer
  Dim iPageSize   As Integer
  Dim iJumpTo     As Integer
  Dim sMsg        As String
  
  On Error Resume Next
  iPageSize = Rows * Cols
  iPageNo = mlStart \ iPageSize + 1
  iPageCt = mlCodesCt \ iPageSize
  If mlCodesCt Mod iPageSize > 0 Then
    iPageCt = iPageCt + 1
  End If
  
  sMsg = "You are on page #" & iPageNo & " on " & iPageCt
  sMsg = sMsg & vbCrLf & vbCrLf & "Enter page number to jump to, between 1 and " & iPageCt & ":"
  iJumpTo = IntChooseBox(sMsg, "Go To Page number", "1", iPageCt)
  If iJumpTo > 0 Then
    mlStart = (iJumpTo - 1) * (Rows * Cols) + 1
    DisplayPage True
    UpdatePageIndicator
  End If
End Sub

Private Sub txtKeyTrap_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  Select Case KeyCode
  Case vbKeyPageDown
    OnPageDown
  Case vbKeyPageUp
    OnPageUp
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
  Case vbKeyReturn
    OnKeyReturn Shift
  End Select
  KeyCode = 0
End Sub

Public Sub WarnOnCanvasUnsync()
  On Error Resume Next
  If msFontName <> FontGetSelectedFont() Then
    Me.cboFontName.BackColor = GetAlertBackgroundColor()
  Else
    Me.cboFontName.BackColor = mlComboBackColor
  End If
End Sub

Private Property Get IMessageReceiver_ClientID() As String
  IMessageReceiver_ClientID = Me.Name
End Property

Private Function IMessageReceiver_OnMessageReceived(ByVal psSenderID As String, ByVal psTopic As String, pvData As Variant) As Long
  Dim fOK       As Boolean
  Dim rowParams As CRow
  
  On Error GoTo OnMessageReceived_Err
  
  Select Case psTopic
  Case MSGTOPIC_LOCKUI
    fOK = MForms.FormSetAllowEdits(Me, False, "", "", 0)
  Case MSGTOPIC_UNLOCKUI
    fOK = MForms.FormSetAllowEdits(Me, True, "", "", 0)
  Case MSGTOPIC_CANUNLOAD
    IMessageReceiver_OnMessageReceived = 0& 'can always unload
  Case MSGTOPIC_CANVASRESIZED
    'reposition window
    Set rowParams = pvData
    rowParams("WindowLeft") = rowParams("WindowLeft") + rowParams("WindowWidth")
    Me.Move rowParams("WindowLeft"), rowParams("WindowTop")
    DoEvents  'Absolutely necessary to get the window size after the event
    rowParams("WindowTop") = Me.WindowTop + Me.WindowHeight
    MessageManager.Broadcast Me.Name, MSGTOPIC_CHARMAPMOVED, rowParams, GetPaletteFormName()
  Case MSGTOPIC_UNLOADNOW
    DoCmd.Close acForm, Me.Name
  Case MSGTOPIC_FONTNAMECHANGED
    Me.cboFontName = pvData
    cboFontName_AfterUpdate
    WarnOnCanvasUnsync
  Case MSGTOPIC_FONTFAMLCHANGED
    Me.cboFontName.RowSource = GetFontsComboSource(FontGetFamilyFontFilter())
    WarnOnCanvasUnsync
  Case MSGTOPIC_FINDCHAR
    Dim lCharCode As Long
    lCharCode = Nz(pvData, 0&)
    If lCharCode <> 0& Then
      FindChar lCharCode
    End If
  End Select

OnMessageReceived_Err:
End Function
