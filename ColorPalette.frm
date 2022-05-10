VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ColorPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements ICsMouseEventSink
Implements IMessageReceiver

'The dialog class associated with this form
Private moDialog      As CPaletteDialog

'The console displaying the palette
Private moConPalette  As CConsoul
Private Const MAX_COLS As Integer = 16

Private Const SUBDIR_PALETTES As String = "palettes"

'Clicked / selected color
Private miSelColorIndex   As Long
Private miSelColDisp      As Long

Private mrcClient   As RECT 'compute on Form_Resize and used in RepositionConsouls

'V02.00.00
'When this method is invoked, even from another form (mainly the Canvas), it
'will prompt to save changes (if palette is dirty).
'So the canvas can abort unloading also.
Public Function CanUnload() As Boolean
  Dim sMsg      As String
  Dim iRet      As VbMsgBoxResult
  
  If Not moDialog.Palette.Dirty Then
    CanUnload = True
  Else
    sMsg = "Save changes made to this palette ?"
    iRet = MsgBox(sMsg, vbYesNo + vbDefaultButton1 + vbQuestion)
    If iRet = vbYes Then
      If SaveToFile() Then
        CanUnload = True
      End If
    Else
      CanUnload = True
    End If
  End If
End Function

Private Sub UpdateDialogTitle()
  On Error Resume Next
  Dim sCaption    As String
  Dim lPixelWidth As Long
  Dim rcClient    As RECT
  
  sCaption = "Palette ("
  If Len(moDialog.Filename) > 0 Then
    'max 2/3 of dialog client width
    Call moConPalette.GetClientRect(rcClient)
    lPixelWidth = rcClient.Right - rcClient.left
    lPixelWidth = (lPixelWidth * 2) / 3
    'Debug.Print "lPixelWidth="; lPixelWidth
    sCaption = sCaption & CompactPath(Me.hWnd, moDialog.Filename, lPixelWidth) & ")"
  Else
    sCaption = sCaption & "new)"
  End If
  If moDialog.Palette.Dirty Then sCaption = sCaption & "*"
  Me.Caption = sCaption
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
  
  moConPalette.MoveWindow 0, iHdHeight, iWidth, iHeight
End Sub

' Managing console windows
Private Function CreateConsouls() As Boolean
  Dim hwndParent As LongPtr
  
  On Error GoTo CreateConsouls_Err
  
  hwndParent = Me.hWnd
  
  'We repeatedly call this function to create/destroy the console windows
  If Not moConPalette Is Nothing Then
    ConsoulEventDispatcher.UnregisterEventSink moConPalette.hWnd
    Set moConPalette = Nothing
  End If
  
  'The console window displaying the palette
  Set moConPalette = New CConsoul
  With moConPalette
    .FontName = "Lucida console"
    .FontSize = 32
    .MaxCapacity = moDialog.Palette.Count \ MAX_COLS + 1
    .BackColor = Me.Section(1).BackColor
    .ForeColor = vbBlack
  End With
  If Not moConPalette.Attach(hwndParent, 0, 0, 0, 0, AddressOf MSupport.OnConsoulMouseButton, piCreateAttributes:=LW_RENDERMODEBYLINE) Then
    MsgBox "Failed to create console window", vbCritical
    GoTo CreateConsouls_Exit
  End If
  
  ConsoulEventDispatcher.RegisterEventSink moConPalette.hWnd, Me, eCsMouseEvent
  moConPalette.ShowWindow True
  
  DisplayPalette
  
  'txtKeyTrap is a textbox hidden behind our console,
  'that we'll use to trap keys and other events for us.
  'We hide and shrink the control, but it will receive/lose focus for us too.
  'txtKeyTrap.Width = 0
  'txtKeyTrap.Height = 0
  txtKeyTrap.SetFocus 'which is sort of a setfocus to the console window
  
  CreateConsouls = True
  
CreateConsouls_Exit:
  Exit Function

CreateConsouls_Err:
  MsgBox "Failed to create consoul's output windows"
End Function

Public Sub RecreateConsole()
  'mlStart = 1
  CreateConsouls
  RepositionConsouls
  UpdateStatusBar
End Sub

Private Sub InitForm()
  RecreateConsole
  DisplayPalette
  UpdateDialogTitle
End Sub

Private Sub cmdAddColor_Click()
  If moDialog.Palette.AddColor(Me.rectAddColor.BackColor) > 0 Then
    Me.rectAddColor.BackStyle = 1
    Me.rectAddColor.BackColor = Me.Section(1).BackColor
    Me.cmdAddColor.Enabled = False
    DisplayPalette
    UpdateDialogTitle
  End If
End Sub

Private Sub LoadFromFile(Optional ByVal pfMerge As Boolean = False)
  Dim iChoice     As Integer
  Dim fLoaded     As Boolean
  Dim sFilename   As String
  Dim iMaxRows    As Integer
  Dim iMaxCols    As Integer
  Dim sMsg        As String
  Dim sInitialDir As String
  
  On Error Resume Next
  
  AppIniFile.GetOption INIOPT_PALLASTLOADPATH, sInitialDir, ""
  sInitialDir = Trim$(sInitialDir)
  If Len(sInitialDir) = 0 Then
    sInitialDir = CombinePath _
        ( _
          CombinePath _
          ( _
            GetSpecialFolder(Application.hWndAccessApp, CSIDL_PERSONAL), APP_NAME _
          ), _
          SUBDIR_PALETTES _
        )
  End If
  If Not ExistDir(sInitialDir) Then
    If Not CreatePath(sInitialDir) Then
      sInitialDir = GetSpecialFolder(Application.hWndAccessApp, CSIDL_PERSONAL)
    End If
  End If

  With Application.FileDialog(msoFileDialogFilePicker)
    If Not pfMerge Then
      .Title = "Load palette from file"
    Else
      .Title = "Insert palette from file"
    End If
    .InitialFileName = NormalizePath(sInitialDir)
    .Filters.Clear
    .Filters.Add "Palette files", "*.pal"
    .FilterIndex = 1
    iChoice = .Show()
    If iChoice <> 0 Then
      sFilename = .SelectedItems(1)
      DoCmd.Hourglass True
      fLoaded = moDialog.LoadFromFile(sFilename, pfMerge)
      If fLoaded Then
        If Not pfMerge Then
          SelectColor 0
        End If
        AppIniFile.SetOption INIOPT_PALLASTLOADPATH, (StripFileName(sFilename))
        DisplayPalette
        UpdateStatusBar
        UpdateDialogTitle
        DoCmd.Hourglass False
      Else
        DoCmd.Hourglass False
        ShowUFError "Failed to load palette file [" & sFilename & "]", moDialog.LastErrDesc
      End If
    End If
  End With
End Sub

Private Function SaveToFile() As Boolean
  Static iDumbCounter As Integer
  
  On Error Resume Next
  If iDumbCounter = 0 Then iDumbCounter = 1
  
  Dim iChoice   As Integer
  Dim fSaved    As Boolean
  Dim sFilename As String
  Dim sInitialDir As String
  
  AppIniFile.GetOption INIOPT_PALLASTSAVEPATH, sInitialDir, ""
  sInitialDir = Trim$(sInitialDir)
  If Len(sInitialDir) = 0 Then
    sInitialDir = CombinePath _
        ( _
          CombinePath _
          ( _
            GetSpecialFolder(Application.hWndAccessApp, CSIDL_PERSONAL), APP_NAME _
          ), _
          SUBDIR_PALETTES _
        )
  End If
  If Not ExistDir(sInitialDir) Then
    If Not CreatePath(sInitialDir) Then
      sInitialDir = GetSpecialFolder(Application.hWndAccessApp, CSIDL_PERSONAL)
    End If
  End If

  With Application.FileDialog(msoFileDialogSaveAs)
    .Title = "Save palette"
    .InitialFileName = CombinePath(sInitialDir, "AsciiPaintPalette" & iDumbCounter & ".pal")
    .Filters.Clear
    .Filters.Add "Palette files", "*.pal"
    .FilterIndex = 1
    iChoice = .Show()
    If iChoice <> 0 Then
      sFilename = .SelectedItems(1)
      DoCmd.Hourglass True
      fSaved = moDialog.SaveToFile(sFilename)
      DoCmd.Hourglass False
      If fSaved Then
        AppIniFile.SetOption INIOPT_PALLASTSAVEPATH, (StripFileName(sFilename))
        UpdateDialogTitle
      Else
        ShowUFError "Failed to save palette to file [" & sFilename & "]", moDialog.LastErrDesc
      End If
    End If
  End With
  
  If fSaved Then
    iDumbCounter = iDumbCounter + 1
    SaveToFile = True
  End If
End Function

Private Sub cmdInsertFile_Click()
  On Error Resume Next
  If Not LockUI() Then Exit Sub
  LoadFromFile True
  UnlockUI
End Sub

Private Sub cmdLoadCanvasColors_Click()
  On Error Resume Next
  If Not LockUI() Then Exit Sub
  DoCmd.Hourglass True
  '/**/Transform to message
  Forms(GetCanvasFormName()).ColorsToPalette moDialog.Palette
  CreateConsouls
  RepositionConsouls
  DisplayPalette
  UpdateStatusBar
  UpdateDialogTitle
  DoCmd.Hourglass False
  UnlockUI
End Sub

Private Sub cmdLoadFile_Click()
  On Error Resume Next
  If Not LockUI() Then Exit Sub
  LoadFromFile False
  UnlockUI
End Sub

Private Sub cmdNewFile_Click()
  Dim iRet      As VbMsgBoxResult
  Dim sMsg      As String
  
  If Not LockUI() Then Exit Sub
  On Error Resume Next
  
  sMsg = "Add default colors to the new palette ?"
  iRet = MsgBox(sMsg, vbQuestion + vbYesNo + vbDefaultButton1)
  
  moDialog.Clear
  If iRet = vbYes Then
    moDialog.Palette.LoadQBColors
    SelectColor 1
  Else
    SelectColor 0
  End If
  DisplayPalette
  UpdateStatusBar
  UpdateDialogTitle
  
  UnlockUI
End Sub

Private Sub cmdSaveFile_Click()
  If Not LockUI() Then Exit Sub
  On Error Resume Next
  SaveToFile
  UpdateDialogTitle
  UnlockUI
End Sub

Private Sub cmdSortPalette_Click()
  On Error Resume Next
  If Not LockUI() Then Exit Sub
  DoCmd.Hourglass True
  moDialog.SortPalette
  CreateConsouls
  RepositionConsouls
  DisplayPalette
  UpdateStatusBar
  UpdateDialogTitle
  DoCmd.Hourglass False
  UnlockUI
End Sub

Private Sub Form_Load()
  Me.TimerInterval = 200
End Sub

Private Sub Form_Open(Cancel As Integer)
  If Len(Me.OpenArgs) > 0 Then
    Set moDialog = GetDialogClass(Me.OpenArgs)
  Else
    Set moDialog = New CPaletteDialog
  End If
  moDialog.Palette.LoadQBColors
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  GetClientRect Me.hWnd, mrcClient
  RepositionConsouls
  UpdateDialogTitle
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
                                      MSGTOPIC_CANUNLOAD, MSGTOPIC_CHARMAPMOVED, _
                                      MSGTOPIC_GETSELCOLOR, MSGTOPIC_ADDNSELCOLOR, _
                                      MSGTOPIC_UNLOADNOW _
                                    )
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not moConPalette Is Nothing Then
    ConsoulEventDispatcher.UnregisterEventSink moConPalette.hWnd
  End If
  MessageManager.Unsubscribe Me.Name, ""
  Set moConPalette = Nothing
  Set moDialog = Nothing
End Sub

Private Function RenderLine(ByVal piLine As Integer) As String
  Dim sRender     As String
  Dim iStartIndex As Long
  Dim iIndex      As Long
  Dim i           As Long
  
  iStartIndex = CLng((piLine - 1)) * MAX_COLS + 1
  For i = 1 To MAX_COLS
    iIndex = (iStartIndex + i - 1)
    If iIndex > moDialog.Palette.Count Then Exit For
    With moDialog.Palette
      If iIndex <> miSelColorIndex Then
        sRender = sRender & VT_BCOLOR(.Color(iIndex)) & " " & VT_RESET()
      Else
        'sRender = sRender & VT_FCOLOR(InverseColor(.Color(iIndex))) & VT_BCOLOR(.Color(iIndex)) & ChrW$(8729) & VT_RESET()
        sRender = sRender & VT_FCOLOR(moConPalette.BackColor) & VT_BCOLOR(.Color(iIndex)) & ChrW$(8729) & VT_RESET()
      End If
    End With
  Next i
  RenderLine = sRender
End Function

Public Sub DisplayPalette()
  Dim i         As Long
  Dim sRender   As String
  Dim iLineCt   As Integer
  
  On Error GoTo DisplayPalette_Err
  
  moConPalette.Clear
  With moDialog.Palette
    iLineCt = .Count \ MAX_COLS
    If (.Count - (CLng(iLineCt) * MAX_COLS)) > 0 Then
      iLineCt = iLineCt + 1
    End If
    For i = 1 To iLineCt
      sRender = RenderLine(i)
      moConPalette.OutputLn sRender
    Next i
  End With

DisplayPalette_Exit:
  Exit Sub

DisplayPalette_Err:
  '/**/
End Sub

Public Sub RefreshView()
  'refresh visible lines
  Dim iRow      As Integer
  For iRow = moConPalette.TopLine To (moConPalette.TopLine + moConPalette.MaxVisibleRows - 2)
    moConPalette.RedrawLine iRow
  Next iRow
End Sub

Public Sub UpdateStatusBar()
  Dim lColor    As Long
  If miSelColorIndex > 0 Then
    lColor = moDialog.Palette.Color(miSelColorIndex)
    Me.rectSelCol.BackStyle = 1
    Me.rectSelCol.BackColor = lColor
    Me.lblSelCol.Caption = GetColorDispString(lColor, miSelColDisp)
    Me.lblColorIndex.Caption = miSelColorIndex & "/" & moDialog.Palette.Count
  Else
    Me.rectSelCol.BackStyle = 0
    Me.rectSelCol.BackColor = Me.Section(1).BackColor
    Me.lblSelCol.Caption = "(no selection)"
    Me.lblColorIndex.Caption = ""
  End If
End Sub

'Will be called by the canvas form
Public Sub SelectColor(ByVal plColorIndex As Long)
  Dim iRow      As Integer
  Dim iCol      As Integer
  
  'This is a public member, mostly used privately, but we'll do bounds checking
  If (plColorIndex = 0) Then
    If miSelColorIndex > 0 Then
      iRow = ((miSelColorIndex - 1) \ MAX_COLS) + 1
      miSelColorIndex = plColorIndex
      moConPalette.SetLine iRow, RenderLine(iRow)
    End If
  End If
  
  If (plColorIndex > moDialog.Palette.Count) Then Exit Sub
  
  If miSelColorIndex <> 0 Then
    'unselect
    If miSelColorIndex > 0 Then
      iRow = ((miSelColorIndex - 1) \ MAX_COLS) + 1
      miSelColorIndex = plColorIndex
      moConPalette.SetLine iRow, RenderLine(iRow)
    End If
  Else
    miSelColorIndex = plColorIndex
  End If
  
  iRow = ((miSelColorIndex - 1) \ MAX_COLS) + 1
  moConPalette.SetLine iRow, RenderLine(iRow)
  If Not moConPalette.IsRowVisible(iRow) Then
    moConPalette.ScrollTo iRow
  End If
  
  moDialog.SelectedColorIndex = miSelColorIndex
  RefreshView
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
  Dim lIndex    As Long
  
  If phWnd = moConPalette.hWnd Then
    If piEvtCode = eWmMouseButton.WM_LBUTTONUP Then
      
      If piCol <= MAX_COLS Then
        lIndex = CLng((piRow - 1)) * MAX_COLS + piCol
        If (lIndex > 0) And (lIndex <= moDialog.Palette.Count) Then
          SelectColor lIndex
          UpdateStatusBar
        End If
      End If
      
    ElseIf piEvtCode = eWmMouseButton.WM_LBUTTONDBLCLK Then
      If miSelColorIndex > 0 Then
        '/**/transform into message
        moDialog.OnColorSelected moDialog.Palette.Color(miSelColorIndex)
      End If
    End If
    txtKeyTrap.SetFocus
  End If
End Function

Private Sub lblSelCol_Click()
  miSelColDisp = miSelColDisp + 1
  If miSelColDisp > 2 Then
    miSelColDisp = 0
  End If
  UpdateStatusBar
End Sub

Private Sub rectAddColor_Click()
  Dim lColor    As Long
  On Error Resume Next
  If PickColor(Me.hWnd, lColor) Then
    Me.rectAddColor.BackStyle = 1
    Me.rectAddColor.BackColor = lColor
    Me.cmdAddColor.Enabled = True
  End If
End Sub

Private Sub OnKeyDelete()
  On Error Resume Next
  If miSelColorIndex > 0 Then
    DoCmd.Hourglass True
    moDialog.Palette.DeleteColor miSelColorIndex
    If miSelColorIndex > moDialog.Palette.Count Then
      SelectColor miSelColorIndex
    End If
    DisplayPalette
    If moDialog.Palette.Count > 0 Then
      If miSelColorIndex <= moDialog.Palette.Count Then
        SelectColor miSelColorIndex 'moDialog.Palette.Count
      End If
    End If
    UpdateDialogTitle
    DoCmd.Hourglass False
  End If
End Sub

Private Sub txtKeyTrap_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  If IsUILocked() Then
    Exit Sub
  End If
  Select Case KeyCode
  Case vbKeyDelete
    If Not LockUI() Then
      KeyCode = 0
      Exit Sub
    End If
    OnKeyDelete
    UnlockUI
    KeyCode = 0
  End Select
End Sub

Private Property Get IMessageReceiver_ClientID() As String
  IMessageReceiver_ClientID = Me.Name
End Property

Private Function IMessageReceiver_OnMessageReceived(ByVal psSenderID As String, ByVal psTopic As String, pvData As Variant) As Long
  Dim fOK       As Boolean
  Dim rowParams As CRow
  Dim vColor    As Variant
  
  Select Case psTopic
  Case MSGTOPIC_LOCKUI
    fOK = MForms.FormSetAllowEdits(Me, False, "", "", 0)
  Case MSGTOPIC_UNLOCKUI
    fOK = MForms.FormSetAllowEdits(Me, True, "", "", 0)
  Case MSGTOPIC_CANUNLOAD
    fOK = CanUnload()
    If fOK Then
      IMessageReceiver_OnMessageReceived = 0&
    Else
      IMessageReceiver_OnMessageReceived = 1& 'breaks the broadcast chain
    End If
  Case MSGTOPIC_CHARMAPMOVED
    'reposition window
    Set rowParams = pvData
    Me.Move rowParams("WindowLeft"), rowParams("WindowTop")
  Case MSGTOPIC_GETSELCOLOR
    Set rowParams = pvData
    vColor = Null
    If moDialog.SelectedColorIndex > 0 Then
      vColor = moDialog.Palette.Color(moDialog.SelectedColorIndex)
      rowParams("color") = vColor
    End If
    IMessageReceiver_OnMessageReceived = 1& 'breaks the broadcast chain
  Case MSGTOPIC_ADDNSELCOLOR
    Dim lColorIndex As Long
    Set rowParams = pvData
    vColor = rowParams("color")
    If Not IsNull(vColor) Then
      lColorIndex = moDialog.Palette.ColorIndex(vColor)
      If lColorIndex > 0& Then
        SelectColor lColorIndex
      Else
        moDialog.Palette.AddColor vColor
        lColorIndex = moDialog.Palette.ColorIndex(vColor)
        If lColorIndex > 0& Then
          RecreateConsole
          DisplayPalette
          SelectColor lColorIndex
        End If
      End If
      IMessageReceiver_OnMessageReceived = 1& 'breaks the broadcast chain
    End If
  Case MSGTOPIC_UNLOADNOW
    DoCmd.Close acForm, Me.Name
  End Select
End Function
