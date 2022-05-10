VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCTabStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'TabStrip properties
' - BackColor
' - ForeColor
' - MaxTabWidth
' - FixedTabWidth
' - TabCount
' - BaseZoneID
' - ActiveTab (Index)
' - StartTab
' Tab properties
' - Caption
' - Width
' - BackColor
' - ForeColor
' - CanClose (shows close button or not)

'01.08.2021|FFO|01.02.00|Rendering tab display as buttons or tab strip

Public Enum eCCTabStyle
  eTabStrip
  eTabButtons
End Enum
Private meTabStyle        As eCCTabStyle
Private msButtonSep       As String

Private mlstTabs          As CList
Private miMaxTabWidth     As Integer
Private mfFixedTabWidth   As Boolean
Private mvBackColor       As Variant
Private mvForeColor       As Variant
Private miActiveTab       As Integer
Private miBaseZoneID      As Integer  'ZoneID for tab click events
Private miStartTab        As Integer

Private mvActiveTabForeColor      As Variant
Private mvActiveTabBackColor      As Variant
Private mvCloseBoxForeColor       As Variant
Private mvCloseBoxBackColor       As Variant
Private mvCloseBoxHoverForeColor  As Variant
Private mvCloseBoxHoverBackColor  As Variant

Private moConsole         As CConsoul
Private miLine            As Integer

'Array of strings where the control is rendered.
'The control needs 4 console lines for tab strip, and 3 for buttons.
Private Const NUMLINES_TABSTRIP As Integer = 4
Private Const NUMLINES_BUTTONS  As Integer = 1
Private masRender(1 To NUMLINES_TABSTRIP) As String

'Handle click event by tracking down zone & row, then click on up if match
Private miMBdownZoneID    As Integer
Private miMBdownRow       As Integer
Private miMBdownCol       As Integer

Private Const ZONEID_CLOSEBOXES_OFFSET As Integer = 100
Private Const ZONEID_LEFTCHEVRON As Integer = 80
Private Const ZONEID_RIGHTCHEVRON As Integer = 81

Event TabClick(ByVal piTabIndex As Integer, ByVal piMouseFlags As Integer, ByVal piZoneID As Integer, ByVal psZoneTag As String)
Event TabClose(ByVal piTabIndex As Integer, ByRef pfCancel As Boolean)
Event TabRemoved(ByVal piTabIndex As Integer, ByVal psTabName As String)

Private Sub Class_Initialize()
  Set mlstTabs = New CList
  mlstTabs.ArrayDefine Array("Name", "Caption", "Width", "BackColor", "ForeColor", "CanClose"), _
                       Array(vbString, vbString, vbInteger, vbVariant, vbVariant, vbBoolean)
  miMaxTabWidth = 8
  mvForeColor = Null
  mvBackColor = Null
  
  mvActiveTabForeColor = Null
  mvActiveTabBackColor = Null
  mvCloseBoxForeColor = Null
  mvCloseBoxBackColor = Null
  mvCloseBoxHoverForeColor = Null
  mvCloseBoxHoverBackColor = Null
  
  meTabStyle = eTabStrip
  miStartTab = 1
  msButtonSep = " "
End Sub

Private Sub Class_Terminate()
  Set moConsole = Nothing
  Set mlstTabs = Nothing
End Sub

Public Property Get StartTab() As Integer
  StartTab = miStartTab
End Property

Public Property Let StartTab(ByVal piStartTab As Integer)
  If (piStartTab < 1) Or (StartTab > (mlstTabs.Count - 1)) Then Exit Property
  miStartTab = piStartTab
End Property

Public Property Get BaseZoneID() As Integer
  BaseZoneID = miBaseZoneID
End Property

Public Property Let BaseZoneID(ByVal piBaseZoneID As Integer)
  miBaseZoneID = piBaseZoneID
End Property

Public Property Get Line() As Integer
  Line = miLine
End Property

Public Property Get TabCount() As Integer
  TabCount = mlstTabs.Count
End Property

Public Function Attach(ByRef poConsole As CConsoul, Optional ByVal piLine As Integer = 0)
  Set moConsole = poConsole
  miLine = piLine
End Function

Public Sub Detach()
  Set moConsole = Nothing
  miLine = 0
End Sub

Public Property Get ActiveTab() As Integer
  ActiveTab = miActiveTab
End Property

Public Property Let ActiveTab(ByVal piActiveTab As Integer)
  If (piActiveTab < 1) Or (piActiveTab > mlstTabs.Count) Then Exit Property
  miActiveTab = piActiveTab
End Property

Public Property Get MaxTabWidth() As Integer
  MaxTabWidth = miMaxTabWidth
End Property

Public Property Let MaxTabWidth(ByVal piMaxTabWidth As Integer)
  miMaxTabWidth = piMaxTabWidth
End Property

Public Property Get FixedTabWidth() As Boolean
  FixedTabWidth = mfFixedTabWidth
End Property

Public Property Let FixedTabWidth(ByVal pfFixedTabWidth As Boolean)
  mfFixedTabWidth = pfFixedTabWidth
End Property

Public Property Get BackColor() As Variant
  If IsNull(mvBackColor) Then
    If Not moConsole Is Nothing Then
      BackColor = moConsole.BackColor
    Else
      BackColor = vbWhite
    End If
  Else
    BackColor = mvBackColor
  End If
End Property

Public Property Let BackColor(ByVal pvBackColor As Variant)
  mvBackColor = pvBackColor
End Property

Public Property Get ForeColor() As Variant
  If IsNull(mvForeColor) Then
    If Not moConsole Is Nothing Then
      ForeColor = moConsole.ForeColor
    Else
      ForeColor = vbBlack
    End If
  Else
    ForeColor = mvForeColor
  End If
End Property

Public Property Let ForeColor(ByVal pvForeColor As Variant)
  mvForeColor = pvForeColor
End Property

Public Property Get TabBackColor(ByVal pvTabIndex As Variant) As Variant
  If VarType(pvTabIndex) = vbString Then
    pvTabIndex = Me.TabIndex(pvTabIndex)
  End If
  If (pvTabIndex < 1) Or (pvTabIndex > mlstTabs.Count) Then Exit Property
  TabBackColor = mlstTabs("BackColor", pvTabIndex)
End Property

Public Property Let TabBackColor(ByVal pvTabIndex As Variant, ByVal pvBackColor As Variant)
  If VarType(pvTabIndex) = vbString Then
    pvTabIndex = Me.TabIndex(pvTabIndex)
  End If
  If (pvTabIndex < 1) Or (pvTabIndex > mlstTabs.Count) Then Exit Property
  mlstTabs("BackColor", pvTabIndex) = pvBackColor
End Property

Public Property Get TabForeColor(ByVal pvTabIndex As Variant) As Variant
  If VarType(pvTabIndex) = vbString Then
    pvTabIndex = Me.TabIndex(pvTabIndex)
  End If
  If (pvTabIndex < 1) Or (pvTabIndex > mlstTabs.Count) Then Exit Property
  TabForeColor = mlstTabs("ForeColor", pvTabIndex)
End Property

Public Property Let TabForeColor(ByVal pvTabIndex As Variant, ByVal pvForeColor As Variant)
  If VarType(pvTabIndex) = vbString Then
    pvTabIndex = Me.TabIndex(pvTabIndex)
  End If
  If (pvTabIndex < 1) Or (pvTabIndex > mlstTabs.Count) Then Exit Property
  mlstTabs("ForeColor", pvTabIndex) = pvForeColor
End Property

Public Property Get ActiveTabBackColor() As Variant
  ActiveTabBackColor = mvActiveTabBackColor
End Property

Public Property Let ActiveTabBackColor(ByVal pvBackColor As Variant)
  mvActiveTabBackColor = pvBackColor
End Property

Public Property Get ActiveTabForeColor() As Variant
  ActiveTabForeColor = mvActiveTabForeColor
End Property

Public Property Let ActiveTabForeColor(ByVal pvForeColor As Variant)
  mvActiveTabForeColor = pvForeColor
End Property

Public Property Get CloseBoxForeColor() As Variant
  If IsNull(mvCloseBoxForeColor) Then
    CloseBoxForeColor = RGB(140, 140, 140)
  Else
    CloseBoxForeColor = mvCloseBoxForeColor
  End If
End Property

Public Property Let CloseBoxForeColor(ByVal pvCloseBoxForeColor As Variant)
  mvCloseBoxForeColor = pvCloseBoxForeColor
End Property

Public Property Get CloseBoxBackColor() As Variant
  If IsNull(mvCloseBoxBackColor) Then
    If Not moConsole Is Nothing Then
      CloseBoxBackColor = moConsole.BackColor
    Else
      CloseBoxBackColor = vbWhite
    End If
  Else
    CloseBoxBackColor = mvCloseBoxBackColor
  End If
End Property

Public Property Let CloseBoxBackColor(ByVal pvCloseBoxBackColor As Variant)
  mvCloseBoxBackColor = pvCloseBoxBackColor
End Property

Public Property Get CloseBoxHoverForeColor() As Variant
  If IsNull(mvCloseBoxHoverForeColor) Then
    CloseBoxHoverForeColor = vbWhite
  Else
    CloseBoxHoverForeColor = mvCloseBoxHoverForeColor
  End If
End Property

Public Property Let CloseBoxHoverForeColor(ByVal pvCloseBoxHoverForeColor As Variant)
  mvCloseBoxHoverForeColor = pvCloseBoxHoverForeColor
End Property

Public Property Get CloseBoxHoverBackColor() As Variant
  If IsNull(mvCloseBoxHoverBackColor) Then
    CloseBoxHoverBackColor = vbRed
  Else
    CloseBoxHoverBackColor = mvCloseBoxHoverBackColor
  End If
End Property

Public Property Let CloseBoxHoverBackColor(ByVal pvCloseBoxHoverBackColor As Variant)
  mvCloseBoxHoverBackColor = pvCloseBoxHoverBackColor
End Property

Public Function AddTab( _
  ByVal psTabName As String, _
  ByVal psCaption As String, _
  Optional ByVal pfCanClose As Boolean = False, _
  Optional ByVal piInsertAfter As Integer = 0, _
  Optional ByVal piInsertBefore As Integer = 0) As Integer
  
  Dim iTabIndex   As Integer
  
  If Len(psTabName) = 0 Then
    Err.Raise 5, "AddTab", "You must supply a tab name."
    Exit Function
  End If
  
  iTabIndex = mlstTabs.Find("Name", psTabName)
  If iTabIndex <> 0 Then
    Err.Raise 457, "AddTab", "A tab named [" & psTabName & "] already exists."
    Exit Function
  End If
  
  Dim rowTab      As New CRow
  mlstTabs.DefineRow rowTab
  rowTab("Name") = psTabName
  rowTab("Caption") = psCaption
  rowTab("Width") = Min(Len(psCaption), miMaxTabWidth)
  rowTab("ForeColor") = Null
  rowTab("BackColor") = Null
  rowTab("CanClose") = pfCanClose
  If piInsertBefore > 0 Then
    AddTab = mlstTabs.AddRow(rowTab, plInsertBefore:=piInsertBefore)
  ElseIf piInsertAfter > 0 Then
    AddTab = mlstTabs.AddRow(rowTab, piInsertAfter)
  Else
    AddTab = mlstTabs.AddRow(rowTab)
  End If
End Function

Public Function TabIndex(ByVal psTabName As String) As Integer
  TabIndex = mlstTabs.Find("Name", psTabName)
End Function

Public Function TabName(ByVal piTabIndex As Integer) As String
  If (piTabIndex < 1) Or (piTabIndex > mlstTabs.Count) Then Exit Function
  TabName = mlstTabs("Name", piTabIndex)
End Function

Public Sub RemoveTab(ByVal pvTabIndex As Variant)
  On Error Resume Next
  Dim sTabName As String
  If VarType(pvTabIndex) = vbString Then
    pvTabIndex = Me.TabIndex(pvTabIndex)
  End If
  If (pvTabIndex < 1) Or (pvTabIndex > mlstTabs.Count) Then Exit Sub
  sTabName = mlstTabs("Name", pvTabIndex)
  mlstTabs.Remove pvTabIndex
  RaiseEvent TabRemoved(CInt(pvTabIndex), sTabName)
End Sub

Public Property Get TabCaption(ByVal pvTabIndex As Variant) As String
  If VarType(pvTabIndex) = vbString Then
    pvTabIndex = Me.TabIndex(pvTabIndex)
  End If
  If (pvTabIndex < 1) Or (pvTabIndex > mlstTabs.Count) Then Exit Property
  TabCaption = mlstTabs("Caption", pvTabIndex)
End Property

'Caller MUST re-render after changing one or more tab captions
Public Property Let TabCaption(ByVal pvTabIndex As Variant, ByVal psCaption As String)
  If VarType(pvTabIndex) = vbString Then
    pvTabIndex = Me.TabIndex(pvTabIndex)
  End If
  If (pvTabIndex < 1) Or (pvTabIndex > mlstTabs.Count) Then Exit Property
  mlstTabs("Caption", pvTabIndex) = psCaption
End Property

Private Sub ClearRender()
  Dim i As Integer
  For i = 1 To UBound(masRender)
    masRender(i) = ""
  Next i
End Sub

Private Function IsValidLineIndex(ByVal piLineIndex As Integer) As Boolean
  Dim iMaxLine As Integer
  If meTabStyle = eTabStrip Then
    iMaxLine = NUMLINES_TABSTRIP
  Else
    iMaxLine = NUMLINES_BUTTONS
  End If
  IsValidLineIndex = CBool((piLineIndex > 0) And (piLineIndex <= iMaxLine))
End Function

Public Property Get RenderedLine(ByVal piLineIndex As Integer) As String
  If Not IsValidLineIndex(piLineIndex) Then Exit Function
  RenderedLine = masRender(piLineIndex)
End Property

Public Property Get TabStyle() As eCCTabStyle
  TabStyle = meTabStyle
End Property

Public Property Let TabStyle(ByVal peStyle As eCCTabStyle)
  meTabStyle = peStyle
End Property

Public Property Get ButtonSep() As String
  ButtonSep = msButtonSep
End Property

Public Property Let ButtonSep(ByVal psButtonSep As String)
  msButtonSep = psButtonSep
End Property

Private Function GetTabInternalWidth(ByVal piTabIndex As Integer)
  If Not mfFixedTabWidth Then
    GetTabInternalWidth = Len(mlstTabs("Caption", piTabIndex))
    If mlstTabs("CanClose", piTabIndex) Then
      GetTabInternalWidth = GetTabInternalWidth + 2
    End If
  Else
    GetTabInternalWidth = miMaxTabWidth
  End If
End Function

'This gives the necessary minimal width occupied by the control.
'The expansion of the horizontal till the end of the console is not considered.
Public Function GetControlMinWidth() As Integer
  Dim i             As Integer
  Dim iTotalWidth   As Integer
  If mlstTabs.Count = 0 Then Exit Function
  For i = 1 To mlstTabs.Count
    iTotalWidth = iTotalWidth + GetTabInternalWidth(i) + 2 '+2=borders
  Next i
  GetControlMinWidth = iTotalWidth
End Function

Private Function RenderCloseButtonText(ByVal piTabIndex As Integer, ByVal pfHovering As Boolean) As String
  Dim iZoneID       As Long
  Dim sButton       As String
  
  iZoneID = miBaseZoneID + ZONEID_CLOSEBOXES_OFFSET + piTabIndex
  If Not pfHovering Then
    If Not IsNull(mvCloseBoxBackColor) Then
      sButton = VT_FCOLOR(Me.CloseBoxForeColor) & VT_BCOLOR(Me.CloseBoxBackColor) & _
                "x" & _
                VT_FCOLOR(Me.ForeColor) & VT_BCOLOR(Me.BackColor)
    Else
      If Not IsNull(Me.TabBackColor(piTabIndex)) Then
        sButton = VT_FCOLOR(Me.CloseBoxForeColor) & VT_BCOLOR(Me.TabBackColor(piTabIndex)) & _
                  "x" & _
                  VT_FCOLOR(Me.ForeColor) & VT_BCOLOR(Me.BackColor)
      Else
        sButton = VT_FCOLOR(Me.CloseBoxForeColor) & _
                  "x" & _
                  VT_FCOLOR(Me.ForeColor)
      End If
    End If
  Else
    sButton = VT_FCOLOR(Me.CloseBoxHoverForeColor) & VT_BCOLOR(Me.CloseBoxHoverBackColor) & _
              "x" & _
              VT_FCOLOR(Me.ForeColor) & VT_BCOLOR(Me.BackColor)
  End If
  RenderCloseButtonText = sButton
End Function

Private Function RenderCloseButton(ByVal piTabIndex As Integer, ByVal pfHovering As Boolean) As String
  Dim iZoneID       As Long
  Dim sButton       As String
  
  iZoneID = miBaseZoneID + ZONEID_CLOSEBOXES_OFFSET + piTabIndex
  If Not pfHovering Then
    sButton = VTX_ZONE_BEGIN(iZoneID, mlstTabs("Caption", piTabIndex)) & _
              RenderCloseButtonText(piTabIndex, False) & _
              VTX_ZONE_END(iZoneID)
  Else
    sButton = VTX_ZONE_BEGIN(iZoneID, mlstTabs("Caption", piTabIndex)) & _
              RenderCloseButtonText(piTabIndex, True) & _
              VTX_ZONE_END(iZoneID)
  End If
  RenderCloseButton = sButton
End Function

Private Function ColorizeTabText(ByVal piTabIndex As Integer, ByVal psText As String) As String
  Dim sColorTagsBegin   As String
  Dim sColorTagsEnd     As String
  
  Dim lForeColor        As String
  Dim lBackColor        As String
  
  If (piTabIndex = miActiveTab) Then
    If Not IsNull(Me.TabForeColor(piTabIndex)) Then
      lForeColor = Me.TabForeColor(piTabIndex)
    Else
      If Not IsNull(mvActiveTabForeColor) Then
        lForeColor = mvActiveTabForeColor
      Else
        lForeColor = Me.ForeColor
      End If
    End If
    If Not IsNull(Me.TabBackColor(piTabIndex)) Then
      lBackColor = Me.TabBackColor(piTabIndex)
    Else
      If Not IsNull(mvActiveTabBackColor) Then
        lBackColor = mvActiveTabBackColor
      Else
        lBackColor = Me.BackColor
      End If
    End If
  Else
    If Not IsNull(Me.TabForeColor(piTabIndex)) Then
      lForeColor = Me.TabForeColor(piTabIndex)
    Else
      lForeColor = Me.ForeColor
    End If
    If Not IsNull(Me.TabBackColor(piTabIndex)) Then
      lBackColor = Me.TabBackColor(piTabIndex)
    Else
      lBackColor = Me.BackColor
    End If
  End If
  
  sColorTagsBegin = VT_FCOLOR(lForeColor) & VT_BCOLOR(lBackColor)
  sColorTagsEnd = VT_FCOLOR(Me.ForeColor) & VT_BCOLOR(Me.BackColor)
  
  psText = sColorTagsBegin & psText & sColorTagsEnd
  
  ColorizeTabText = psText
End Function

Private Function GetTabText(ByVal piTabIndex As Integer) As String
  Dim iWidth      As Integer
  Dim sText    As String
  Dim iLenCaption As Integer
  
  sText = mlstTabs("Caption", piTabIndex)
  If meTabStyle = eTabButtons Then
    sText = " " & sText & " "
  End If
  
  iLenCaption = Len(sText)
  If Not mfFixedTabWidth Then
    iWidth = Len(sText)
  Else
    iWidth = miMaxTabWidth
  End If
  If iLenCaption > iWidth Then
    sText = left$(sText, iWidth - 1) & ChrW$(&H2026&) '"..."
  Else
    sText = StrBlock(sText, " ", iWidth)
  End If
  If mlstTabs("CanClose", piTabIndex) Then
    sText = sText & " "
  End If
  
  GetTabText = sText
End Function

Private Function WrapTabZone(ByVal piTabIndex As Integer, ByVal psText As String) As String
  If miBaseZoneID <> 0 Then
    'Use the formatted caption for the zone tag, as we're going to use it to replace the zone text when hovering
    WrapTabZone = VTX_ZONE_BEGIN(miBaseZoneID + piTabIndex, VT_Purge(psText)) & psText & VTX_ZONE_END(miBaseZoneID + piTabIndex)
  End If
End Function

Private Function GetTabElement(ByVal piTabIndex As Integer) As String
  Dim sElement    As String
  
  sElement = GetTabText(piTabIndex)
  'warning, zone must be inside colors as we invert the zone (for buttons) on hover
  sElement = ColorizeTabText(piTabIndex, WrapTabZone(piTabIndex, sElement))
  If mlstTabs("CanClose", piTabIndex) Then
    sElement = sElement & RenderCloseButton(piTabIndex, False)
  End If
  
  GetTabElement = sElement
End Function

'This will store the tabstrip rendering in a cached array
'so we can give it back line by line.
Private Sub RenderTabStrip()
  Dim iLine         As Integer
  Dim iTab          As Integer
  Dim sTab          As String
  Dim sLine         As String
  Dim sCaption      As String
  Dim iTabWidth     As Integer
  Dim iUsedWidth    As Integer
  Dim iMaxUsedWidth As Integer
  Dim sChevron      As String
  Dim i             As Integer
  Dim iAdditionalWidth  As Integer
  Dim sLeft         As String
  Dim sRight        As String
  Dim iMaxCols      As Integer
  
  On Error GoTo RenderTabStrip_Err
  
  If Not moConsole Is Nothing Then
    iMaxCols = moConsole.MaxVisibleCols
    If moConsole.LineCount > moConsole.MaxVisibleRows Then
      iMaxCols = iMaxCols - 1
    End If
  End If
  
  For iLine = 1 To NUMLINES_TABSTRIP
    sLine = ""
    iUsedWidth = 0
    
    'Chevron left ?
    If miStartTab > 1 Then
      Select Case iLine
      Case 1
        sLine = sLine & " " 'will be replaced later by left chevron if necessary
      Case 2
        sLine = sLine & " "
      Case 3
        sLine = sLine & " "
      Case 4
        sLine = sLine & ChrW$(&H2500) '-
      End Select
      iUsedWidth = 1
    End If
    
    For iTab = miStartTab To mlstTabs.Count
      sTab = ""
      iTabWidth = GetTabInternalWidth(iTab)
      iUsedWidth = iUsedWidth + iTabWidth
      If (iTab = miStartTab) Or (iTab = mlstTabs.Count) Then
        iAdditionalWidth = 2
      Else
        iAdditionalWidth = 1
      End If
      iUsedWidth = iUsedWidth + iAdditionalWidth
      sCaption = GetTabElement(iTab)
      
      Select Case iLine
      Case 1
        If iTab = miActiveTab Then
          sTab = sTab & ChrW$(&H250C) & String$(iTabWidth, ChrW$(&H2500)) & ChrW$(&H2510) 'left top corner, -, right top corner
        Else
          If (iTab + 1) = miActiveTab Then
            sTab = sTab & Space$(iTabWidth + iAdditionalWidth - 1)
          Else
            sTab = sTab & Space$(iTabWidth + iAdditionalWidth)
          End If
        End If
      Case 2  'text for active tab
        If iTab = miActiveTab Then
          If iTab = 1 Then
            sTab = sTab & ChrW$(&H2502) '|
          Else
            sTab = sTab & ChrW$(&H2524) '-45°T
          End If
          sTab = sTab & sCaption  '-45°T, spaces
          If iTab = mlstTabs.Count Then
            sTab = sTab & ChrW$(&H2502) '|
          End If
        Else
          If (miActiveTab <> 0) And (iTab = (miActiveTab + 1)) Then
            sTab = sTab & ChrW$(&H251C) '-45°T
          Else
            If iTab > 1 Then
              sTab = sTab & ChrW$(&H252C) 'T
            Else
              sTab = sTab & ChrW$(&H250C)  'left top corner
            End If
          End If
          sTab = sTab & String$(iTabWidth, ChrW$(&H2500)) '-
          If iTab = mlstTabs.Count Then
            sTab = sTab & ChrW$(&H2510)  'right top corner
          End If
        End If
      Case 3  'text for non active tabs
        If iTab = miActiveTab Then
          sTab = sTab & ChrW$(&H2502)
          If Not IsNull(mvActiveTabBackColor) Then
            sTab = sTab & VT_BCOLOR(mvActiveTabBackColor)
          End If
          sTab = sTab & ColorizeTabText(iTab, Space$(iTabWidth))
          If iTab = mlstTabs.Count Then
            sTab = sTab & ChrW$(&H2502) '|
          End If
          If Not IsNull(mvActiveTabBackColor) Then
            sTab = sTab & VT_BCOLOR(Me.BackColor)
          End If
        Else
          sTab = sTab & ChrW$(&H2502) & sCaption
          If iTab = mlstTabs.Count Then
            sTab = sTab & ChrW$(&H2502) '|
          End If
        End If
      Case 4
        sTab = sTab & ChrW$(&H2534) & String$(iTabWidth, ChrW$(&H2500)) ' 180°"T"
        If iTab = mlstTabs.Count Then
          sTab = sTab & ChrW$(&H2534)
        End If
      End Select
      
      sLine = sLine & sTab
    Next iTab
    masRender(iLine) = sLine
    If iUsedWidth > iMaxUsedWidth Then
      iMaxUsedWidth = iUsedWidth
    End If
  Next iLine
  
  If Not moConsole Is Nothing Then
    Dim iRemainingChars As Integer
    If iMaxCols > 0 Then
      iRemainingChars = iMaxCols - iMaxUsedWidth
      If iMaxUsedWidth <= iMaxCols Then
        sLine = String$(iRemainingChars + 1, ChrW$(&H2500)) '-
        masRender(4) = masRender(4) & sLine
      Else
        'put the right arrow
        sChevron = VTX_ZONE_BEGIN(ZONEID_RIGHTCHEVRON) & ChrW$(&H25BA) & VTX_ZONE_END(ZONEID_RIGHTCHEVRON) '&HAB=<< (>> = &HBB)
        sLeft = left$(masRender(1), iMaxCols - 3)
        sRight = Mid$(masRender(1), iMaxCols - 2)
        sLine = sLeft & sChevron & sRight
        masRender(1) = sLine
      End If
    End If
    If miStartTab > 1 Then
      sRight = Right$(masRender(1), Len(masRender(1)) - 1)
      masRender(1) = VTX_ZONE_BEGIN(ZONEID_LEFTCHEVRON) & ChrW$(&H25C4) & VTX_ZONE_END(ZONEID_LEFTCHEVRON) & sRight
    End If
    For i = 1 To NUMLINES_TABSTRIP
      masRender(i) = VT_FCOLOR(Me.ForeColor) & VT_BCOLOR(Me.BackColor) & masRender(i)
    Next i
  End If

RenderTabStrip_Output:
  If Not moConsole Is Nothing Then
    If miLine > 0 Then  'the strip has been RenderTabStriped once
      For i = 1 To NUMLINES_TABSTRIP
        If moConsole.LineCount < miLine + i - 1 Then
          Do
            moConsole.OutputLn ""
          Loop Until moConsole.LineCount = (miLine + i - 1)
        End If
        moConsole.SetLine miLine + i - 1, masRender(i), False, True
      Next i
      moConsole.RefreshWindow
    Else
      For i = 1 To NUMLINES_TABSTRIP
        iLine = moConsole.OutputLn(masRender(i))
        If i = 1 Then
          miLine = i
        End If
      Next i
    End If
  End If

RenderTabStrip_Exit:
  Exit Sub
RenderTabStrip_Err:
  Resume RenderTabStrip_Exit
  Resume
End Sub

Public Function TotalLinesRequiredForRender() As Integer
  If meTabStyle = eTabStrip Then
    TotalLinesRequiredForRender = NUMLINES_TABSTRIP
  Else
    TotalLinesRequiredForRender = NUMLINES_BUTTONS
  End If
End Function

Private Sub RenderButtonBar()
  Dim iLine         As Integer
  Dim iTab          As Integer
  Dim sTab          As String
  Dim sLine         As String
  Dim sCaption      As String
  Dim iTabWidth     As Integer
  Dim iUsedWidth    As Integer
  Dim iMaxUsedWidth As Integer
  Dim sChevron      As String
  Dim i             As Integer
  Dim sLeft         As String
  Dim sRight        As String
  Dim iMaxCols      As Integer
  
  On Error GoTo RenderButtonBar_Err
  
  If Not moConsole Is Nothing Then
    iMaxCols = moConsole.MaxVisibleCols
    If moConsole.LineCount > moConsole.MaxVisibleRows Then
      iMaxCols = iMaxCols - 1
    End If
  End If
  
  'NUMLINES_BUTTONS = 1
  sLine = ""
  iUsedWidth = 0
    
  'Chevron left ?
  If miStartTab > 1 Then
    sLine = sLine & " "
    iUsedWidth = 1
  End If
    
  For iTab = miStartTab To mlstTabs.Count
    sTab = ""
    iTabWidth = GetTabInternalWidth(iTab) + 2 'wrapped with two spaces
    iUsedWidth = iUsedWidth + iTabWidth
    
    sCaption = GetTabElement(iTab)
    sTab = sCaption
    sTab = ColorizeTabText(iTab, sTab)
    
    sLine = sLine & sTab
    If iTab < mlstTabs.Count Then
      sLine = sLine & msButtonSep
      iUsedWidth = iUsedWidth + Len(msButtonSep)
    End If
  Next iTab
    
  masRender(1) = sLine
  If iUsedWidth > iMaxUsedWidth Then
    iMaxUsedWidth = iUsedWidth
  End If

  If Not moConsole Is Nothing Then
    Dim iRemainingChars As Integer
    If iMaxCols > 0 Then
      iRemainingChars = iMaxCols - iMaxUsedWidth
      If iMaxUsedWidth > iMaxCols Then
        'put the right arrow
        sChevron = VTX_ZONE_BEGIN(ZONEID_RIGHTCHEVRON) & ChrW$(&H25BA) & VTX_ZONE_END(ZONEID_RIGHTCHEVRON) '&HAB=<< (>> = &HBB)
        sLeft = left$(masRender(1), iMaxCols - 3)
        sRight = Mid$(masRender(1), iMaxCols - 2)
        sLine = sLeft & sChevron & sRight
        masRender(1) = sLine
      End If
    End If
    If miStartTab > 1 Then
      sRight = Right$(masRender(1), Len(masRender(1)) - 1)
      masRender(1) = VTX_ZONE_BEGIN(ZONEID_LEFTCHEVRON) & ChrW$(&H25C4) & VTX_ZONE_END(ZONEID_LEFTCHEVRON) & sRight
    End If
  End If

  masRender(1) = VTX_SPILL(True) & masRender(1) & VT_BCOLOR(Me.BackColor) & " " & VTX_SPILL(False)
  
RenderTabStrip_Output:
  If Not moConsole Is Nothing Then
    If miLine > 0 Then  'the strip has been Rendered once
      For i = 1 To NUMLINES_BUTTONS
        If moConsole.LineCount < miLine + i - 1 Then
          Do
            moConsole.OutputLn ""
          Loop Until moConsole.LineCount = (miLine + i - 1)
        End If
        moConsole.SetLine miLine + i - 1, masRender(i), False, True
      Next i
      moConsole.RefreshWindow
    Else
      For i = 1 To NUMLINES_BUTTONS
        iLine = moConsole.OutputLn(masRender(i))
        If i = 1 Then
          miLine = i
        End If
      Next i
    End If
  End If

RenderButtonBar_Exit:
  Exit Sub
RenderButtonBar_Err:
  Resume RenderButtonBar_Exit
  Resume
End Sub

Public Sub Render()
  ClearRender
  If mlstTabs.Count > 0 Then
    If meTabStyle = eTabStrip Then
      RenderTabStrip
    Else
      RenderButtonBar
    End If
  End If
End Sub

Public Function IsMouseHit( _
  ByVal piRow As Integer, _
  ByVal piCol As Integer) As Boolean
  IsMouseHit = CBool((piRow >= Me.Line) And _
                      (piRow < (Me.Line + NUMLINES_TABSTRIP) And _
                      (piCol >= 1) And (piCol <= Len(masRender(3)))))
End Function

Private Sub OnZoneClick(ByVal piEvtCode As Integer, _
                        ByVal piMouseFlags As Integer, _
                        ByVal piZoneID As Integer, _
                        ByVal piRow As Integer, _
                        ByVal piCol As Integer, _
                        ByVal piPosX As Integer, _
                        ByVal piPosY As Integer)
  On Error Resume Next
  Dim iTabIndex     As Integer
  Dim sZoneTag      As String
  
  'don't forget that we have a base line at Me.Line
  piRow = piRow - miLine + 1 'Now piRow starts at 1
  'The tag of the zone is the full tab caption
  sZoneTag = moConsole.GetZoneTag(piRow, piZoneID)
  
  'Determine the tab that was clicked
  iTabIndex = piZoneID - miBaseZoneID
  
  RaiseEvent TabClick(iTabIndex, piMouseFlags, piZoneID, sZoneTag)
End Sub

Private Sub OnTabClose(ByVal piTabIndex As Integer)
  Dim fRetCancel  As Boolean
  Dim iZoneID     As Integer
  Dim sTabName    As String
  
  On Error Resume Next
  
  RaiseEvent TabClose(piTabIndex, fRetCancel)
  If Not fRetCancel Then
    sTabName = mlstTabs("Name", piTabIndex)
    mlstTabs.Remove piTabIndex
    If piTabIndex > mlstTabs.Count Then
      miActiveTab = mlstTabs.Count
    Else
      miActiveTab = piTabIndex
    End If
    Me.Render
    RaiseEvent TabRemoved(piTabIndex, sTabName)
    If miActiveTab > 0 Then
      iZoneID = miBaseZoneID + miActiveTab
      RaiseEvent TabClick(miActiveTab, WM_LBUTTONUP, iZoneID, moConsole.GetZoneTag(2, iZoneID))
    End If
  End If
End Sub

Public Function OnMouseButton( _
  ByVal phWnd As Long, _
  ByVal piEvtCode As Integer, _
  ByVal pwParam As Integer, _
  ByVal piZoneID As Integer, _
  ByVal piRow As Integer, _
  ByVal piCol As Integer, _
  ByVal piPosX As Integer, _
  ByVal piPosY As Integer, _
  ByVal plUserData As Long) As Integer
  
  Dim sText   As String
  Dim sTag    As String
  Dim iTabIndex As Integer
  
  On Error Resume Next
  
  Select Case piEvtCode
  Case WM_USER_ZONEENTER
    'Debug.Print "ENTER Zone "; piZoneID; "enter, tag="; sTag; ", text="; sText
    If piZoneID = ZONEID_LEFTCHEVRON Then
      Exit Function
    ElseIf piZoneID = ZONEID_RIGHTCHEVRON Then
      Exit Function
    ElseIf piZoneID <= (miBaseZoneID + ZONEID_CLOSEBOXES_OFFSET) Then
      iTabIndex = piZoneID - miBaseZoneID
      If meTabStyle = eTabStrip Then
        sText = VT_UNDL_ON() & GetTabText(iTabIndex) & VT_UNDL_OFF()
      Else
        sText = VT_INV_ON() & GetTabText(iTabIndex) & VT_INV_OFF()
      End If
    Else
      iTabIndex = piZoneID - (miBaseZoneID + ZONEID_CLOSEBOXES_OFFSET)
      sText = RenderCloseButtonText(iTabIndex, True)
    End If
    moConsole.ReplaceZone piRow, piZoneID, sText
    moConsole.RefreshWindow
    
  Case WM_USER_ZONELEAVE
    'Debug.Print "LEAVE Zone "; piZoneID; " leave, Text = "; sText
    If piZoneID = ZONEID_LEFTCHEVRON Then
      Exit Function
    ElseIf piZoneID = ZONEID_RIGHTCHEVRON Then
      Exit Function
    ElseIf piZoneID <= (miBaseZoneID + ZONEID_CLOSEBOXES_OFFSET) Then
      iTabIndex = piZoneID - miBaseZoneID
      sText = GetTabText(iTabIndex)
    Else
      iTabIndex = piZoneID - (miBaseZoneID + ZONEID_CLOSEBOXES_OFFSET)
      sText = RenderCloseButtonText(iTabIndex, False)
    End If
    moConsole.ReplaceZone piRow, piZoneID, sText
    moConsole.RefreshWindow
  
  'A click happens on (left) mouse button up, if it's on the
  'same click target (a calendar cell or a zone) the button down was.
  Case eWmMouseButton.WM_LBUTTONDOWN, eWmMouseButton.WM_RBUTTONDOWN
    'Debug.Print "ButtonDown(row="; piRow; ", col="; piCol; ", zone="; piZoneID; ")"
    If (piRow < miLine) Or (piRow >= (miLine + NUMLINES_TABSTRIP)) Then Exit Function
    
    miMBdownZoneID = piZoneID
    miMBdownRow = piRow
    miMBdownCol = piCol
    
  Case eWmMouseButton.WM_LBUTTONUP, eWmMouseButton.WM_RBUTTONUP
    'Debug.Print "ButtonUp(row="; piRow; ", col="; piCol; ", zone="; piZoneID; ")"
    If (piRow < miLine) Or (piRow >= (miLine + NUMLINES_TABSTRIP)) Then Exit Function 'should be redundant if caller checks hit (IsMouseHit())
    
    If (piRow = miMBdownRow) And (piCol = miMBdownCol) Then
      If piZoneID > 0 Then
        If piZoneID = miMBdownZoneID Then
          'Debug.Print "OnZoneClick(piEvtCode=" & piEvtCode & ", pwParam=" & pwParam & ", piZoneID=" & piZoneID & ", piRow=" & piRow & ", piCol=" & piCol & ", piPosX=" & piPosX & ", piPosY=" & piPosY & ")"
          If piZoneID = ZONEID_LEFTCHEVRON Then
            miStartTab = miStartTab - 1
            Render
          ElseIf piZoneID = ZONEID_RIGHTCHEVRON Then
            miStartTab = miStartTab + 1
            Render
          ElseIf piZoneID <= (miBaseZoneID + ZONEID_CLOSEBOXES_OFFSET) Then
            OnZoneClick piEvtCode, pwParam, piZoneID, piRow, piCol, piPosX, piPosY
          Else
            OnTabClose piZoneID - (miBaseZoneID + ZONEID_CLOSEBOXES_OFFSET)
          End If
          'prepare for next click detection
          miMBdownZoneID = 0
          miMBdownRow = 0
          miMBdownCol = 0
        End If
      Else
        'we could determine the clicked tabindex here and raise the TabClick event
        '/**/TODO
        
        'prepare for next click detection
        miMBdownZoneID = 0
        miMBdownRow = 0
        miMBdownCol = 0
      End If
    End If
  
  End Select
End Function

