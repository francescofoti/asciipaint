VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRibbon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private WithEvents mtabsRibbon As CCTabStrip
Attribute mtabsRibbon.VB_VarHelpID = -1
Private mconRibbon      As CConsoul
Private mlstRibControls As CList
Private moForm          As Form
Private moCtlFocusTrap  As Control

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
  Set mtabsRibbon = New CCTabStrip
End Sub

Private Sub Class_Terminate()
  Set moCtlFocusTrap = Nothing
  If Not mtabsRibbon Is Nothing Then
    mtabsRibbon.Detach
  End If
  If Not mconRibbon Is Nothing Then
    ConsoulEventDispatcher.UnregisterEventSink mconRibbon.hWnd
  End If
  Set mtabsRibbon = Nothing
  Set mlstRibControls = Nothing
  Set mconRibbon = Nothing
End Sub

Public Property Get TabControl() As CCTabStrip
  Set TabControl = mtabsRibbon
End Property

Public Property Get TabConsole() As CConsoul
  Set TabConsole = mconRibbon
End Property

Public Property Get ControlsList() As CList
  Set ControlsList = mlstRibControls
End Property

Public Property Get Form() As Form
  Set Form = moForm
End Property

Public Sub OnMousWheel(ByVal Count As Long)
  If mconRibbon Is Nothing Then Exit Sub
  If mtabsRibbon Is Nothing Then Exit Sub
  If Count < 0 Then
    If mtabsRibbon.ActiveTab > 1 Then
      mtabsRibbon_TabClick mtabsRibbon.ActiveTab - 1, eWmMouseButton.WM_LBUTTONUP, 0, 0
      'mtabsRibbon.ActiveTab = mtabsRibbon.ActiveTab - 1
      'mtabsRibbon.Render
    End If
  Else
    If mtabsRibbon.ActiveTab < mtabsRibbon.TabCount Then
      mtabsRibbon_TabClick mtabsRibbon.ActiveTab + 1, eWmMouseButton.WM_LBUTTONUP, 0, 0
      'mtabsRibbon.ActiveTab = mtabsRibbon.ActiveTab + 1
      'mtabsRibbon.Render
    End If
  End If
End Sub

Public Function Init( _
    ByRef poHostForm As Form, _
    ByVal plTabsForeColor As Long, _
    ByVal plTabsBackColor As Long, _
    ByVal plActiveTabBackColor As Long, _
    ByRef pIIConsoleEventSink As Object, _
    ByRef poCtlFocusTrap As Control _
  ) As Boolean
  Const LOCAL_ERR_CTX As String = "Init"
  On Error GoTo Init_Err
  Dim fOK As Boolean
  Dim hwndParent As Long
  ClearErr
  
  hwndParent = poHostForm.hWnd
  Set moForm = poHostForm
  Set moCtlFocusTrap = poCtlFocusTrap
  
  Set mconRibbon = New CConsoul
  mconRibbon.FontName = "Lucida Console"
  mconRibbon.FontSize = 10
  mconRibbon.MaxCapacity = 8
  mconRibbon.ForeColor = plTabsForeColor
  mconRibbon.BackColor = plTabsBackColor
  
  'Create the console window and tell the library that we want click feedback
  If Not mconRibbon.Attach( _
      hwndParent, 0, 0, 0, 0, _
      AddressOf MSupport.OnConsoulMouseButton, _
      piCreateAttributes:=LW_RENDERMODEBYLINE Or LW_TRACK_ZONES _
    ) Then
    SetErr LOCAL_ERR_CTX, -1&, "Failed to create ribbon window"
    GoTo Init_Exit
  End If
  mconRibbon.LineSpacing(elsTop) = 8
  mconRibbon.LineSpacing(elsBottom) = 8
  mconRibbon.LinePadding(elsTop) = -2
  mconRibbon.LinePadding(elsBottom) = -4
  mconRibbon.ShowWindow True
  mtabsRibbon.TabStyle = eTabButtons
  mtabsRibbon.BaseZoneID = 100
  mtabsRibbon.Attach mconRibbon
  mtabsRibbon.ActiveTab = 1
  mtabsRibbon.ActiveTabBackColor = plActiveTabBackColor
  mtabsRibbon.Render
  mconRibbon.ScrollTop
  mconRibbon.ShowScrollBar False
  ConsoulEventDispatcher.RegisterEventSink mconRibbon.hWnd, pIIConsoleEventSink, eCsMouseEvent
  
  Init = True
  
Init_Exit:
  Exit Function

Init_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume Init_Exit
  Resume
End Function

Public Sub PositionBandControls(ByVal piTop As Integer)
  Dim i       As Integer
  Dim iBand   As Integer
  Dim vCtl    As Variant
  Dim iRefBandTop As Integer
  Dim iDelta  As Integer
  Dim sTag    As String
  Dim fPreserveHeight As Boolean
  Dim iInitialHeight  As Integer
  Dim lstCtls   As CList
  Dim oNewList  As CList
  Dim lstBandCtls As CList
  
  Const TAGPARAM_RIBBONINDEX    As String = "rbindex"
  Const TAGPARAM_PRESERVEHEIGHT As String = "rbpsvheight"
  
  Set mlstRibControls = New CList
  mlstRibControls.ArrayDefine Array("itab", "controls", "idelta"), Array(vbInteger, vbObject, vbInteger)
  For i = 1 To mtabsRibbon.TabCount
    iRefBandTop = moForm.Controls("rectRibRef" & i).Top
    iDelta = iRefBandTop - piTop  'assumed that the controls are *below*
    Set oNewList = New CList
    oNewList.ArrayDefine Array("name"), Array(vbString)
    mlstRibControls.AddValues i, oNewList, iDelta
    Set oNewList = Nothing
  Next i
  
  On Error Resume Next
  Set lstCtls = New CList
  lstCtls.ArrayDefine Array("control", "initialheight"), Array(vbObject, vbInteger)
  For i = 0 To moForm.Controls.Count - 1
    Set vCtl = moForm.Controls(i)
    'If vCtl.Name = "cmdSaveFileAs" Then Stop
    sTag = GetTagParam(vCtl.Tag, TAGPARAM_RIBBONINDEX)
    iBand = Val(sTag)
      
    If iBand > 0 Then
      Set lstBandCtls = mlstRibControls("controls", iBand)
      lstBandCtls.AddValues vCtl.Name
      Set lstBandCtls = Nothing
      
      If iBand <> 1 Then
        fPreserveHeight = CBool(Len(GetTagParam(vCtl.Tag, TAGPARAM_PRESERVEHEIGHT)) > 0)
        If fPreserveHeight Then
          iInitialHeight = vCtl.Height
          lstCtls.AddValues vCtl, iInitialHeight
        End If
        iDelta = mlstRibControls("idelta", iBand)
        vCtl.Top = vCtl.Top - iDelta
      End If
    End If
    
    Set vCtl = Nothing
  Next i
  
  If lstCtls.Count > 0 Then
    For i = 1 To lstCtls.Count
      lstCtls("control", i).Height = lstCtls("initialheight", i)
    Next i
  End If
  
  Set lstCtls = Nothing
  Set oNewList = Nothing
End Sub

Private Sub ShowRibbonBand(ByVal piBand As Integer, ByVal pfShow As Boolean)
  Dim i   As Integer
  Dim lstCtls As CList
  Dim sCtlName As String
  
  On Error Resume Next
  If Not moCtlFocusTrap Is Nothing Then
    moCtlFocusTrap.Visible = True
    moCtlFocusTrap.SetFocus
  End If
  Set lstCtls = mlstRibControls("controls", piBand)
  For i = 1 To lstCtls.Count
    sCtlName = lstCtls("name", i)
    moForm.Controls(sCtlName).Visible = pfShow
  Next i
  If Not moCtlFocusTrap Is Nothing Then
    moCtlFocusTrap.Visible = False
  End If
End Sub

Private Sub mtabsRibbon_TabClick(ByVal piTabIndex As Integer, ByVal piMouseFlags As Integer, ByVal piZoneID As Integer, ByVal psZoneTag As String)
  Static iPrevTab As Integer
  On Error Resume Next
  If piTabIndex <> iPrevTab Then
    If iPrevTab = 0 Then
      iPrevTab = 1
    End If
    ShowRibbonBand iPrevTab, False
  End If
  mtabsRibbon.ActiveTab = piTabIndex
  mtabsRibbon.Render
  ShowRibbonBand piTabIndex, True
  'Me.cmdDummy.Visible = False
  iPrevTab = piTabIndex
End Sub

Public Sub OnResize()
  mtabsRibbon.Render
  mconRibbon.ScrollTop
  mconRibbon.ShowScrollBar False
End Sub

Public Function OnTabsMouseButton( _
  ByVal phWnd As LongPtr, _
  ByVal piEvtCode As Integer, _
  ByVal pwParam As Integer, _
  ByVal piZoneID As Integer, _
  ByVal piRow As Integer, _
  ByVal piCol As Integer, _
  ByVal piPosX As Integer, _
  ByVal piPosY As Integer) As Integer
  OnTabsMouseButton = mtabsRibbon.OnMouseButton(phWnd, piEvtCode, pwParam, piZoneID, piRow, piCol, piPosX, piPosY, 0)
End Function


